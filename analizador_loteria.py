#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Analizador de Series de Lotería - 3ª y 4ª Cifra
================================================
Aplica las 4 fórmulas de análisis a cualquier Excel de resultados,
detecta qué fórmula coincidió en cada sorteo para ambas cifras, y predice las siguientes.

Estructura Excel esperada (fila más reciente arriba):
  Fecha | 1ª cifra (B) | 2ª cifra (C) | 3ª cifra (D) | 4ª cifra (E)
Tolera filas vacías (recortes). Funciona con cualquier lotería.
"""

import pandas as pd
import json
import sys


# ============================================================
# DEFINICIÓN DE LAS 4 FÓRMULAS (genéricas para cualquier cifra)
# ============================================================
# Cada fórmula predice una cifra objetivo usando datos de días anteriores
# (filas de abajo = n+1, n+2). Resultado = última cifra (módulo 10).
#
# Para 4ª CIFRA (columna E):
#   F1: Horizontal SUMA - usa cifras del día anterior (n+1)
#       izq = B(n+1)+C(n+1)   |   der = D(n+1)+E(n+1)
#   F2: Horizontal MULTIPLICACIÓN - usa día anterior (n+1)
#       izq = B(n+1)×C(n+1)   |   der = D(n+1)×E(n+1)
#   F3: Vertical E - usa 4ª cifra de n+1 y n+2
#       suma = E(n+1)+E(n+2)  |   mult = E(n+1)×E(n+2)
#   F4: Vertical B - usa 1ª cifra de n+1 y n+2
#       suma = B(n+1)+B(n+2)  |   mult = B(n+1)×B(n+2)
#
# Para 3ª CIFRA (columna D):
#   F1: Horizontal SUMA - usa cifras del día anterior (n+1)
#       izq = B(n+1)+C(n+1)   |   der = D(n+1)+E(n+1)
#   F2: Horizontal MULTIPLICACIÓN - usa día anterior (n+1)
#       izq = B(n+1)×C(n+1)   |   der = D(n+1)×E(n+1)
#   F3: Vertical D - usa 3ª cifra de n+1 y n+2
#       suma = D(n+1)+D(n+2)  |   mult = D(n+1)×D(n+2)
#   F4: Vertical B - usa 1ª cifra de n+1 y n+2
#       suma = B(n+1)+B(n+2)  |   mult = B(n+1)×B(n+2)

FORMULAS = {
    "F1-izq": {"nombre": "Fórmula 1 (Horizontal Suma Izq)", "tipo": "horizontal", "op": "suma"},
    "F1-der": {"nombre": "Fórmula 1 (Horizontal Suma Der)", "tipo": "horizontal", "op": "suma"},
    "F2-izq": {"nombre": "Fórmula 2 (Horizontal Mult Izq)", "tipo": "horizontal", "op": "mult"},
    "F2-der": {"nombre": "Fórmula 2 (Horizontal Mult Der)", "tipo": "horizontal", "op": "mult"},
    "F3-suma": {"nombre": "Fórmula 3 (Vertical Suma)", "tipo": "vertical", "op": "suma"},
    "F3-mult": {"nombre": "Fórmula 3 (Vertical Mult)", "tipo": "vertical", "op": "mult"},
    "F4-suma": {"nombre": "Fórmula 4 (Vertical B Suma)", "tipo": "vertical", "op": "suma"},
    "F4-mult": {"nombre": "Fórmula 4 (Vertical B Mult)", "tipo": "vertical", "op": "mult"},
}


def calcular_candidatos(row_n1, row_n2, col_objetivo="E"):
    """
    Calcula los 8 candidatos para la cifra objetivo del día n,
    usando el día anterior (n+1) y dos días atrás (n+2).

    col_objetivo: "E" (4ª cifra) o "D" (3ª cifra)
    row = dict con claves B, C, D, E (enteros).
    Devuelve dict {clave_formula: (valor_predicho, texto_calculo)}
    """
    B1, C1, D1, E1 = row_n1["B"], row_n1["C"], row_n1["D"], row_n1["E"]
    out = {}

    # F1 - Horizontal suma (solo n+1)
    out["F1-izq"] = ((B1 + C1) % 10, f"{B1}+{C1}={(B1+C1)%10}")
    out["F1-der"] = ((D1 + E1) % 10, f"{D1}+{E1}={(D1+E1)%10}")

    # F2 - Horizontal mult (solo n+1)
    out["F2-izq"] = ((B1 * C1) % 10, f"{B1}×{C1}={(B1*C1)%10}")
    out["F2-der"] = ((D1 * E1) % 10, f"{D1}×{E1}={(D1*E1)%10}")

    # F3 y F4 requieren n+2
    if row_n2 is not None:
        B2 = row_n2["B"]
        # F3: Vertical usando la columna objetivo
        col_n1 = row_n1[col_objetivo]
        col_n2 = row_n2[col_objetivo]
        out["F3-suma"] = ((col_n1 + col_n2) % 10, f"{col_n1}+{col_n2}={(col_n1+col_n2)%10}")
        out["F3-mult"] = ((col_n1 * col_n2) % 10, f"{col_n1}×{col_n2}={(col_n1*col_n2)%10}")

        # F4: Vertical B
        out["F4-suma"] = ((B1 + B2) % 10, f"{B1}+{B2}={(B1+B2)%10}")
        out["F4-mult"] = ((B1 * B2) % 10, f"{B1}×{B2}={(B1*B2)%10}")
    else:
        for k in ["F3-suma", "F3-mult", "F4-suma", "F4-mult"]:
            out[k] = (None, "-")

    return out


def cargar_datos(path, hoja=0):
    """Carga el Excel, elimina filas vacías, normaliza columnas."""
    df = pd.read_excel(path, sheet_name=hoja)
    cols = df.columns.tolist()
    rename_map = {}
    nombres_std = ["Fecha", "B", "C", "D", "E"]
    for i, c in enumerate(cols[:5]):
        rename_map[c] = nombres_std[i]
    df = df.rename(columns=rename_map)
    df = df[["Fecha", "B", "C", "D", "E"]]
    df = df.dropna(subset=["B", "C", "D", "E"]).reset_index(drop=True)
    for c in ["B", "C", "D", "E"]:
        df[c] = df[c].astype(int)
    df["Fecha"] = df["Fecha"].astype(str)
    return df


def analizar(df, col_objetivo="E"):
    """
    Para cada fila (sorteo), calcula candidatos y marca qué fórmulas
    coincidieron con la cifra real. Devuelve lista de registros histórico.
    """
    registros = []
    n = len(df)
    for i in range(n):
        fecha = df.iloc[i]["Fecha"]
        cifra_real = int(df.iloc[i][col_objetivo])

        row_n1 = None
        row_n2 = None
        if i + 1 < n:
            r = df.iloc[i + 1]
            row_n1 = {"B": int(r["B"]), "C": int(r["C"]), "D": int(r["D"]), "E": int(r["E"])}
        if i + 2 < n:
            r = df.iloc[i + 2]
            row_n2 = {"B": int(r["B"]), "C": int(r["C"]), "D": int(r["D"]), "E": int(r["E"])}

        reg = {
            "fecha": fecha,
            "cifras": [int(df.iloc[i]["B"]), int(df.iloc[i]["C"]),
                       int(df.iloc[i]["D"]), int(df.iloc[i]["E"])],
            "cifra_real": cifra_real,
            "col_objetivo": col_objetivo,
            "candidatos": {},
            "aciertos": [],
        }

        if row_n1 is not None:
            cand = calcular_candidatos(row_n1, row_n2, col_objetivo)
            for k, (val, txt) in cand.items():
                reg["candidatos"][k] = {"valor": val, "calculo": txt,
                                        "acierto": (val == cifra_real)}
                if val == cifra_real:
                    reg["aciertos"].append(k)
        registros.append(reg)
    return registros


def estadisticas(registros):
    """Tasa de acierto por fórmula sobre todo el histórico."""
    stats = {k: {"aciertos": 0, "evaluados": 0} for k in FORMULAS}
    for reg in registros:
        for k in FORMULAS:
            if k in reg["candidatos"] and reg["candidatos"][k]["valor"] is not None:
                stats[k]["evaluados"] += 1
                if reg["candidatos"][k]["acierto"]:
                    stats[k]["aciertos"] += 1
    for k in stats:
        ev = stats[k]["evaluados"]
        stats[k]["tasa"] = round(100 * stats[k]["aciertos"] / ev, 1) if ev else 0.0
        stats[k]["nombre"] = FORMULAS[k]["nombre"]
    return stats


def analizar_recientes(registros, ventana=10):
    """
    Mira los últimos `ventana` sorteos (los más recientes = primeras filas).
    Cuenta hits por fórmula y detecta rachas / fórmula 'caliente'.
    """
    recientes = registros[:ventana]
    conteo = {k: 0 for k in FORMULAS}
    secuencia = []
    for reg in recientes:
        secuencia.append({"fecha": reg["fecha"], "cifra_real": reg["cifra_real"],
                          "aciertos": reg["aciertos"]})
        for k in reg["aciertos"]:
            conteo[k] += 1

    caliente = max(conteo, key=conteo.get) if any(conteo.values()) else None

    ultima_vez = {}
    for k in FORMULAS:
        ultima_vez[k] = None
        for idx, reg in enumerate(registros):
            if k in reg["aciertos"]:
                ultima_vez[k] = idx
                break

    return {"conteo": conteo, "secuencia": secuencia,
            "caliente": caliente, "ultima_vez": ultima_vez}


def predecir_siguiente(df, stats, reciente, col_objetivo="E"):
    """
    Predice la cifra del PRÓXIMO sorteo usando la última y penúltima fecha.
    """
    if len(df) < 2:
        return None

    r1 = df.iloc[0]
    r2 = df.iloc[1]
    row_n1 = {"B": int(r1["B"]), "C": int(r1["C"]), "D": int(r1["D"]), "E": int(r1["E"])}
    row_n2 = {"B": int(r2["B"]), "C": int(r2["C"]), "D": int(r2["D"]), "E": int(r2["E"])}

    cand = calcular_candidatos(row_n1, row_n2, col_objetivo)

    max_reciente = max(reciente["conteo"].values()) if any(reciente["conteo"].values()) else 1
    predicciones = []
    for k, (val, txt) in cand.items():
        if val is None:
            continue
        tasa = stats[k]["tasa"]
        mom = reciente["conteo"][k]
        mom_norm = (mom / max_reciente) * 100 if max_reciente else 0
        score = round(0.6 * tasa + 0.4 * mom_norm, 1)
        predicciones.append({
            "formula": k,
            "nombre": FORMULAS[k]["nombre"],
            "digito_predicho": val,
            "calculo": txt,
            "tasa_historica": tasa,
            "hits_recientes": mom,
            "score": score,
        })

    predicciones.sort(key=lambda x: x["score"], reverse=True)

    por_digito = {}
    for p in predicciones:
        d = p["digito_predicho"]
        if d not in por_digito:
            por_digito[d] = {"digito": d, "formulas": [], "score_total": 0}
        por_digito[d]["formulas"].append(p["formula"])
        por_digito[d]["score_total"] += p["score"]
    ranking_digitos = sorted(por_digito.values(),
                             key=lambda x: x["score_total"], reverse=True)

    return {
        "fecha_base": str(r1["Fecha"]),
        "cifras_base": [row_n1["B"], row_n1["C"], row_n1["D"], row_n1["E"]],
        "col_objetivo": col_objetivo,
        "predicciones": predicciones,
        "ranking_digitos": ranking_digitos,
    }


def generar_alertas(stats, reciente, prediccion, col_objetivo="E"):
    """Genera mensajes de alerta legibles."""
    cifra_label = "3ª" if col_objetivo == "D" else "4ª"
    alertas = []

    if reciente["caliente"] and reciente["conteo"][reciente["caliente"]] > 0:
        k = reciente["caliente"]
        alertas.append({
            "nivel": "info",
            "titulo": "🔥 Fórmula más activa recientemente",
            "mensaje": f"{FORMULAS[k]['nombre']} pegó {reciente['conteo'][k]} veces "
                       f"en los últimos {len(reciente['secuencia'])} sorteos."
        })

    if reciente["secuencia"] and reciente["secuencia"][0]["aciertos"]:
        ultimas = reciente["secuencia"][0]["aciertos"]
        nombres = ", ".join(FORMULAS[k]["nombre"] for k in ultimas)
        alertas.append({
            "nivel": "ok",
            "titulo": "✅ Último sorteo cumplió fórmula",
            "mensaje": f"El sorteo {reciente['secuencia'][0]['fecha']} "
                       f"coincidió con: {nombres}."
        })
    elif reciente["secuencia"]:
        alertas.append({
            "nivel": "warn",
            "titulo": "⚠️ Último sorteo sin coincidencia",
            "mensaje": f"El sorteo {reciente['secuencia'][0]['fecha']} "
                       f"no coincidió con ninguna fórmula."
        })

    if prediccion and prediccion["ranking_digitos"]:
        top = prediccion["ranking_digitos"][0]
        top_pred = prediccion["predicciones"][0]
        alertas.append({
            "nivel": "pred",
            "titulo": f"🎯 PRÓXIMO RESULTADO ({cifra_label} cifra)",
            "mensaje": f"Dígito más probable: {top['digito']} "
                       f"(apoyado por {len(top['formulas'])} fórmula(s)). "
                       f"Fórmula líder: {top_pred['nombre']} → {top_pred['digito_predicho']} "
                       f"(score {top_pred['score']})."
        })

    return alertas


def analisis_completo(path, ventana=10):
    """Analiza ambas cifras (3ª y 4ª) de forma completa."""
    df = cargar_datos(path)

    resultados_por_cifra = {}

    for col_objetivo, etiqueta in [("D", "3ª Cifra"), ("E", "4ª Cifra")]:
        registros = analizar(df, col_objetivo)
        stats = estadisticas(registros)
        reciente = analizar_recientes(registros, ventana)
        prediccion = predecir_siguiente(df, stats, reciente, col_objetivo)
        alertas = generar_alertas(stats, reciente, prediccion, col_objetivo)

        resultados_por_cifra[col_objetivo] = {
            "etiqueta": etiqueta,
            "total_sorteos": len(df),
            "fecha_mas_reciente": str(df.iloc[0]["Fecha"]),
            "estadisticas": stats,
            "reciente": reciente,
            "prediccion": prediccion,
            "alertas": alertas,
            "historico": registros,
        }

    return resultados_por_cifra


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else "datos/dorado_tarde_ejemplo.xlsx"
    resultados = analisis_completo(path)

    for col_obj, res in resultados.items():
        print("\n" + "=" * 70)
        print(f"ANÁLISIS DE SERIE - {res['etiqueta']}")
        print("=" * 70)
        print(f"Total sorteos analizados: {res['total_sorteos']}")
        print(f"Fecha más reciente: {res['fecha_mas_reciente']}")

        print(f"\n--- TASA DE ACIERTO POR FÓRMULA (histórico) ---")
        for k, s in sorted(res["estadisticas"].items(),
                           key=lambda x: x[1]["tasa"], reverse=True):
            print(f"  {s['nombre']:40s} {s['aciertos']:3d}/{s['evaluados']:3d} = {s['tasa']:5.1f}%")

        print(f"\n--- ACTIVIDAD RECIENTE (últimos {len(res['reciente']['secuencia'])} sorteos) ---")
        for s in res["reciente"]["secuencia"]:
            hits = ", ".join(s["aciertos"]) if s["aciertos"] else "sin coincidencia"
            print(f"  {s['fecha']}  {col_obj}={s['cifra_real']}  →  {hits}")

        print(f"\n--- ALERTAS ---")
        for a in res["alertas"]:
            print(f"  [{a['nivel'].upper()}] {a['titulo']}")
            print(f"       {a['mensaje']}")

        print(f"\n--- PREDICCIÓN PRÓXIMO SORTEO ---")
        pred = res["prediccion"]
        print(f"  Base (último sorteo): {pred['fecha_base']} = {pred['cifras_base']}")
        print(f"\n  Ranking de dígitos candidatos:")
        for rd in pred["ranking_digitos"]:
            fs = ", ".join(rd["formulas"])
            print(f"    Dígito {rd['digito']}  (score {rd['score_total']:.1f})  ←  {fs}")
        print(f"\n  Detalle por fórmula:")
        for p in pred["predicciones"]:
            print(f"    {p['nombre']:40s} → {p['digito_predicho']}  "
                  f"({p['calculo']})  tasa={p['tasa_historica']}%  score={p['score']}")

    # Guardar JSON para el dashboard
    with open("analisis_datos.json", "w", encoding="utf-8") as f:
        json.dump(resultados, f, ensure_ascii=False, indent=2)
    print(f"\n✅ Datos guardados en analisis_datos.json")

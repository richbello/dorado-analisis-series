# Análisis de Series — Dorado Tarde (3ª y 4ª cifra)

Herramienta para analizar resultados históricos del chance y calcular, con 4 fórmulas, los dígitos candidatos de la **3ª** y **4ª cifra** del próximo sorteo.

## Contenido

| Archivo | Qué es |
|---|---|
| `index.html` | Dashboard web. Se abre en el navegador, carga un Excel y muestra la predicción de la 3ª y 4ª cifra. |
| `analizador_loteria.py` | Mismo análisis en Python (consola + `analisis_datos.json`). |
| `datos/dorado_tarde_ejemplo.xlsx` | Ejemplo con 100 sorteos (12/06/2026 – 10/10/2026). |
| `docs/FORMULAS.md` | Explicación detallada de las fórmulas. |

## Estructura del Excel

| Fecha | 1ª cifra | 2ª cifra | 3ª cifra | 4ª cifra |
|---|---|---|---|---|
| 10/10/2026 | 7 | 1 | 0 | 2 |

La primera fila puede ser encabezado. Las filas vacías se ignoran y el orden de fechas se ajusta solo (más reciente arriba).

## Fórmulas

Para el sorteo `n` se usan el día anterior (`n+1`) y el de antes (`n+2`); siempre se toma la última cifra (mód 10).

| Fórmula | Cálculo |
|---|---|
| F1 Suma horizontal | `B+C` y `D+E` del día anterior |
| F2 Mult. horizontal | `B×C` y `D×E` del día anterior |
| F3 Vertical objetivo | `D(n+1)+D(n+2)` y `D(n+1)×D(n+2)` (para la 4ª cifra se usa E) |
| F4 Vertical B | `B(n+1)+B(n+2)` y `B(n+1)×B(n+2)` |

**Score** de cada fórmula = 60 % tasa de acierto histórica + 40 % aciertos en los últimos 10 sorteos. Los dígitos se ordenan sumando los scores de las fórmulas que los apoyan.

## Uso

**Dashboard:** abrir `index.html` en el navegador → **Cargar Excel**.
También funciona publicado con GitHub Pages.

**Python:**
```bash
pip install -r requirements.txt
python analizador_loteria.py datos/dorado_tarde_ejemplo.xlsx
```

## Nota

Con 100 sorteos, cada fórmula acierta entre 8 % y 14 %, cerca del 10 % que daría elegir un dígito al azar. El ranking sirve para ordenar candidatos, no es una probabilidad garantizada.

## Autor

Richard Bello Roncancio — [@richbello](https://github.com/richbello)

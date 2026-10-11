# Fórmulas de análisis — 3ª cifra (D) y 4ª cifra (E)

## Estructura del Excel

Fila más reciente arriba. Columnas: **A** Fecha · **B** 1ª cifra · **C** 2ª cifra · **D** 3ª cifra · **E** 4ª cifra.

Para calcular la cifra del sorteo `n`:
- `n+1` = día anterior (fila de abajo)
- `n+2` = dos días atrás (dos filas abajo)
- Siempre se toma la última cifra del resultado (módulo 10).

Las 4 fórmulas son iguales para la 3ª y la 4ª cifra; solo cambia la columna objetivo (D o E) en la Fórmula 3.

## Fórmula 1 — Suma horizontal
- Izquierda: `(B(n+1) + C(n+1)) mod 10`
- Derecha: `(D(n+1) + E(n+1)) mod 10`

## Fórmula 2 — Multiplicación horizontal
- Izquierda: `(B(n+1) × C(n+1)) mod 10`
- Derecha: `(D(n+1) × E(n+1)) mod 10`

## Fórmula 3 — Vertical sobre la columna objetivo
- 3ª cifra: `(D(n+1) + D(n+2)) mod 10` y `(D(n+1) × D(n+2)) mod 10`
- 4ª cifra: `(E(n+1) + E(n+2)) mod 10` y `(E(n+1) × E(n+2)) mod 10`

## Fórmula 4 — Vertical sobre la 1ª cifra (B)
- `(B(n+1) + B(n+2)) mod 10` y `(B(n+1) × B(n+2)) mod 10`

## Ejemplo (próximo sorteo después del 10/10/2026)

Base: 10/10/2026 = 7-1-0-2 · 09/10/2026 = 0-4-5-7

| Fórmula | 3ª cifra | 4ª cifra |
|---|---|---|
| F1-izq | 7+1 = 8 | 7+1 = 8 |
| F1-der | 0+2 = 2 | 0+2 = 2 |
| F2-izq | 7×1 = 7 | 7×1 = 7 |
| F2-der | 0×2 = 0 | 0×2 = 0 |
| F3-suma | 0+5 = 5 | 2+7 = 9 |
| F3-mult | 0×5 = 0 | 2×7 = 4 |
| F4-suma | 7+0 = 7 | 7+0 = 7 |
| F4-mult | 7×0 = 0 | 7×0 = 0 |

## Score y ranking

- Score de cada fórmula = 0.6 × tasa histórica (%) + 0.4 × momentum, donde momentum = aciertos en los últimos 10 sorteos ÷ máximo de aciertos de cualquier fórmula × 100.
- Cada dígito suma el score de las fórmulas que lo predicen; el de mayor total es el candidato principal.

## Tasas de acierto (100 sorteos, 12/06/2026 – 10/10/2026, 3ª cifra)

| Fórmula | Aciertos | Tasa |
|---|---|---|
| F1-izq | 14/99 | 14.1 % |
| F3-suma | 13/98 | 13.3 % |
| F4-suma | 11/98 | 11.2 % |
| F1-der | 11/99 | 11.1 % |
| F2-izq | 10/99 | 10.1 % |
| F4-mult | 9/98 | 9.2 % |
| F2-der | 9/99 | 9.1 % |
| F3-mult | 8/98 | 8.2 % |

Elegir un dígito al azar acierta ~10 %. Con esta cantidad de datos las diferencias entre fórmulas no son concluyentes; conviene volver a medir a medida que se agregan sorteos.

# Análisis del ETF QQQ — Ingeniería Económica 2026

Análisis del ETF **QQQ** (Invesco QQQ Trust, que replica el índice Nasdaq-100) en tres partes:

1. **Composición del fondo**: qué empresas lo forman, cuánto pesa cada una y en qué sectores.
2. **Pronóstico con 5 modelos predictivos** (ingenuo, regresión lineal, ARIMA, Random Forest y Gradient Boosting) a 1 semana, 3 meses y 12 meses, y comparación entre ellos.
3. **Cadena de suministro** y relación causa-efecto: por qué sube o baja la cotización.

## Cómo verlo

**1. Leer el informe (sin instalar nada):** [informe/Informe_Analisis_QQQ.docx](informe/Informe_Analisis_QQQ.docx)

**2. Ejecutarlo en vivo (sin instalar nada):**

[![Abrir en Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/Chaparroo/analisis-qqq-ingenieria-economica-/blob/master/analisis_QQQ.ipynb)

Clic en el botón y luego **Entorno de ejecución → Ejecutar todas**. La primera celda descarga los datos de este repositorio sola.

**3. Ejecutarlo en tu computador** (con Python instalado):

```bash
pip install -r requirements.txt
jupyter notebook analisis_QQQ.ipynb
```

> Los precios se descargan en vivo de Yahoo Finance, así que se necesita internet y las cifras se actualizan al día en que se ejecute. La sección 2.2 tarda unos minutos.

## Archivos

| Archivo | Qué es |
|---|---|
| `analisis_QQQ.ipynb` | **Notebook principal** con todo el análisis. |
| `informe/Informe_Analisis_QQQ.docx` | El análisis en Word. |
| `informe/cadena_suministro_qqq.png` | Diagrama de la cadena de suministro. |
| `informe/generar_informe_word.py` | Vuelve a crear el Word desde el notebook: `python informe/generar_informe_word.py`. |
| `datos/` | Composición oficial de Invesco, sectores de cada empresa y registro de pronósticos. |

**Fuentes:** composición oficial de [Invesco](https://www.invesco.com/qqq-etf) (al 23 de octubre de 2026) y precios de Yahoo Finance (librería `yfinance`).

> Trabajo académico. Los pronósticos son ejercicios estadísticos, no recomendaciones de inversión.

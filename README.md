# Investment Lab

Simulador de eventos discretos para comparar cuatro estrategias de inversión de un inversionista de Medellín con 700 millones de pesos y diez años de horizonte. El capital se reparte entre una cuenta remunerada, CDT a 12 meses, el ETF iCOLCAP y un apartamento para arrendar. La meta es terminar con al menos 1.000 millones en pesos de hoy sin quedarse sin liquidez en el camino.

Capital, meta, horizonte, liquidez mínima, pesos de las estrategias y cada parámetro de los modelos se pueden cambiar desde la aplicación. Ningún número del modelo está escrito en el código del motor: todo sale de la hoja `PARAMETROS` de `data/master_datos.xlsx`, con su fuente y su estado (Estimado, Supuesto, Configurable o Pendiente).

## Correrlo en local

```bash
pip install -r requirements.txt
streamlit run app/main.py
```

Para los tests y los scripts de estimación:

```bash
pip install -r requirements-dev.txt
pytest -q
python tools/correr_experimentos.py 12     # corrida completa del informe con 12 procesos, unos 9 minutos
```

## Qué hay en cada carpeta

| Carpeta | Contenido |
|---|---|
| `src/investment_sim/engine` | Motor: lista de eventos con `heapq`, estado del sistema y un manejador por cada evento del diagrama |
| `src/investment_sim/models` | Un módulo por proceso aleatorio: Markov, inflación, Junta y CDT, ETF, inmueble, vacancia, liquidez |
| `src/investment_sim/strategies` | Las cuatro estrategias y la regla de revisión de E4 |
| `src/investment_sim/experiments` | Réplicas con números aleatorios comunes, capital-meta, tornado y validación |
| `src/investment_sim/metrics` | Intervalos de confianza (t, Wilson, bootstrap), CVaR, diferencias pareadas |
| `app/` | Interfaz en Streamlit con nueve páginas |
| `data/master_datos.xlsx` | Parámetros, fuentes, diccionario y datos procesados |
| `data/raw/` | Series que alimentan las estimaciones (IPC, TPM, CDT semanal, PIB, ETF, IPVU, decisiones de la Junta) |
| `data/fuentes/etf/` | Descargas originales y series procesadas del análisis de ICOLCAP (BanRep, FRED, Yahoo, BlackRock, OCDE); `metadata_fuentes.csv` dice de dónde sale cada una |
| `data/fuentes/inmueble/` | Datos del análisis inmobiliario (avisos de arriendo, Lonja, predial, IPVU, tiempos hasta arrendar); `raw/fuentes_y_fechas_de_consulta.csv` tiene la fecha de cada consulta |
| `outputs/resultados/` | Resultados de la corrida del informe: 1.068 réplicas por estrategia y escenario |
| `tests/` | 59 pruebas: eventos, cálculos financieros, generadores, estrategias, reproducibilidad |
| `docs/` | Informe, registro de decisiones y diagramas del equipo |

## Reproducibilidad

La réplica `r` usa la semilla `semilla_base + r`, y de ella salen 13 generadores independientes (uno por proceso) con `numpy.random.SeedSequence.spawn`. Las cuatro estrategias de una misma réplica ven la misma economía. Con la semilla 20261107 y los parámetros de referencia, la corrida da exactamente los números de `outputs/resultados/resumen.csv`, con uno o con varios procesos.

## Publicarlo en Streamlit Community Cloud

1. Subir este repositorio a GitHub (puede ser privado).
2. En share.streamlit.io, *Create app*, elegir el repositorio, rama `main` y archivo `app/main.py`.
3. Python 3.12. Las dependencias se instalan solas desde `requirements.txt`.

En la nube conviene dejar el control de procesos en 1 o 2 y no pasar de unas 2.000 réplicas por ejecución, porque el plan gratuito tiene poca CPU y memoria.

## Límites conocidos

Están detallados en `docs/decision_log.md`. Los dos más visibles: la tasa real simulada (TPM menos inflación) promedia 3,4 puntos contra 1,2 en la historia, y la vacancia simulada (2,5 % del tiempo) es más alta que el 0,6 % que reporta Ciencuadras para Medellín. Las reparaciones extraordinarias están apagadas porque no conseguimos datos.

Herramienta académica. Los resultados describen el modelo y sus supuestos; no son una recomendación de inversión.

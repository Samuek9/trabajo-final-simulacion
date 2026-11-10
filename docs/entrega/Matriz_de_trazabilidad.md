# Matriz de trazabilidad (resultados verificados)

| Resultado en el informe | Fuente de verdad | Estado |
|---|---|---|
| P(meta) por estrategia y escenario (Tabla 57) | `outputs/resultados/resumen.csv`; Excel, hoja «Resultados, estrategias» | Verificado, 12 de 12 |
| Precisión de P(meta) según réplicas (Tabla 50) | `precision_replicas.csv`; Excel, «Comparación de estrategias» | Verificado |
| Réplicas por configuración (1.068) | `corrida.json`; Excel, «Resultados, estrategias» | Verificado |
| Capital-meta: capitales, metas, réplicas | `capital_meta.csv`, `capital_necesario.csv` | Verificado |
| Comparaciones pareadas y corrección de Holm | `comparaciones_pareadas.csv`; Excel, «Comparación de estrategias» | Verificado en el Excel (63 de 72 significativas); texto pendiente de cotejo línea a línea |
| Sensibilidad (20 factores, 8 modelos, interacción) | `tornado.csv`, `sensibilidad_modelos.csv`, `interaccion.csv` | Verificado en el Excel; texto pendiente |
| Validación frente a observado | `validacion.csv`; Excel, «Validación del simulador» | Verificado en el Excel; texto pendiente |
| r₀ de la valorización | PARAMETROS/MODELO `val.r0`; Excel «Valorización, modelo» | Verificado |
| Venta urgente (10 %, 4 a 7 meses) | `inm.descuento_urgente`, `inm.venta_min`, `inm.venta_max` | Verificado |
| Vacancia por estado (p y duración) | `vac.p_*`, `vac.media_*`; Excel «Vacancia, modelo» | Verificado |
| Generadores aleatorios (13) | `engine/simulation.py`, `FLUJOS` | Verificado |
| Evento de crisis forzada | `engine/events.py` | Verificado |
| Logit de la Junta, α de la cuenta, margen del CDT, ETF, inflación, Markov | Excel, hojas de cada proceso | Cifras del Excel verificadas contra su hoja de cálculo; texto del informe pendiente de cotejo línea a línea |

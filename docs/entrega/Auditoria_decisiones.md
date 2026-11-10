# Auditoría de decisiones sobre el Excel

Copia de la hoja AUDITORIA que estaba en `data/master_datos.xlsx`, retirada del libro porque es registro de trabajo y no documentación de los modelos.

| Auditoría y cambios del libro |  |  |  |  |  |  |  |
| Registro de los hallazgos de organización y de los cambios hechos. Las correcciones metodológicas se hicieron solo con autorización. |  |  |  |  |  |  |  |
| ID | Hoja | Ubicación | Hallazgo | Tipo | Gravedad | Acción | Estado |
| A01 | INF | Serie mensual y anual | El 6,29 % de 2026 es la inflación de septiembre, no de diciembre; el modelo lo usa solo como inflación inicial (P033). | Aclaración temporal | Media | Se conserva el dato; la hoja Inflación lo rotula como septiembre de 2026 | Resuelto |
| A02 | PARAMETROS | Hoja completa | Mezclaba valores de inversionista con valores estimados de los procesos aleatorios. | Estructural | Alta | PARAMETROS solo trae lo editable; los valores estimados pasaron a la hoja oculta MODELO con las mismas claves | Resuelto |
| A03 | PARAMETROS | Columnas Fuente y Soporte | Citaban secciones del informe y no aportaban a quien edita. | Visual | Baja | Se quitaron de la hoja visible; siguen en MODELO | Resuelto |
| A04 | MODELOS | Columnas de alternativa y parámetros | Mostraban más de un modelo por proceso. | Estructural | Media | Queda solo el modelo elegido y la prueba que lo respalda | Resuelto |
| A05 | PARAMETROS | P153 (necesidades de liquidez) | El monto mediano era 20 millones fijos, sin relación con el capital. | Metodológico | Alta | Pasa a ser 3 % del capital inicial (liq.mediana_pct); se corrieron de nuevo los experimentos | Resuelto, con autorización del usuario |
| A06 | PARAMETROS | Reparaciones (P160 a P162) | Sin datos de costo ni frecuencia. | Dato faltante | Media | Salen de la vista y del informe; quedan apagadas en MODELO | Resuelto |
| A07 | RESULTADOS | Capital-meta | La malla de capitales y metas estaba fija en el código. | Compatibilidad | Media | Se calcula como múltiplos del capital y la meta configurados; con la referencia da la misma malla | Resuelto |
| A08 | Todas | Hojas del equipo | Algunas notas mencionaban el enunciado o al equipo. | Visual | Baja | Se reescribieron en la copia del libro; los originales no se tocan | Resuelto |
| A09 | Todas | Datos originales y procesados | Conviven en la misma hoja con pruebas y texto, en bloques horizontales. | Estructural | Media | Inflación: las hojas Inflación e Inflación, modelo leen de INF con fórmulas, sin mover nada. Completo | Resuelto |
| A10 | VALIDACION | Autocorrelación de la inflación | La simulada (0,19) queda por debajo de la observada (0,42) porque se mide con 10 años por réplica. | Posible limitación estadística | Baja | Se documenta en el informe; no se cambia el modelo | Documentado |
| A11 | PIB | Nota de crisis y conteo desde 2011 | Decía que crisis es crecimiento negativo y contaba 4 trimestres de crisis desde 2011, pero los umbrales vigentes son 1 % y 4 % (12 trimestres de crisis desde 2006). | Inconsistencia | Media | La hoja PIB, estados recalcula el conteo con fórmula y describe la regla con los umbrales vigentes; la hoja original queda oculta sin cambios | Resuelto |
| A12 | INM | Costos y canon en pesos fijos | La hoja original calculaba predial, administración y pintura sobre un apartamento de 300 millones. | Estructural | Baja | La hoja Inmueble, datos expresa los costos como proporción del valor comercial; la original queda oculta sin cambios | Resuelto |

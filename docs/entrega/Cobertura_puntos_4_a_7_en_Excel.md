# Cobertura de los puntos 4 a 7 en el Excel

Cotejo del enunciado contra `data/master_datos.xlsx` (hojas visibles).

## Punto 4. Parámetros configurables
Cubierto en la hoja PARAMETROS: capital inicial, horizonte, meta en patrimonio real, liquidez mínima y otros parámetros (réplicas, semilla, plazo del CDT, necesidades de liquidez, pesos de las estrategias, valor del inmueble). Los experimentos que varían capital y meta los ejecuta el simulador y no forman parte del libro.

## Punto 5. Universo de inversión
Los cuatro grupos de activos están en las estrategias (PARAMETROS) y en las hojas de cada proceso. **Abierto:** la tabla de reglas financieras por activo (inversión mínima, comisiones, penalizaciones, tiempos de venta, costos) está en el informe (Tablas 2 a 6) y en la hoja oculta MODELO, pero no hay una tabla equivalente visible en el Excel.

## Punto 6. Fuentes de incertidumbre
Cubierto en MODELOS: diez procesos, con los cinco obligatorios (ETF, inflación, tasas, valorización y vacancia) más cuenta, margen del CDT, reajuste del canon, estado de la economía y necesidades de liquidez.

## Punto 7. Estudio y modelación (11 elementos)
Todas las hojas de modelo traen datos y fuente, evidencia estadística, especificación con parámetros estimados, validación o lectura, y dependencia. En esta revisión se agregó el elemento 11 (limitaciones) en las ocho hojas donde faltaba (Markov ya lo tenía) y, en Vacancia, la comparación de las cuatro distribuciones (log-verosimilitud, error máximo, mediana y P90), que estaba solo en el informe.

| Proceso | Candidatos y comparación | Atípicos | Limitaciones |
|---|---|---|---|
| Inflación | normal frente a lognormal (chi-cuadrado) y memoria entre años | sin tratamiento en el informe | agregadas |
| Junta (TPM) | prueba de razón de verosimilitud con y sin brecha de Taylor, y normalidad de los cambios | sin tratamiento en el informe | agregadas |
| Cuenta | proporción única (no hay candidatos) | sin tratamiento (cinco fechas) | agregadas |
| Margen del CDT | modelos candidatos con RMSE y distribución del error | 90 semanas fuera de Tukey, se conservan | agregadas |
| ETF | modelos candidatos con log L y KS | 3 meses fuera de Tukey | agregadas |
| Valorización | modelos candidatos con log L y AIC | ningún año pasa de 2,5 desviaciones | agregadas |
| Reajuste del canon | regresión frente a seguir todo el IPC | sin tratamiento en el informe | agregadas |
| Vacancia | cuatro distribuciones con curva, parámetros y tabla de ajuste | no aplica (curva agregada) | agregadas |
| Estado de la economía | cadena de Markov frente a estados independientes | no aplica | ya estaban |
| Necesidades de liquidez | sin datos de monto: supuesto documentado en MODELOS | no aplica | en MODELOS |

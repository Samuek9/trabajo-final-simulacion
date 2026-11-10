# Informe de cambios y verificación

Documento base: `docs/informe_final.docx` (73 páginas, 61 tablas, 16 imágenes). Las copias de entrega están en esta carpeta y el original no se modificó.

## Qué se hizo en esta revisión
- Se auditó el DOCX completo (títulos, tablas, referencias, índice) y se comparó con el código, el Excel y `outputs/resultados`.
- No se cambió ningún dato, parámetro, prueba ni resultado. Las verificaciones no encontraron discrepancias que justificaran cambiarlos.
- El texto no contiene las fórmulas de relleno señaladas (cero apariciones de «Cabe destacar», «Es importante», «En este sentido», «En conclusión», «Como se puede observar», «Por consiguiente», «De esta manera») ni rayas largas ni referencias a «enunciado», «equipo», «Hoja», `[L…]` o scratchpad.

## Cifras verificadas
- Tabla 57 (12 configuraciones): P(meta) coincide con `resumen.csv` en las 12 filas.
- Tabla 50 (precisión según réplicas): coincide con `precision_replicas.csv` y con la hoja «Comparación de estrategias» del Excel.
- 1.068 réplicas: consistente. La única mención de «1.000 réplicas» es el B del bootstrap para los p de KS y AD (sección 7.5), que es otra cosa.
- r₀ = 0,0809: coincide con `val.r0` (P138) y con el Excel.
- Venta urgente del inmueble, descuento de 10 % y 4 a 7 meses: coincide con P118 a P120.
- Probabilidad de no renovar 5 %, 10 % y 30 %, y duración media 1, 1,5 y 6 meses: coinciden con P140 a P146. La vacancia implícita de 0,41 % en expansión coincide con el Excel.
- 13 generadores de números aleatorios: coincide con `FLUJOS` en `engine/simulation.py`.
- «Crisis forzada»: es un evento real del código (`CRISIS_INICIO` y `CRISIS_FIN`, escenario de estrés); su mención en el informe es correcta.
- El diseño capital-meta (7 capitales de 400 a 1.000 M, 5 metas de 700 a 1.500 M, 500 réplicas, escenarios base y estrés) coincide con `capital_meta.csv` y con el Excel.

## Referencias
No se verificaron en línea los metadatos (DOI, URL, volúmenes) de las referencias; esa comprobación queda pendiente (ver Pendientes).

## Elementos sin poder comprobarse
Ver `Lista_de_pendientes.md`.

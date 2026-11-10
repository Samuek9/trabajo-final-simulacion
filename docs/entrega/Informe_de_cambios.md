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

## Cotejo automático de la sección 7
Se extrajeron las cifras con decimales de cada subsección 7.1 a 7.8 y se buscaron en la hoja de cálculo del proceso. Coincidieron todas las de 7.4 (55 de 55) y casi todas las de 7.1 (40 de 43), 7.2 (48 de 54) y 7.3 (9 de 11). Las que no coinciden son cifras calculadas fuera del Excel (parámetros de modelos candidatos del ETF, correlaciones con rezago, p de pruebas hechas en Python) o rótulos redondeados, y están en la lista de pendientes.

## Corrección en el Excel
El intervalo de α de la cuenta digital: el informe (7.3) usa el promedio de cada cuenta (0,755 a 1,046, t con 3 grados de libertad), correcto porque las 19 tasas vienen de cuatro cuentas; la hoja original del Excel traía el de las 19 tasas (0,829 a 0,967). La hoja «Cuenta, modelo» ahora calcula y muestra el intervalo por cuenta (fórmulas, reproduce 0,755 a 1,046) y deja el otro rotulado «19 tasas». El informe no cambió.

## Referencias
No se verificaron en línea los metadatos (DOI, URL, volúmenes) de las referencias; esa comprobación queda pendiente (ver Pendientes).

## Elementos sin poder comprobarse
Ver `Lista_de_pendientes.md`.


## Corrección de «NaN» en el capital necesario para 80 % (revisión visual del PDF)

Al revisar el PDF a tamaño completo aparecía «NaN millones» en cuatro lugares (sección 17, respuesta 4 de la sección 18 y las conclusiones). La causa: la malla de capital llega a 1.000 millones y con E2 la probabilidad de la meta de 1.000 millones es 75,8 % ahí, así que el capital para 80 % queda fuera de la malla y la interpolación devuelve NaN. No se inventó una cifra: el texto ahora dice que ni 1.000 millones bastan (76 %) y que no se extrapola. Cambio en `tools/construir_informe.py` (función `cap_fuera` y frases condicionales); el informe se regeneró (73 páginas) y no quedan «NaN». Los datos y resultados no cambiaron.


## Verificación de referencias y revisión visual

- DOI: los 13 DOI de revistas y libros se consultaron en Crossref y coinciden con autor, título, revista, volumen, número, páginas y año de la lista (el del libro de Burnham y Anderson resuelve en Springer). Para McCullagh (1980), Crossref da 109-127 y la lista 109-142, que es la paginación habitual del artículo con su discusión; no se cambió.
- URL: 21 de 22 responden 200. La del primer informe trimestral de Ciencuadras (2025) devuelve 404 desde este equipo, aunque un buscador la lista; hay que abrirla en un navegador antes de entregar.
- Revisión visual de las 73 páginas del PDF: sin tablas cortadas ni figuras fuera de margen. Detalles menores: la última página del anexo B tiene solo 7 filas y una URL larga de Fogafín se parte en la lista de referencias.


## Cuatro puntos sin ubicar y cifras de Python de la sección 7

- Inflación entre diciembres: confirmado en `models/inflation.py` (`interpolar`, línea recta entre diciembres); coincide con 7.1.5 y con la Tabla 45.
- Participación del arriendo: `resumen.csv` da 44,0 % (E2), 67,5 % (E3) y 45,6 % (E4); el texto dice 44 % y 67 %. Coincide.
- Costos totales: el código suma comisión del ETF, 4x1000, retención, intereses perdidos, costo de compra, comisión inmobiliaria, descuento de venta urgente y costo de venta; coincide con la definición de la Tabla 52 (no incluye predial ni administración).
- Primer cumplimiento de la meta: `t_meta` es el primer mes con patrimonio real (inmueble neto de costos de venta) en la meta. Entre réplicas que la alcanzan alguna vez (10 a 24 %, por estrategia) hay más que las que la tienen al final (6 a 19 %). La nota de la Tabla 58 ahora lo dice.
- Cifras de Python de la sección 7: se cotejaron contra los análisis originales del equipo (`insumos_equipo`): de unas 620 cifras con decimales quedaron 32 sin coincidencia directa, y todas tienen respaldo en algún archivo del proyecto. Se recalcularon de forma independiente: Shapiro-Wilk de la inflación (0,057 en nivel; 0,81 en logaritmo), Welch de la valorización del IPVU por estado (0,275 de una cola; 0,550 de dos colas), los parámetros de la t del ETF (ν = 5,42; 0,668 %; 4,132 %) y los AIC del logit de la TPM (336,4 y 331,5).
- Correcciones: en 7.5.9 «la máxima verosimilitud da 5,34» pasa a 5,42 (valor reproducible, igual al de la Tabla 28); el p de Welch de 0,275 se rotula «de una cola, 0,550 de dos colas» en el informe y en tres celdas del Excel.
- Siguen sin recalcular a mano (salen de simulaciones de 3.000 trayectorias, no de una hoja): 5,3 %/5,4 % de rendimiento anual y 40,5 %/40,8 % de caída máxima de la sección 7.5.

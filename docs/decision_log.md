# Registro de decisiones

Cada entrada dice qué se decidió, por qué y qué queda abierto. Las marcadas **Equipo** necesitan visto bueno del grupo antes de la entrega.

## Modelos que se reconciliaron

El Word de los puntos 1 a 7 y el diagrama de flujo del equipo no siempre usaban el mismo modelo para un proceso. No borré ninguno: el de referencia es el que se simula por defecto y el otro queda como opción en la aplicación y entra al análisis de sensibilidad a la elección de modelos.

| Proceso | Referencia | Alternativa | Razón |
|---|---|---|---|
| Estados económicos | Markov del equipo (crisis si el PIB crece menos de 1 %, expansión si más de 4 %) | ninguna | Es la del diagrama y la hoja MARKOV |
| ETF | t de Student con ν = 5 y volatilidad por estado | Jones-Faddy sin estados | La del equipo conecta el ETF con la economía; Jones-Faddy recoge mejor la asimetría (−1,12 en los datos) |
| Valorización | Normal del logaritmo por estado (IPVNBR Medellín) | AR(1) del IPVU nacional | La del equipo es de Medellín y depende del estado |
| Margen del CDT | M5: TPM + a + b·ΔTPM + residuo observado | Normal por fase de la TPM | M5 se estimó con 18 años de datos semanales y reproduce la reacción del margen a los cambios de tasa. Los valores por fase del diagrama (2,68 / 0,86 / 0,16) eran de una versión anterior |
| Reajuste del canon | Modelo DANE con tope en el IPC previo (Ley 820) | Sube todo el IPC | El diagrama asume que siempre sube el IPC completo; los datos muestran que en promedio sube menos |
| Junta del Banco | Logit ordenado con tendencia de la inflación y brecha de Taylor | Tabla por tendencia de 7.2 | Ver la entrada siguiente |

Efecto de cambiar cada uno por su alternativa (400 réplicas, puntos de probabilidad de meta, E1 / E2 / E3 / E4): CDT por fase +1,0 / +0,3 / −0,5 / 0,0; canon con todo el IPC 0,0 / +2,5 / +1,8 / +1,3; Jones-Faddy −0,5 / +0,3 / −1,0 / +1,3; AR(1) del IPVU 0,0 / −13,7 / −18,5 / −12,5; tabla de la Junta +47,8 / +26,8 / +11,8 / +19,3. Las dos últimas cambian la conclusión y van en la discusión del informe.

## Junta del Banco de la República (Equipo)

Con la tabla de 7.2 (46 reuniones), la TPM se va al tope de 15 % en buena parte de las réplicas y la P(meta) de E1 pasa de 0,04 a 0,56. Eso no pasa en la historia: el máximo desde 2001 fue 13,25 %. Por eso la referencia es un logit ordenado con la tendencia de la inflación y la brecha de Taylor, TPM − (1 + π + 0,5(π − 3)), estimado con 202 decisiones de 2001 a 2026 fechadas por su vigencia (ver Correcciones de la auditoría).

Con el logit corregido la TPM mediana al año 10 queda en 7,25 % y 0,8 % de las réplicas toca el tope. La tasa real media de los 10 años es 3,44 puntos contra 1,23 histórica; en los años 6 a 10 baja a 2,70 y sigue bajando hacia lo que pide la regla de Taylor con la inflación simulada (4,6 %): 1,8 puntos. El exceso viene del arranque (12,25 % de TPM con 6,3 % de inflación) y es una limitación declarada. Estimar solo con 2017-2026 no cambia la P(meta) en más de 1 punto.

## Supuestos del equipo que quedaron como parámetros

- Necesidades de liquidez: una cada 24 meses en expansión y normal, una cada 12 en crisis; monto lognormal con mediana de 20 millones en pesos de hoy y σ = 0,8. Es el factor que más mueve el resultado después de la inflación y la valorización.
- Vacancia: no renovar con probabilidad 5 / 10 / 30 % según el estado; duración gamma con forma 1,1 y media 1 / 1,5 / 6 meses. La validación da 3 % del tiempo desocupado contra 0,6 % de Ciencuadras 2026. Lo dejé así a propósito, más conservador. **Equipo**: confirmar.
- Estrés: los recortes de la TPM durante la crisis forzada (0,5 / 1 / 1,5 / 2 puntos) no salen de datos. **Equipo**.

## Decisiones de implementación

- La inflación de diciembre se sortea al inicio de cada año (t = 0, 12, 24...) y entre diciembres la ruta es lineal. Así la Junta y la cuenta conocen la inflación del mes sin mirar el futuro.
- El arriendo del mes en que cambia el contrato se paga con el canon viejo.
- 4x1000: se cobra sobre los retiros de la cuenta que pasan la exención mensual de 350 UVT y sobre lo que se recibe al vender ETF. UVT 2026 = 52.374. **Pendiente** confirmar el valor oficial.
- Si ni la venta urgente del inmueble alcanza, lo que falta queda como deuda y cuenta como episodio de iliquidez. La deuda se paga con el primer ingreso que llegue a la cuenta.
- Pintura, administración y monto de las necesidades se indexan con el índice de precios simulado.
- Patrimonio: el inmueble va a valor de mercado cada mes y, al final, neto de los costos de venta (4,84 %).
- E4 fija en la revisión el monto con que se renueva el CDT y aplaza las compras de ETF hasta que el CDT venza, para no pagar el castigo por venta anticipada.

## Correcciones de la auditoría independiente (noviembre de 2026)

- Valorización (C1): el modelo de referencia pasa a ser el AR(1) del IPVU (media 7,55 %, φ 0,676, σ 2,36 %). El IPVNBR mide vivienda nueva y el inmueble es usado; con el IPVU el estado económico no cambia la valorización (Welch, p = 0,275). La normal por estado con IPVNBR queda solo como alternativa en la sensibilidad. Adverso: media −2 puntos. Estrés: −5 % anual mientras dura la crisis.
- Junta (A2 y N2): el logit se estimaba con reuniones mensuales (2001-2016) y se simulaba con 8 por año; ahora se estima con el calendario de 8 reuniones en todos los años. Además, la TPM de fin de mes en el mes de reunión no era la decisión de esa reunión (el cambio entra en vigencia casi siempre el primer día hábil del mes siguiente) y la brecha de Taylor usaba la TPM del mes anterior, que ya traía parte del cambio que se quería explicar. Ahora cada reunión se mide con la TPM del día 15 del mes siguiente y la brecha usa la TPM vigente antes de decidir. Con 202 decisiones: tendencia 1,647 (p < 0,001), brecha −0,202 (p = 0,011), cortes −1,599 y 1,526. Antes la brecha daba −0,141 (p = 0,062). El logit estimado solo con 2017-2026 (78 reuniones, brecha −0,265) entra a la sensibilidad de modelos.
- Retención (M4): 4 % para los intereses del CDT y 7 % para los de la cuenta de ahorros (tarifa general de rendimientos financieros).
- Interés del CDT (B10): K·((1 + r)^(plazo/12) − 1); antes era simple y solo era correcto con plazo de 12 meses.
- ETF (M2): la media sale de la misma muestra que las volatilidades por estado (179 meses, ago-2011 a jun-2026).
- Cuenta (M8): α = 0,8978 con 4 cuentas. Ualá se excluye porque es un depósito de bajo monto con tope cercano a $10 M.
- Necesidades de liquidez (A1): la frecuencia (1 cada 24 meses) es un supuesto del equipo y los resultados se reportan con y sin necesidades. La IEFIC 2017-2018 del DANE (Medellín, 10 % de hogares con más ingreso) queda como referencia: un mes con déficit cubierto con ahorros o venta de activos cada 16 a 23 meses. No sirve para calibrar porque cuenta déficits de cualquier tamaño y la pregunta de cómo se cubren es sobre conducta habitual, no sobre el último mes. El monto y el doble en crisis tampoco tienen datos. Los microdatos no se publican (Ley 79 de 1993), solo data/fuentes/liquidez/iefic_resumen_medellin.csv.
- Venta urgente (M10): de 4 a 7 meses. La Lonja de Medellín (edición 275) reporta 225 días para vender vivienda usada; el descuento de 10 % es lo que se cede para no pasar de ahí y sigue como supuesto con sensibilidad de 5 % a 20 %.
- Vacancia (M7): la duración se apoya en Palomma (46 días); la Lonja reporta 121 días en oferta y entra a la sensibilidad con 4 meses; la probabilidad de no renovar y la vacancia en crisis siguen como supuestos.
- Canon (M6): se deja de usar Aptología (avisos de un agregador sin respaldo institucional, descarga truncada). El canon de 0,60 % del valor se apoya en la Lonja de Medellín (edición 275): mediana de canon ÷ precio de 0,59 % en 11 barrios y tamaños de Medellín.
- Umbral de crisis (A3): PIB < 1 % en todo el libro y en los datos. Con 0 % solo quedarían 4 trimestres de crisis desde 2006 y la de 2009 no contaría.
- Tornado (N1): el factor de valorización movía las medias de la normal por estado, que con el AR(1) ya no se usan, y su rango era 0. Ahora mueve la media del AR(1) del IPVU (5,55 % / 9,55 %).
- Cuenta en el libro (M8): la hoja TPM_Cuenta excluye Ualá y da α = 0,8978 con 19 tasas, igual que PARAMETROS.
- Estados económicos: las limitaciones de la cadena (intervalos amplios con 81 transiciones, persistencia en parte mecánica del crecimiento interanual, crisis simuladas más largas que las observadas) quedan en la hoja DECISIONES (D13).
- Inversión mínima (punto 5): cuenta sin monto mínimo; CDT de $50.000 a $500.000 según la entidad (prensa, 2026); ETF una unidad; inmueble completo. Ninguna limita al modelo porque las posiciones son de decenas de millones.
- Hoja INM del Excel: los avisos de Aptología se reemplazan en la copia del master por la tabla de la Lonja (edición 275) con la que se fija el canon.
- Impuestos (M5): no se modelan el impuesto de renta sobre intereses y arriendos ni la ganancia ocasional en la venta del inmueble; la retención se trata como costo definitivo aunque es un anticipo.

## Pendientes

- Reparaciones extraordinarias: apagadas (frecuencia 0). Sin datos de costo ni de frecuencia.
- Nombres de los integrantes en la portada del informe.

## Réplicas

El piloto de 200 réplicas pide entre 500 y 704 réplicas para un error relativo de 2 % en el patrimonio real medio. Usamos 1.068, que además da ±3 puntos en el peor caso para la probabilidad de meta. Con 2.400 réplicas de E2 la probabilidad de meta pasa de 37,9 % a 36,5 %, dentro del intervalo de 1.068.

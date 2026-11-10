"""Pone al día los puntos 1 a 7 del Word del equipo dentro de la copia que arma construir_informe.py.

El documento del equipo no se modifica: los cambios se hacen sobre el objeto Document ya abierto.
Cada cambio responde a una corrección de la auditoría (docs/decision_log.md) o al modelo que
el simulador usa hoy. Las referencias de las secciones se juntan en una sola lista al final.
"""
from __future__ import annotations

import copy
import json
import re
from pathlib import Path

import pandas as pd
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

RAIZ = Path(__file__).resolve().parents[1]


# ---------------------------------------------------------------- utilidades
def _cuerpo(doc):
    return list(doc.element.body.iterchildren())


def _texto(el) -> str:
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def _es_titulo(el) -> bool:
    s = el.find(".//" + qn("w:pStyle"))
    return s is not None and ("tulo" in s.get(qn("w:val")) or "eading" in s.get(qn("w:val")))


def _i(doc, inicio: str, titulo: bool | None = None, desde: int = 0) -> int:
    for k, el in enumerate(_cuerpo(doc)):
        if k < desde or el.tag != qn("w:p"):
            continue
        if _texto(el).strip().startswith(inicio) and (titulo is None or _es_titulo(el) == titulo):
            return k
    raise KeyError(inicio)


def _borrar(doc, i: int, j: int) -> None:
    """Borra los elementos del cuerpo con índice en [i, j)."""
    for el in _cuerpo(doc)[i:j]:
        el.getparent().remove(el)


def _borrar_entre(doc, inicio: str, hasta: str, titulo_inicio=None, titulo_hasta=None) -> None:
    i = _i(doc, inicio, titulo_inicio)
    _borrar(doc, i, _i(doc, hasta, titulo_hasta, desde=i + 1))


def _poner(par: Paragraph, texto: str) -> None:
    runs = par.runs
    runs[0].text = texto
    for run in runs[1:]:
        run._r.getparent().remove(run._r)


def _parrafo(doc, inicio: str) -> Paragraph:
    el = _cuerpo(doc)[_i(doc, inicio)]
    return Paragraph(el, doc.element.body)


def reemplazar(doc, viejo: str, nuevo: str) -> int:
    """Reemplaza texto en párrafos y celdas; devuelve cuántos párrafos cambió."""
    n = 0
    pars = list(doc.paragraphs) + [p for t in doc.tables for f in t.rows for c in f.cells for p in c.paragraphs]
    for p in pars:
        if viejo not in p.text:
            continue
        hecho = False
        for run in p.runs:
            if viejo in run.text:
                run.text = run.text.replace(viejo, nuevo)
                hecho = True
        if not hecho:
            _poner(p, p.text.replace(viejo, nuevo))
        n += 1
    return n


def _tabla(doc, encabezado: str) -> Table:
    for t in doc.tables:
        if t.rows[0].cells[0].text.strip() == encabezado.split(" | ")[0] and " | ".join(c.text for c in t.rows[0].cells).startswith(encabezado):
            return t
    raise KeyError(encabezado)


def _fila(t: Table, primera: str):
    for f in t.rows:
        if f.cells[0].text.strip().startswith(primera):
            return f
    raise KeyError(primera)


def _celdas(fila, valores: dict[int, str]) -> None:
    for j, v in valores.items():
        _poner(fila.cells[j].paragraphs[0], v)


def _agregar_fila(t: Table, valores: list[str]) -> None:
    nueva = copy.deepcopy(t.rows[-1]._tr)
    t._tbl.append(nueva)
    _celdas(t.rows[-1], dict(enumerate(valores)))


def _mover_antes(doc, ancla_idx: int, generar) -> None:
    """Escribe con generar() al final del documento y mueve lo escrito antes del elemento ancla_idx."""
    ancla = _cuerpo(doc)[ancla_idx]
    n0 = len(_cuerpo(doc))
    generar()
    nuevos = _cuerpo(doc)[n0 - 1:-1]
    for el in nuevos:
        ancla.addprevious(el)


# ---------------------------------------------------------------- cambios
def actualizar(w, P, c, pc) -> list[str]:
    """Aplica los cambios a w.doc y devuelve las referencias de los puntos 1 a 7 para la lista final."""
    doc = w.doc
    j = json.loads((RAIZ / "data" / "metadata" / "junta_logit.json").read_text(encoding="utf-8"))

    # Archivos y nombres de hojas del libro viejo.
    reemplazar(doc, "la hoja Parametros_simulador", "la hoja PARAMETROS del Excel maestro o la aplicación")
    reemplazar(doc, "Estos valores se cambian desde la hoja PARAMETROS del Excel maestro o la aplicación. Los del caso base son los que usamos "
                    "para comparar estrategias; los rangos son los que recorreremos en el experimento de capital y meta.",
               "Estos valores se cambian desde la hoja PARAMETROS del Excel maestro o desde la aplicación. Los del caso base son los "
               "que usamos para comparar estrategias; los rangos son los que recorrimos en el experimento de capital y meta (sección 17).")

    # ---------------- 5. Universo de inversión
    _poner(_parrafo(doc, "La cuenta es el colchón del portafolio"),
           "La cuenta es el colchón del portafolio. Recibe los cánones, los intereses de los CDT y las ventas de activos, y de ella "
           "salen los gastos del inmueble y los imprevistos. Las cuentas digitales pagan una fracción estable de la tasa del Banco de "
           f"la República; con 19 tasas de cuatro cuentas entre diciembre de 2023 y octubre de 2026 estimamos que esa fracción es "
           f"{c(P.f('cta.alfa'), 3)} (sección 7.3).")
    t = _tabla(doc, "Regla | Valor en el modelo | Tipo y fuente")
    tasa = P.f("cta.alfa") * P.f("tpm.inicial") / 100
    _celdas(_fila(t, "Rendimiento"), {1: f"r = α × TPM con α = {c(P.f('cta.alfa'), 3)}. Con la TPM en 12,25 % paga {pc(tasa, 2)} E.A.",
                                      2: "Estimado con 19 tasas de 4 cuentas (sección 7.3)"})
    _celdas(_fila(t, "Cambio de tasa"), {1: "Se recalcula en cada reunión de la Junta"})
    _celdas(_fila(t, "Monto mínimo"), {1: "No tiene"})

    t = _tabla(doc, "Regla | Valor en el modelo | Sustento")
    _celdas(_fila(t, "Monto mínimo"), {1: "Entre $50.000 y $500.000 según la entidad y el canal. No limita al modelo: el CDT más "
                                          "pequeño de las estrategias es de $80 millones",
                                       2: "Noticias Caracol (enero de 2026) y El País (septiembre de 2026)"})
    _celdas(_fila(t, "Seguro de depósitos"), {1: "Fogafín cubre hasta $50 millones por persona y por entidad. Con más de eso en CDT habría "
                                                 "que repartirlo entre varias entidades; el modelo no lo representa",
                                              2: "Fogafín"})
    _celdas(_fila(t, "Retención en la fuente"), {1: f"{pc(P.f('imp.retencion'), 0)} sobre los intereses del CDT y "
                                                    f"{pc(P.f('imp.retencion_cuenta'), 0)} sobre los de la cuenta. Se trata como costo "
                                                    "final, aunque es un anticipo del impuesto de renta"})

    t = [x for x in doc.tables if x.rows[0].cells[0].text == "Regla" and any(f.cells[0].text.startswith("Comisión de administración") for f in x.rows)][0]
    _celdas(_fila(t, "Rendimiento mensual"), {1: f"t de Student con ν = 5 y volatilidad según el estado económico ({pc(P.f('etf.vol_N'), 1)} "
                                                 f"mensual en expansión y normal, {pc(P.f('etf.vol_C'), 1)} en crisis); media "
                                                 f"{pc(P.f('etf.mu'), 2)} mensual. Alternativa: t asimétrica de Jones y Faddy",
                                              2: "Estimado (sección 7.5)"})
    _agregar_fila(t, ["Inversión mínima", "Una unidad del fondo, que se compra en bolsa por medio de una comisionista. No limita al "
                                          "modelo, que compra millones de pesos en ETF", "Mecánica de la Bolsa de Valores de Colombia"])

    t = [x for x in doc.tables if x.rows[0].cells[0].text == "Regla" and any(f.cells[0].text.startswith("Precio de compra") for f in x.rows)][0]
    _celdas(_fila(t, "Canon inicial"), {2: "Lonja de Medellín (2025), edición 275: canon ÷ precio de venta mediano de 0,59 % mensual en "
                                           "11 combinaciones de barrio y tamaño de Medellín (sección 5.4.1)"})
    _celdas(_fila(t, "Venta ordinaria"), {2: "Lonja de Medellín (2025): 225 días para vender vivienda usada; comisión publicada"})
    _celdas(_fila(t, "Venta urgente"), {1: f"Descuento de {pc(P.f('inm.descuento_urgente'), 0)} (sensibilidad 5 % a 20 %); incluye una "
                                           "comisión más alta para que la inmobiliaria la venda rápido",
                                        2: "Supuesto"})
    _agregar_fila(t, ["Inversión mínima", "El apartamento completo: $300 millones más 2,5 % de costos de compra. No se puede comprar ni "
                                          "vender por partes", "Decisión del equipo"])
    _poner(_parrafo(doc, "El canon de 0,60 % ($1,8 millones)"),
           "Con 0,60 % el canon inicial es $1,8 millones al mes. Al vencer cada contrato el canon nuevo se fija otra vez en 0,60 % "
           "del valor actualizado del apartamento.")

    # 5.4.1: el canon sale de la Lonja (se quita el estudio con avisos de Aptología).
    i = _i(doc, "5.4.1 De dónde sale el canon", True)
    _borrar(doc, i + 1, _i(doc, "5.4.2 Predial", True))
    lonja = pd.read_csv(RAIZ / "data" / "fuentes" / "inmueble" / "processed" / "A3_lonja_canon_precio_admin.csv")
    lonja = lonja[(lonja.municipio == "Medellín") & lonja.canon_sobre_precio_mensual.notna()]

    def canon():
        w.p("Para el canon usamos el boletín de mercado inmobiliario residencial de la Lonja de Propiedad Raíz de Medellín y "
            "Antioquia (edición 275, agosto de 2025). La Lonja publica, por barrio y por tamaño, el precio de venta y el canon por "
            "metro cuadrado de vivienda usada. Preferimos esta fuente a los portales de avisos porque es gremial, publica su "
            "metodología y reporta los dos precios del mismo inmueble típico, que es justo lo que necesitamos para la relación "
            "canon ÷ precio. En el boletín los encabezados de canon y precio de esa tabla están intercambiados; los ubicamos por "
            "su orden de magnitud.")
        filas = [[x.barrio, x.banda_m2.replace("-", " a ") + " m²", c(x.precio_venta_m2), c(x.canon_m2), pc(x.canon_sobre_precio_mensual, 2)]
                 for x in lonja.itertuples()]
        w.tabla("5.4", "Canon y precio de venta por m² en barrios de Medellín (Lonja, edición 275).",
                ["Barrio", "Tamaño", "Precio de venta por m²", "Canon mensual por m²", "Canon ÷ precio"], filas)
        r = lonja.canon_sobre_precio_mensual
        w.p(f"En las {len(r)} combinaciones de Medellín con los dos datos, el canon mensual es entre {pc(r.min(), 2)} y "
            f"{pc(r.max(), 2)} del precio, con mediana de {pc(r.median(), 2)}. Usamos 0,60 %, y en la sensibilidad lo movemos entre "
            "0,50 % y 0,70 %. Antes habíamos usado 0,40 % ($1,2 millones), que quedaba por debajo de todos los barrios del "
            "boletín.")
    _mover_antes(doc, _i(doc, "5.4.2 Predial", True), canon)

    # ---------------- 6 y 7: tablas generales
    t = _tabla(doc, "Fuente de incertidumbre | Qué cambia en el simulador")
    _celdas(_fila(t, "Margen del CDT"), {2: "Margen sobre la TPM según el cambio reciente de la TPM"})
    _celdas(_fila(t, "Necesidad extraordinaria"), {4: "Sección 12"})
    _poner(_parrafo(doc, "Los procesos están conectados."),
           "Los procesos están conectados. La inflación simulada alimenta las decisiones de la Junta, la TPM resultante fija la tasa "
           "de la cuenta y la de los CDT nuevos, y el estado de la economía (secciones 7.8 y 8) mueve la vacancia, la volatilidad del "
           "ETF y la frecuencia de los gastos imprevistos. El ETF no se correlaciona directamente con la inflación ni con las tasas "
           "porque los datos no muestran esa relación (sección 7.5).")
    t = _tabla(doc, "Proceso | Depende de | Alimenta a")
    _celdas(_fila(t, "TPM"), {1: "Tendencia de la inflación y distancia a la regla de Taylor"})
    _celdas(_fila(t, "ETF"), {1: "Estado de la economía (solo la volatilidad)"})
    _celdas(_fila(t, "Valorización"), {1: "Su valor del año anterior"})

    # ---------------- 7.2: logit ordenado que usa el simulador
    _poner(_parrafo(doc, "Hoja del libro: TPM_Cuenta. Dónde está cada paso"),
           "El primer estudio (7.2.1 a 7.2.6) usa las 46 reuniones de 2021 a 2026 y una tabla de decisiones según la tendencia de "
           "la inflación. El modelo que usa el simulador (7.2.7) es un logit ordenado estimado con 25 años de decisiones; allí se "
           "explica por qué se cambió. La TPM oficial diaria desde 1998 (SUAMECA) también sirve en 7.4 para estimar el margen "
           "del CDT.")
    cal = j["calibracion"]

    def logit():
        w.h("7.2.7 Modelo que usa el simulador: logit ordenado", 3)
        w.p("Cuando simulamos diez años con la tabla de 7.2.3, la TPM terminó en el tope de 15 % en cerca de la mitad de las "
            "réplicas, algo que no pasó nunca desde 2001 (el máximo fue 13,25 %). La tabla solo mira hacia dónde va la inflación: si "
            "la inflación sube, la Junta sube, aunque la tasa ya esté muy por encima de la inflación. Le faltaba una variable que "
            "frene las subidas. Por eso estimamos un modelo con más datos y con esa variable.")
        w.p(f"Datos. Usamos la TPM diaria del Banco de la República y la inflación mensual del DANE de 2001 a 2026, con una "
            f"decisión por cada mes del calendario actual de la Junta (ocho reuniones al año): {j['n']} decisiones, de las cuales "
            f"{j['decisiones']['1']} fueron subidas, {j['decisiones']['-1']} bajadas y {j['decisiones']['0']} sin cambio. Antes "
            "de 2017 la Junta se reunía cada mes, así que en esos años la decisión es el cambio acumulado desde la reunión anterior "
            "del calendario. Cada decisión se mide con la TPM del día 15 del mes siguiente, porque el cambio entra en vigencia casi "
            "siempre el primer día hábil después de la reunión; así la variable explicativa no contiene el cambio que se quiere "
            "explicar.")
        w.p("Modelo. Un logit ordenado con dos variables: la tendencia de la inflación (cambio en los tres meses previos) y la "
            "distancia entre la TPM vigente y la que pediría una regla de Taylor, TPM − (1 + π + 0,5·(π − meta)). Si la tasa ya está "
            "por encima de la regla, la brecha es positiva y la probabilidad de subir baja. P(baja) = F(k1 − x) y P(sube) = 1 − "
            "F(k2 − x), con x = b_tend·tendencia + b_Taylor·brecha y F la función logística.")
        w.tabla("7.2", f"Logit ordenado de la Junta, {j['desde'][:4]} a {j['hasta'][:4]}.", ["Coeficiente", "Valor", "p"],
                [["Tendencia de la inflación (b_tend)", c(j["b_tend"], 3), "< 0,001"],
                 ["Brecha de Taylor (b_Taylor)", c(j["b_taylor"], 3), c(j["p_taylor"], 3)],
                 ["Corte k1", c(j["k1"], 3), ""], ["Corte k2", c(j["k2"], 3), ""]])
        w.p(f"Pruebas. Las dos variables son significativas al 5 %. Agregar la brecha de Taylor al modelo que solo tiene la "
            f"tendencia baja el AIC de {c(j['aic_solo_tendencia'], 1)} a {c(j['aic_con_taylor'], 1)} (razón de verosimilitud "
            f"{c(j['lr'], 2)}, p = {c(j['lr_p'], 3)}). El logit ordenado supone que las dos variables pesan igual para pasar de bajar "
            f"a mantener que para pasar de mantener a subir. Lo revisamos con dos logit binarios y una prueba de Brant aproximada: "
            f"W = {c(j['brant_w'], 2)} con 2 grados de libertad, p = {c(j['brant_p'], 2)}, así que no se rechaza. La tabla compara "
            "lo que predice el modelo con lo que pasó, agrupando las reuniones en tres partes iguales según el índice x:")
        w.tabla("7.2", "Probabilidades predichas frente a frecuencias observadas.",
                ["Reuniones", "P(baja) predicha", "Bajas observadas", "P(sube) predicha", "Subidas observadas"],
                [[k.capitalize() + " x", pc(v["p_baja"]), pc(v["obs_baja"]), pc(v["p_sube"]), pc(v["obs_sube"])] for k, v in cal.items()])
        sub = ", ".join(f"{k.replace('.', ',')} ({pc(v, 0)})" for k, v in j["subidas"].items())
        baj = ", ".join(f"{k.replace('.', ',')} ({pc(v, 0)})" for k, v in j["bajadas"].items())
        w.p(f"Tamaño del cambio. Si la Junta mueve la tasa, el tamaño se sortea con las frecuencias observadas en las mismas "
            f"decisiones. Subidas: {sub}. Bajadas: {baj}. La TPM se mantiene entre 1,75 % y 15 %.")
        w.p("Validación y límites. Con el logit la TPM toca el tope solo en 0,4 % de las réplicas y su mediana al año 10 es 7,25 %. "
            "La tasa real simulada sigue siendo más alta que la histórica, sobre todo al principio, porque hoy la tasa real es de 6 "
            "puntos (sección 16.2). Si el logit se estima solo con 2017 a 2026, cuando la Junta ya tenía ocho reuniones, la brecha "
            "pesa algo más (−0,265), pero los resultados del simulador cambian menos de un punto (sección 16.3). La tabla de 7.2.3 "
            "queda como alternativa en la aplicación.")
    _mover_antes(doc, _i(doc, "7.3 Rendimiento de la cuenta remunerada", True), logit)

    # ---------------- 7.3: cuenta sin Ualá
    _poner(_parrafo(doc, "Recopilamos 24 tasas"),
           "Recopilamos tasas efectivas anuales de cuentas digitales en cinco fechas, de diciembre de 2023 a octubre de 2026, "
           "publicadas en El Tiempo, El Colombiano y Rankia. Usamos 19 tasas de cuatro cuentas (Nu, Pibank, Lulo y RappiPay). "
           "Dejamos por fuera a Ualá porque su tasa alta solo aplica hasta un saldo cercano a $10 millones, y el inversionista "
           "tendría decenas de millones en la cuenta. Para cada tasa calculamos la razón α = tasa de la cuenta / TPM vigente en esa "
           "fecha. Con cinco fechas no hay una serie de tiempo que analizar, y no hay valores atípicos.")
    reemplazar(doc, "La estimación es el promedio de las 24 razones:", "La estimación es el promedio de las 19 razones:")
    t = _tabla(doc, "Concepto | Valor")
    _celdas(_fila(t, "α"), {1: c(P.f("cta.alfa"), 3)})
    _celdas(_fila(t, "Desviación estándar"), {1: "0,143"})
    _celdas(_fila(t, "Intervalo de confianza"), {1: "0,829 a 0,967"})
    _poner(_parrafo(doc, "Si la cuenta paga una fracción fija de la TPM"),
           "Si la cuenta paga una fracción fija de la TPM, la razón no debería depender del nivel de la TPM. La correlación entre las "
           "dos es −0,26 (t = −1,09; p = 0,29), así que no se rechaza que α sea constante.")
    _poner(_parrafo(doc, "Cada vez que cambia la TPM se recalcula"),
           f"Cada vez que cambia la TPM se recalcula la tasa de la cuenta. Con la TPM en 12,25 % la cuenta paga {pc(tasa, 2)} E.A., "
           f"es decir {pc((1 + tasa) ** (1 / 12) - 1, 3)} mensual. Sobre los intereses se descuenta una retención de "
           f"{pc(P.f('imp.retencion_cuenta'), 0)}. La dependencia con el resto del modelo pasa toda por la TPM.")
    _poner(_parrafo(doc, "Son 24 datos de cinco entidades"),
           "Son 19 datos de cuatro entidades, tomados de prensa y no de una base oficial. Los promedios ponderados de captación de "
           "la Superintendencia Financiera dan un α más bajo, cercano a 0,82, porque incluyen cuentas tradicionales que pagan menos; "
           "en el escenario adverso usamos 0,80. El modelo no incluye cambios de política de cada entidad, y no encontramos topes "
           "de saldo en las cuatro cuentas que usamos.")

    # ---------------- 7.5: ETF
    i5, i6 = _i(doc, "7.5 Renta variable", True), _i(doc, "7.6 Valorización del inmueble", True)
    refs_etf = [_texto(el) for el in _cuerpo(doc)[_i(doc, "Referencias del ETF", True) + 1:i6] if _texto(el).strip()]
    _borrar_entre(doc, "Archivos y reproducibilidad del ETF", "7.6 Valorización del inmueble", True, True)
    _borrar(doc, _i(doc, "Hoja del libro: ETF. El código"), _i(doc, "Hoja del libro: ETF. El código") + 1)
    _borrar(doc, _i(doc, "Datos descargados y procesados el 29"), _i(doc, "Datos descargados y procesados el 29") + 1)
    reemplazar(doc, " (carpeta src). Ninguna se fijó a mano.", ". Ninguna se fijó a mano.")
    # Tablas de comparación de fuentes, figura de control y detalles de calidad de Yahoo.
    _borrar_entre(doc, "Revisamos tres fuentes.", "La serie principal es el rendimiento total mensual")
    _borrar_entre(doc, "Tabla 2. Rendimiento total", "Yahoo Finance falla justo")
    _borrar_entre(doc, "Tabla 3. Distribuciones por año", "Para confirmar la serie")
    # Figura 1 (imagen y leyenda), nota de calidad de Yahoo y 7.5.1.3.
    k = _i(doc, "Figura 1. Izquierda")
    _borrar(doc, k - 1, _i(doc, "7.5.2 Análisis exploratorio", True))
    reemplazar(doc, " (Tabla 2).", ".")
    reemplazar(doc, " (Tabla 3).", ".")
    reemplazar(doc, " (Figura 1).", ".")
    # Subperiodos: se deja el texto, sin la tabla ni la figura de autocorrelación.
    _borrar_entre(doc, "Tabla 7. Rendimiento por subperiodos", "Entre 2011 y 2023 el ICOLCAP")
    _borrar_entre(doc, "Figura 4. Autocorrelación", "Los rendimientos no tienen memoria lineal")
    k = _i(doc, "Los rendimientos no tienen memoria lineal")
    if _cuerpo(doc)[k - 1].find(".//" + qn("w:drawing")) is not None:
        _borrar(doc, k - 1, k)
    _borrar_entre(doc, "Tabla 9. Sensibilidad de los estadísticos", "La Tabla 9 deja ver")
    reemplazar(doc, "La Tabla 9 deja ver algo importante. Sin esos tres meses", "Hay algo importante. Sin esos tres meses")
    _borrar_entre(doc, "La mezcla separa con claridad", "7.5.7 Bondad de ajuste", None, True)
    # Bondad de ajuste: se deja solo la figura Q-Q.
    k5, k7 = _i(doc, "Figura 5. Histograma de rendimientos"), _i(doc, "Figura 7. Gráficos Q-Q")
    _borrar(doc, k5 - 1, k7 - 1)
    k = _i(doc, "Figura 7. Gráficos Q-Q")
    _borrar(doc, k + 1, _i(doc, "7.5.8 Comparación entre modelos", True))
    _borrar_entre(doc, "Tabla 14. Cuantiles extremos", "En la muestra hubo dos meses")
    _borrar_entre(doc, "Los cuatro modelos asimétricos", "En la muestra hubo dos meses")
    _poner(_parrafo(doc, "Seleccionamos la t asimétrica de Jones y Faddy"),
           "Si se mira solo la distribución, sin estados de la economía, el mejor modelo es la t asimétrica de Jones y Faddy con "
           "a = 2,545, b = 3,448, loc = 2,471 % y escala = 4,062 %. Las razones, en orden de peso:")
    _poner(_parrafo(doc, "Dejamos dos alternativas documentadas"),
           "En 7.5.11 explicamos por qué el simulador termina usando una t simétrica con volatilidad según el estado económico, y "
           "deja esta como alternativa.")
    _borrar_entre(doc, "7.5.9.1 Incertidumbre de los parámetros", "7.5.10 Dependencia con otras variables", True, True)
    _poner(_parrafo(doc, "El enunciado no permite suponer que todas las variables son independientes. Para saber"),
           "Las variables del modelo no se pueden tratar como independientes. Para saber con qué se mueve el ICOLCAP "
           "armamos un panel mensual con la tasa de cambio, el S&P 500, el petróleo Brent, la inflación, la TPM y la tasa de los "
           "CDT sobre los mismos 182 meses, y medimos la dependencia de tres formas: correlaciones contemporáneas, correlaciones "
           "con rezago y comportamiento por estado económico. Usamos Pearson para la relación lineal y Spearman y Kendall, que se "
           "basan en rangos y no se dejan arrastrar por los meses extremos.")
    _borrar_entre(doc, "Figura 9. Matriz de correlación", "Para ver el efecto conjunto")
    k = _i(doc, "Para ver el efecto conjunto")
    while _cuerpo(doc)[k - 1].find(".//" + qn("w:drawing")) is not None:
        _borrar(doc, k - 1, k)
        k -= 1
    _borrar_entre(doc, "7.5.10.3 Dependencia en las colas", "7.5.10.4 Comportamiento por estado", True, True)
    _borrar_entre(doc, "7.5.10.5 Regímenes del propio ETF", "7.5.10.6 Qué implica", True, True)
    reemplazar(doc, "7.5.10.4 Comportamiento", "7.5.10.3 Comportamiento")
    reemplazar(doc, "7.5.10.6 Qué implica", "7.5.10.4 Qué implica")
    _poner(_parrafo(doc, "Por eso proponemos que el ETF se conecte"),
           "Por eso el ETF se conecta con las demás variables a través de la cadena de estados económicos (sección 8). El estado "
           "cambia la volatilidad del fondo y no su media, como se ve en 7.5.11. Así una crisis vuelve más riesgoso al ETF al mismo "
           "tiempo que alarga la vacancia y trae más gastos imprevistos, sin forzar correlaciones lineales que los datos no "
           "muestran. No modelamos la tasa de cambio, que es la variable que más se mueve con el ETF (−0,52).")
    # Estado de la economía con el umbral de 1 % que usa el simulador.
    k = _i(doc, "Los 12 meses de crisis del PIB")
    t_old = _cuerpo(doc)[k - 2] if _cuerpo(doc)[k - 2].tag == qn("w:tbl") else None
    if t_old is not None:
        t_old.getparent().remove(t_old)
    k = _i(doc, "Los 12 meses de crisis del PIB")

    def estados_etf():
        w.tabla("7.5", "Rendimiento del ETF según el estado de la economía (umbral de crisis de 1 %, ago-2011 a jun-2026).",
                ["Estado", "Meses", "Media log mensual", "Desviación"],
                [["Expansión", "44", "−0,38 %", "5,26 %"], ["Normal", "111", "0,63 %", "4,46 %"], ["Crisis", "24", "0,48 %", "8,78 %"]])
    _mover_antes(doc, k, estados_etf)
    _poner(_parrafo(doc, "Los 12 meses de crisis del PIB"),
           "Con el umbral de crisis de 1 % que usa el simulador (sección 8), la media del ETF no cambia con el estado "
           "(Kruskal-Wallis, p = 0,39), pero la desviación en crisis es casi el doble de la de los otros dos estados (Levene, "
           "p = 0,004). Esa diferencia descansa en pocos meses: sin marzo de 2020 deja de ser significativa (p = 0,080). La "
           "tomamos igual, porque es lo que se espera de una crisis y el modelo necesita que el estado de la economía mueva la renta variable, "
           "pero la reportamos con esa advertencia.")
    _poner(_parrafo(doc, "La t por máxima verosimilitud centra en 0,67 %"),
           "La t por máxima verosimilitud centra en 0,67 % y sobreestimaría el rendimiento en cerca de 3 puntos al año, porque no "
           "capta la asimetría. Fijando la posición en la media, la t simple reproduce el rendimiento, aunque espera menos caídas "
           "extremas que la t asimétrica (0,94 frente a 1,29 meses, con 2 observados). En 3.000 trayectorias de 182 meses, la t "
           "asimétrica da rendimiento anual de 5,4 % y caída máxima mediana de 40,8 %, y la t de 5,3 % y 40,5 %, frente a 5,3 % y "
           "41,6 % observados: las dos reproducen la historia casi igual.")

    def modelo_etf():
        w.p(f"Modelo que usa el simulador. Elegimos la t de Student con ν = 5 (la máxima verosimilitud da 5,34), la media de "
            f"{pc(P.f('etf.mu'), 2)} mensual y la desviación del estado: {pc(P.f('etf.vol_N'), 1)} en expansión y normal, "
            f"{pc(P.f('etf.vol_C'), 1)} en crisis. La media y las desviaciones salen de los mismos 179 meses que tienen estado "
            "económico (agosto de 2011 a junio de 2026). La razón para dejar la Jones-Faddy como alternativa es la dependencia con "
            "la economía: la Jones-Faddy no tiene estados, y la t con volatilidad por estado sí conecta el ETF con la economía. El "
            "costo es que la t es simétrica y no reproduce la asimetría de −1,12 de los datos. Usar la Jones-Faddy en el simulador "
            "cambia la probabilidad de meta unos tres puntos (sección 16.3).")
        w.p("Limitaciones. La muestra tiene 15 años y casi todo el rendimiento positivo del fondo está en 2024 a 2026; hasta 2023 la "
            "media mensual era negativa. Las colas pesadas las ponen tres meses (marzo de 2020, junio de 2022 y uno más). Y la "
            "volatilidad de crisis se estimó con 24 meses.")
    _mover_antes(doc, _i(doc, "7.6 Valorización del inmueble", True), modelo_etf)
    reemplazar(doc, "t asimétrica de Jones y Faddy (elegida)", "t asimétrica de Jones y Faddy (mejor ajuste)")
    t = _tabla(doc, "Prueba (nivel 5 %) | Estadístico | p | Decisión")
    _celdas(_fila(t, "Rendimiento igual en los tres estados"), {0: "Rendimiento igual en los tres estados del PIB, umbral de 1 % (Kruskal-Wallis)",
                                                                2: "0,39", 3: "No se rechaza"})
    _celdas(_fila(t, "Varianza igual en los tres estados"), {0: "Varianza igual en los tres estados del PIB, umbral de 1 % (Levene)",
                                                             2: "0,004", 3: "Se rechaza"})
    _poner(_parrafo(doc, "La lectura es la de los apartados anteriores."),
           "La lectura es la de los apartados anteriores. El rendimiento medio no es distinto de cero (el fondo ganó 0,43 % mensual, "
           "pero con una desviación de 5,4 %). Los meses son independientes entre sí y no hay agrupamiento de volatilidad, así que se "
           "pueden sortear por separado. La distribución no es normal ni simétrica, y el nivel medio cambió en 2024. El estado de la "
           "economía no explica la media del rendimiento, pero sí su varianza, como se ve en la tabla siguiente.")
    _renumerar(doc, "7.5 Renta variable", "7.6 Valorización del inmueble", "7.5")
    i5, i6 = _i(doc, "7.5 Renta variable", True), _i(doc, "7.6 Valorización del inmueble", True)
    for el in _cuerpo(doc)[i5:i6]:
        if el.tag == qn("w:p") and re.search(r"\bpuntos? \d", _texto(el)):
            p = Paragraph(el, doc.element.body)
            nuevo = re.sub(r"\bpuntos (\d+) a (\d+)\b", r"pasos \1 a \2", p.text)
            nuevo = re.sub(r"\bpunto (\d+)\b(?! del enunciado)", r"paso \1", nuevo)
            nuevo = nuevo.replace("paso 8 del enunciado", "punto 8 del enunciado")
            if nuevo != p.text:
                _poner(p, nuevo)

    # ---------------- 7.6 y 7.7
    _poner(_parrafo(doc, "El valor del inmueble se actualiza una vez al año con el AR(1) nominal"),
           "El valor del inmueble se actualiza una vez al año con el AR(1) nominal:")
    _poner(_parrafo(doc, "Aquí r_t es la variación logarítmica"),
           "Aquí r_t es la variación logarítmica del precio en el año t y V_t el valor del apartamento. La variación esperada "
           "converge a 7,55 % anual y cada año la sorpresa e_t la aleja de ese nivel. La valorización no cambia con el estado de la "
           "economía: entre estados la diferencia no es significativa (Welch, p = 0,275). Solo en el escenario de estrés, mientras "
           "dura la crisis forzada, el apartamento se desvaloriza 5 % al año; en el adverso la media baja 2 puntos.")
    reemplazar(doc, "Hoja del libro: INM (IPVU y su modelo) y PIB. El análisis detallado está en el notebook Simulación "
                    "Inmueble/notebooks/02_validacion_inmobiliaria.ipynb.", "Hoja del libro: INM (IPVU y su modelo) y PIB.")
    reemplazar(doc, " No tenemos una serie de IPC de arriendos, así que el reajuste del canon se simula con la inflación simulada, como "
                    "dice la Ley 820, y no se prueba si los arriendos realmente suben como el IPC.", "")

    def vacancia():
        res = pd.read_csv(RAIZ / "outputs" / "resultados" / "resumen.csv").set_index(["escenario", "estrategia"])
        vac = float(res.loc[("base", "E2"), "tiempo_vacante_media"])
        w.p(f"Validación. Con estos valores el apartamento pasa {pc(vac, 1)} del tiempo desocupado en la simulación, más que el "
            "0,6 % que reporta Ciencuadras para Medellín. Lo dejamos así a propósito: preferimos un modelo que exagere la vacancia "
            "a uno que la ignore. La Lonja reporta 121 días en oferta antes de arrendar, bastante más que los 46 días de Palomma; "
            "nos quedamos con Palomma porque mide el segmento y la ciudad del inmueble, y probamos la media de 4 meses de la Lonja "
            "en la sensibilidad (sección 16.3).")
    _mover_antes(doc, _i(doc, "7.7.11 Limitaciones", True), vacancia)

    # ---------------- 7.8 Estado de la economía
    _poner(_parrafo(doc, "El enunciado pide analizar la dependencia entre procesos y habla de estados"),
           "Para analizar la dependencia entre procesos se define el estado de la economía. Se distinguen tres "
           "estados con el crecimiento anual del PIB real trimestral (DANE, base 2015, encadenado con las bases 2005 y 1994): crisis "
           "si crece menos de 1 % frente al mismo trimestre del año anterior, expansión si crece 4 % o más y normal en otro caso. "
           "Con 1 % entran como crisis la recesión de 2008 y 2009, la pandemia y el estancamiento de 2023 y 2024; con 0 % solo "
           "quedarían cuatro trimestres desde 2006. Se usaron los 126 trimestres, de 1995 a junio de 2026.")
    _poner(_parrafo(doc, "El crecimiento medio es 3,2 %"),
           "El crecimiento medio es 3,2 % con desviación de 3,9 puntos. Hay 21 trimestres de crisis (1996 y 1997, 1998 y 1999, "
           "2002, 2008 y 2009, 2020, y 2023 y 2024), 58 normales y 47 de expansión; desde 2006 son 12, 40 y 30. Con esa definición "
           "el estado de la economía cambia la volatilidad del ETF pero no su media (7.5), no cambia la valorización (7.6) y es la "
           "base de los escenarios de vacancia (7.7). La matriz de transición entre estados está en la sección 8.")

    # ---------------- 7.9 Resumen de modelos
    t = _tabla(doc, "Proceso | Modelo para el simulador")
    _celdas(_fila(t, "TPM"), {1: "Logit ordenado: tendencia de la inflación y brecha de Taylor; tamaños empíricos",
                              2: f"b_tend = {c(j['b_tend'], 2)}; b_Taylor = {c(j['b_taylor'], 2)}; cortes {c(j['k1'], 2)} y {c(j['k2'], 2)}",
                              3: "TPM_Logit"})
    _celdas(_fila(t, "Cuenta"), {2: f"α = {c(P.f('cta.alfa'), 3)}"})
    _celdas(_fila(t, "ETF"), {1: "t de Student (ν = 5) con volatilidad según el estado, mensual",
                              2: f"Media {pc(P.f('etf.mu'), 2)}; desviación {pc(P.f('etf.vol_N'), 1)} (expansión y normal) y "
                                 f"{pc(P.f('etf.vol_C'), 1)} (crisis)", 3: "ETF, MARKOV_ETF"})
    _celdas(_fila(t, "Estado de la economía"), {2: "Umbrales 1 % y 4 %", 3: "PIB, MARKOV"})
    _agregar_fila(t, ["Necesidades de liquidez", "Poisson con monto lognormal (supuesto, sección 12)",
                      "1 cada 24 meses (12 en crisis); mediana 3 % del capital; σ = 0,8", ""])

    # ---------------- referencias de los puntos 1 a 7, a la lista única
    k = _i(doc, "Referencias de las secciones 5.4, 7.6 y 7.7", True)
    k2 = _i(doc, "7.9 Resumen de los modelos elegidos", True)
    refs = [_texto(el) for el in _cuerpo(doc)[k + 1:k2] if _texto(el).strip()]
    _borrar(doc, k, k2)
    correcciones_auditoria(w, P, c, pc)
    return [r for r in refs_etf + refs if not r.startswith("Aptología")]


def _renumerar(doc, desde: str, hasta: str, seccion: str) -> None:
    """Pasa 'Tabla 5.' y 'Figura 3.' a 'Tabla 7.5-a.' y actualiza las menciones dentro de la sección."""
    i, j = _i(doc, desde, True), _i(doc, hasta, True)
    pars = [Paragraph(el, doc.element.body) for el in _cuerpo(doc)[i:j] if el.tag == qn("w:p")]
    mapa = {}
    cuenta = {"Tabla": 0, "Figura": 0}
    for p in pars:
        m = re.match(r"^(Tabla|Figura) (\d+|" + re.escape(seccion) + r"-[a-z])\. ", p.text)
        if m:
            tipo, n = m.group(1), m.group(2)
            letra = "abcdefghijklmnopqrstuvwxyz"[cuenta[tipo]]
            cuenta[tipo] += 1
            mapa[(tipo, n)] = f"{tipo} {seccion}-{letra}"
            _poner(p, mapa[(tipo, n)] + ". " + p.text[m.end():])
    for p in pars:
        nuevo = re.sub(r"\b(Tabla|Figura) (\d+)\b(?![.,]\d)", lambda m: mapa.get((m.group(1), m.group(2)), m.group(0)), p.text)
        if nuevo != p.text:
            _poner(p, nuevo)


def correcciones_auditoria(w, P, c, pc) -> None:
    """Cambios de la auditoría final (9 de noviembre de 2026) sobre el texto de los puntos 1 a 7."""
    doc = w.doc
    res = RAIZ / "outputs" / "resultados"
    val = pd.read_csv(res / "validacion.csv")
    cm = pd.read_csv(res / "capital_meta.csv")
    n = P.i("exp.replicas")

    # §4: lo que de verdad se corrió.
    t = _tabla(doc, "Parámetro | Valor base")
    caps, metas = sorted(cm.capital.unique()), sorted(cm.meta.unique())
    _celdas(_fila(t, "Capital inicial"), {2: f"${c(caps[0] / 1e6)} a ${c(caps[-1] / 1e6)} millones (sección 17)"})
    _celdas(_fila(t, "Horizonte"), {2: "Fijo en el experimento; se puede cambiar en la aplicación"})
    _celdas(_fila(t, "Meta financiera"), {2: f"${c(metas[0] / 1e6)} a ${c(metas[-1] / 1e6)} millones (sección 17)"})
    _celdas(_fila(t, "Liquidez mínima"), {2: "3 % y 10 % en la sensibilidad (sección 16.3)"})
    _celdas(_fila(t, "Réplicas por estrategia"), {1: c(n)})

    # §5.2: una sola cita para la retención del CDT (la misma del Excel).
    reemplazar(doc, "Decreto 1625 de 2016, art. 1.2.4.2.32 (tarifa ajustada por el Decreto 2418 de 2013), citado en el Concepto DIAN "
                    "3325 de 2024 [L1]",
               "4 % fijado por el Decreto 2418 de 2013, compilado en el Decreto 1625 de 2016 y mantenido por el Decreto 572 de 2025 "
               "(Ministerio de Hacienda y Crédito Público, 2013, 2016, 2025)")
    reemplazar(doc, " [L1]", "")

    # §5.4: no hay venta ordinaria en el modelo; el inmueble se valora neto al final.
    t = next(x for x in doc.tables if x.rows[0].cells[0].text.strip() == "Regla" and "Precio de compra" in
             " ".join(f.cells[0].text for f in x.rows))
    for f in t.rows:
        nombre = f.cells[0].text.strip()
        if nombre == "Venta ordinaria":
            _celdas(f, {0: "Valor al final del horizonte",
                        1: "Precio de mercado menos los costos de venta: comisión 3 % más IVA, notaría 0,27 % y retención en la "
                           "fuente de 1 % (4,84 % en total). El modelo no hace una venta ordinaria antes del final",
                        2: "Comisión publicada; Lonja de Medellín (2025): 225 días para vender vivienda usada"})
        elif nombre == "Venta urgente":
            _celdas(f, {1: "Se decide cuando el efectivo, el ETF y los CDT no alcanzan; el dinero llega entre 4 y 7 meses después "
                           "(uniforme). Descuento de 10 % (sensibilidad 5 % a 20 %) más los costos de venta",
                        2: "Plazo: Lonja de Medellín (2025), 225 días; descuento: supuesto"})
        elif nombre == "Precio de compra":
            _celdas(f, {1: "$300.000.000 en E2 y E4 (50 m², $6,0 M/m²). El valor es proporcional al capital: E3, que pone 55 %, "
                           "compra uno de $385 millones (unos 64 m²) con las mismas reglas"})
    reemplazar(doc, "Tarda de 4 a 7 meses en venderse", "Venta urgente: el dinero llega de 4 a 7 meses después")

    # §7.1: inflación mensual y normalidad según la prueba.
    reemplazar(doc, "• Cada mes los precios suben una doceava parte de esa inflación, y con eso se calcula el patrimonio real.",
               "• Entre un diciembre y el siguiente la inflación anual cambia en línea recta, y cada mes el índice de precios sube "
               "lo que corresponde a la inflación anual de ese momento. Con ese índice se calcula el patrimonio real.")
    t = _tabla(doc, "Pregunta | Resultado")
    _celdas(_fila(t, "¿La inflación sigue una distribución normal?"),
            {1: "Jarque-Bera (con asimetría y curtosis muestrales): p = 0,002; Shapiro-Wilk: p = 0,057; Lilliefors: p = 0,29 "
                "(2001 a 2025, n = 25)",
             2: "Depende de la prueba. Jarque-Bera rechaza la normal y las otras dos no, porque con 25 datos tienen poca potencia. "
                "La asimetría hacia arriba sí se ve (1,21), y por eso no usamos la normal."})
    _celdas(_fila(t, "¿Sigue una distribución lognormal?"), {1: "Jarque-Bera: p = 0,82; Shapiro-Wilk: p = 0,81"})

    # §7.2.7: el tope sale de la validación.
    tope = float(val[val.medida.str.startswith("Réplicas con la TPM en el tope")].simulado.iloc[0])
    med = float(val[val.medida.str.startswith("TPM mediana")].simulado.iloc[0])
    reemplazar(doc, "Con el logit la TPM toca el tope solo en 0,4 % de las réplicas y su mediana al año 10 es 7,25 %.",
               f"Con el logit la TPM termina en el tope de 15 % en {pc(tope, 1)} de las réplicas y su mediana al año 10 es "
               f"{c(med, 2)} %.")

    # §7.3: IC de α por entidad (las 19 tasas no son independientes).
    t = _tabla(doc, "Concepto | Valor")
    _celdas(_fila(t, "Intervalo de confianza"), {0: "Intervalo de confianza de 95 % por entidad (t con 3 grados de libertad)",
                                                1: "0,755 a 1,046"})
    reemplazar(doc, "así que no se rechaza que α sea constante.",
               "así que no se rechaza que α sea constante. Como las 19 tasas vienen de cuatro cuentas, el intervalo de α se calcula "
               "con el promedio de cada cuenta y no con las 19 tasas como si fueran independientes; por eso es ancho.")
    reemplazar(doc, "publicadas en El Tiempo, El Colombiano y Rankia.",
               "publicadas en El Tiempo, El Colombiano y Rankia.")

    # §7.4.10: el estado no mueve la TPM.
    reemplazar(doc, "La tasa depende del estado económico a través de la TPM: el estado fija los cambios de la TPM y estos el margen.",
               "En el modelo la tasa no depende del estado económico. Depende de la TPM, que la Junta mueve según la inflación y "
               "la brecha de Taylor (7.2.7), y el margen depende del cambio reciente de la TPM.")

    # §7.5: referencias cruzadas viejas, notas internas y muletillas.
    reemplazar(doc, " La Tabla 7 muestra de dónde sale.", " La explicación está en los subperiodos.")
    reemplazar(doc, "y lo retomamos en los puntos 9 y 10.", "y por eso la llevamos al análisis de sensibilidad (sección 16.3).")
    reemplazar(doc, "Conclusión del punto: tratar", "Tratar")
    reemplazar(doc, "Hay algo importante. Sin esos tres meses", "Sin esos tres meses")
    reemplazar(doc, "La diferencia no es menor. Para la normal,", "Para la normal,")
    reemplazar(doc, " Las Figuras 5 a 8 muestran dónde se diferencian.", "")
    reemplazar(doc, "que compensa con creces el sesgo", "menor que el sesgo")
    reemplazar(doc, " La anotamos como una observación para discutir con el resto del equipo.", "")
    reemplazar(doc, "como sugiere la Tabla 9", "como muestran los meses atípicos")
    reemplazar(doc, "Como la asimetría no es significativa (p = 0,14), ajustamos",
               "Aunque los datos son asimétricos, la razón de verosimilitud entre la t asimétrica y la t simétrica no es significativa "
               "(p = 0,14): el parámetro de asimetría no mejora el ajuste lo suficiente. Por eso ajustamos")
    t = _tabla(doc, "Estadístico | Valor | Estadístico | Valor")
    _celdas(_fila(t, "Jarque-Bera"), {1: "361,7 (< 0,001)"})
    t = _tabla(doc, "Prueba (nivel 5 %) | Estadístico")
    _celdas(_fila(t, "Varianza igual en los tres estados"),
            {0: "Varianza igual en los tres estados del PIB, umbral de 1 % (Levene centrado en la media; con la mediana p = 0,008)"})

    # §7.6.10: umbral del contraste por estados.
    reemplazar(doc, "6,9 % con PIB débil (menos de 1,5 %)", "6,9 % con PIB débil (menos de 1 %)")

    # §7.7.2: la tabla usa la duración del parámetro.
    t = _tabla(doc, "Estado | p | D (meses)")
    for f in t.rows[1:]:
        letra = {"Expansión": "E", "Normal": "N", "Crisis": "C"}.get(f.cells[0].text.strip())
        if letra:
            p_, d_ = P.f(f"vac.p_{letra}"), P.f(f"vac.media_{letra}")
            _celdas(f, {1: pc(p_, 0), 2: c(d_, 1), 3: pc(p_ * d_ / (12 + p_ * d_), 2)})
    pe, de = P.f("vac.p_E"), P.f("vac.media_E")
    reemplazar(doc, "El 0,62 % de expansión coincide con el 0,6 % que reporta Ciencuadras para Medellín en 2026.",
               f"El {pc(pe * de / (12 + pe * de), 2)} de expansión queda cerca del 0,6 % que reporta Ciencuadras (2026) para "
               "Medellín.")

    # ---------------- segunda auditoría
    sm = pd.read_csv(res / "sensibilidad_modelos.csv")
    g = sm[sm.modelo.str.startswith("Junta: logit")]
    dmax = float((g.alt_p_meta - g.ref_p_meta).abs().max() * 100)
    reemplazar(doc, "pero los resultados del simulador cambian menos de un punto (sección 16.3).",
               f"pero la probabilidad de meta cambia a lo sumo {c(dmax, 1)} puntos (sección 16.3).")
    # Sensibilidades: el texto dice solo lo que se corrió.
    reemplazar(doc, "Se analiza la sensibilidad sin memoria y con memoria de 60 %.",
               "En la sensibilidad (sección 16.3) se prueba sin memoria y con memoria de 60 %, con el mismo nivel de largo plazo.")
    reemplazar(doc, "se analiza la sensibilidad con un nivel normal de 3 %.",
               "en la sensibilidad (sección 16.3) el nivel normal baja a 3,5 % y sube a 6,5 %.")
    reemplazar(doc, "La sensibilidad duplica la desviación de ε y usa la alternativa M6, que agrega el nivel de la TPM (coeficiente −0,09; p = 0,014; R² sube a 0,35).",
               "En la sensibilidad movemos la constante del margen medio punto arriba y abajo y probamos el margen por fase como "
               "modelo alternativo. El modelo M6, que agrega el nivel de la TPM (coeficiente −0,09; p = 0,014; R² sube a 0,35), no "
               "lo usamos: agrega un parámetro y no lo evaluamos fuera de muestra.")
    reemplazar(doc, "pero son pocos episodios independientes; queda como sensibilidad.",
               "pero son pocos episodios independientes, así que el modelo no lo usa.")
    reemplazar(doc, "6 meses (crisis, con sensibilidad de 4 a 9).",
               "6 meses (crisis); en la sensibilidad las duraciones se multiplican por 0,5 y por 2.")
    # Redacción de 7.5.
    reemplazar(doc, "p-valores", "valores p")
    reemplazar(doc, "p-valor", "valor p")
    reemplazar(doc, " Todas las cifras que aparecen aquí salen de los archivos descargados y de los scripts que acompañan el documento. "
                    "Ninguna se fijó a mano.", "")
    reemplazar(doc, "Un cuidado de método: las tablas usuales", "Las tablas usuales")
    reemplazar(doc, "En diez años la diferencia no es pequeña: con la media histórica", "En diez años, con la media histórica")
    reemplazar(doc, "Juntando lo anterior, la dependencia", "La dependencia")
    reemplazar(doc, "Esta subsección completa el análisis con las pruebas de hipótesis que pide el enunciado, compara el modelo "
                    "elegido con una t de Student simple y revisa la dependencia",
               "Aquí se reúnen las pruebas de hipótesis, la comparación del modelo elegido con una t de Student "
               "simple y la dependencia")
    reemplazar(doc, "La lectura es la de los apartados anteriores. ", "")
    reemplazar(doc, "(marzo de 2020, junio de 2022 y uno más)", "(marzo de 2020, junio de 2022 y enero de 2026)")
    reemplazar(doc, "en el planteamiento del proyecto)", "en las reglas del ETF de la sección 5.3)")
    reemplazar(doc, "candidata que aparecía en el plan inicial del proyecto", "candidata que propusimos al empezar el estudio")

    # 7.9: por qué se eligió cada modelo (lo que hay que poder defender en la sustentación).
    t = _tabla(doc, "Proceso | Modelo para el simulador")
    razones = {
        "Inflación": "Siempre positiva, con saltos hacia arriba y memoria de un año a otro",
        "TPM": "Reproduce las decisiones observadas; la tabla lleva la TPM a un tope que nunca tuvo",
        "Cuenta": "α no depende del nivel de la TPM (p = 0,29)",
        "CDT": "El margen se abre cuando la TPM sube; residuos no normales, por eso se sortean de los datos",
        "ETF": "Colas pesadas, sin memoria, y el estado mueve su volatilidad (sección 8)",
        "Valoriz": "El IPVU mide vivienda usada; años seguidos parecidos (φ significativo); no depende del estado",
        "Vacancia": "La gamma reproduce la curva de Palomma; la probabilidad de no renovar es supuesto",
        "Estado de la economía": "El estado depende del trimestre anterior (χ² = 81,8; p < 0,001)",
        "Necesidades": "Sin datos: supuesto, por eso se muestran resultados con y sin ellas",
    }
    tbl = t._tbl
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None and len(grid) == len(t.rows[0].cells):
        grid.append(copy.deepcopy(grid[-1]))
    for k, fila in enumerate(t.rows):
        nueva = copy.deepcopy(fila.cells[-1]._tc)
        fila._tr.append(nueva)
        texto = "Por qué" if k == 0 else next((v for clave, v in razones.items() if fila.cells[0].text.strip().startswith(clave)), "")
        _poner(fila.cells[-1].paragraphs[0], texto or " ")

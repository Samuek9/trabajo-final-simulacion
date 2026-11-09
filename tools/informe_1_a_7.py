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
                                      2: "Estimado con 19 tasas de 4 cuentas (hoja TPM_Cuenta)"})
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
                                              2: "Estimado (hojas ETF y MARKOV_ETF, sección 7.5)"})
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
           "Hojas del libro: TPM_Cuenta y TPM_Logit. En 7.2.1 a 7.2.6 está el primer estudio, con las 46 reuniones de 2021 a 2026 y "
           "una tabla de decisiones según la tendencia de la inflación. En 7.2.7 está el modelo que usa el simulador, un logit "
           "ordenado estimado con 25 años de decisiones, y por qué cambiamos. La TPM oficial diaria desde 1998 (SUAMECA) se usa "
           "también en 7.4 para estimar el margen del CDT.")
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
    _poner(_parrafo(doc, "Hoja del libro: ETF. El código"),
           "Hoja del libro: ETF. Los pasos del enunciado están en 7.5.1 (fuente), 7.5.2 a 7.5.4 (exploración, análisis temporal y "
           "atípicos), 7.5.5 a 7.5.8 (modelos, estimación, bondad de ajuste y comparación), 7.5.9 (selección), 7.5.10 (dependencia) "
           "y 7.5.11 (pruebas de hipótesis y modelo que usa el simulador). Las limitaciones están al final de 7.5.11.")
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
           "El enunciado no permite suponer que todas las variables son independientes. Para saber con qué se mueve el ICOLCAP "
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
           "tomamos igual, porque es lo que se espera de una crisis y el enunciado pide que los estados muevan la renta variable, "
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
            "económico (agosto de 2011 a junio de 2026). La razón para dejar la Jones-Faddy como alternativa es el punto 8 del "
            "enunciado: la Jones-Faddy no tiene estados, y la t con volatilidad por estado sí conecta el ETF con la economía. El "
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
        val = pd.read_csv(RAIZ / "outputs" / "resultados" / "validacion.csv")
        vac = float(val[val.proceso == "Vacancia"].simulado.iloc[0])
        w.p(f"Validación. Con estos valores el apartamento pasa {pc(vac, 1)} del tiempo desocupado en la simulación, más que el "
            "0,6 % que reporta Ciencuadras para Medellín. Lo dejamos así a propósito: preferimos un modelo que exagere la vacancia "
            "a uno que la ignore. La Lonja reporta 121 días en oferta antes de arrendar, bastante más que los 46 días de Palomma; "
            "nos quedamos con Palomma porque mide el segmento y la ciudad del inmueble, y probamos la media de 4 meses de la Lonja "
            "en la sensibilidad (sección 16.3).")
    _mover_antes(doc, _i(doc, "7.7.11 Limitaciones", True), vacancia)

    # ---------------- 7.8 Estado de la economía
    _poner(_parrafo(doc, "El enunciado pide analizar la dependencia entre procesos y habla de estados"),
           "El enunciado pide analizar la dependencia entre procesos y habla de estados de la economía. Para eso definimos tres "
           "estados con el crecimiento anual del PIB real trimestral (DANE, base 2015, encadenado con las bases 2005 y 1994): crisis "
           "si crece menos de 1 % frente al mismo trimestre del año anterior, expansión si crece 4 % o más y normal en otro caso. "
           "Con 1 % entran como crisis la recesión de 2008 y 2009, la pandemia y el estancamiento de 2023 y 2024; con 0 % solo "
           "quedarían cuatro trimestres desde 2006. La hoja PIB trae los 126 trimestres, de 1995 a junio de 2026, y los umbrales se "
           "pueden cambiar.")
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
                      "1 cada 24 meses (12 en crisis); mediana $20 M; σ = 0,8", "PARAMETROS"])

    # ---------------- referencias de los puntos 1 a 7, a la lista única
    k = _i(doc, "Referencias de las secciones 5.4, 7.6 y 7.7", True)
    k2 = _i(doc, "7.9 Resumen de los modelos elegidos", True)
    refs = [_texto(el) for el in _cuerpo(doc)[k + 1:k2] if _texto(el).strip()]
    _borrar(doc, k, k2)
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

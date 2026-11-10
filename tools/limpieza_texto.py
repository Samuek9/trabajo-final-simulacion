"""Deja el texto del informe como un informe: sin referencias al enunciado, a las hojas del Excel ni al equipo.

Se aplica sobre el Word ya armado, antes del formato APA. Los textos de las secciones 1 a 7 vienen del Word del equipo
(que no se modifica) y algunos hablan del proceso de elaboración; aquí se reescriben o se quitan.
"""
from __future__ import annotations

from docx.oxml.ns import qn
from docx.table import Table

from informe_1_a_7 import _borrar, _cuerpo, _i, _parrafo, _poner, _texto, reemplazar

REESCRIBIR = {
    "Un inversionista colombiano tiene $700 millones":
        "Una persona quiere diseñar una estrategia de inversión para un horizonte que ella define. Debe decidir cómo repartir su "
        "capital entre una cuenta remunerada, CDT, un ETF de acciones colombianas y un apartamento para arrendar. Durante ese "
        "tiempo van a cambiar la inflación, las tasas del Banco de la República, el precio de las acciones y el de la vivienda. "
        "El apartamento puede quedarse vacío, y en cualquier momento puede aparecer un gasto grande que obligue a vender algo.",
    "Estos valores se cambian desde la hoja PARAMETROS":
        "Estos valores se cambian desde el Excel o desde la aplicación. Los de referencia son los que se usaron para comparar "
        "estrategias; los rangos son los que se recorrieron en el experimento de capital y meta (sección 17).",
    "Ningún valor de esta tabla es fijo.":
        "Ningún valor de esta tabla es fijo. El capital, el horizonte, la meta, la liquidez mínima, las estrategias y las "
        "necesidades de liquidez se cambian desde el Excel o desde la aplicación, y el simulador responde sin reconstruirse. Los "
        "valores de referencia son los de la configuración con la que se corrieron los experimentos de este informe. Los procesos "
        "aleatorios (inflación, tasas, ETF, inmueble, vacancia y estados de la economía) se estimaron con datos y quedan fijos. "
        "Las cifras de mercado (precios, tasas, cánones) cambian con el tiempo, y por eso cada una dice su fuente y su fecha de "
        "consulta: para actualizarla basta volver a la fuente.",
    "La meta se mide en pesos de hoy para que la inflación no la cumpla por sí sola":
        "La meta se mide en pesos de hoy para que la inflación no la cumpla por sí sola: con inflación promedio de 4,5 %, un "
        "patrimonio nominal de $155 millones dentro de diez años vale apenas $100 millones de hoy.",
    "Ejemplo: si el patrimonio pasa de $700 millones":
        "Ejemplo: si el patrimonio pasa de $100 millones a $150 millones en 10 años con una inflación promedio de 4,5 %, su valor "
        "real es de $96,6 millones de hoy: en pesos ganó $50 millones, pero en poder de compra perdió $3,4 millones.",
    "El enunciado pide como mínimo cinco fuentes":
        "El modelo incorpora siete fuentes de incertidumbre. Cinco son las que todo estudio de este tipo debe cubrir: rendimiento "
        "de la renta variable, inflación, tasas de interés, valorización de la propiedad raíz y vacancia. A ellas se suman dos que "
        "los datos justifican: el margen del CDT sobre la TPM y las necesidades extraordinarias de liquidez. Además se define el "
        "estado de la economía con el PIB (sección 7.8), que sirve para analizar la dependencia entre procesos.",
    "Liquidez extraordinaria. Es el gasto imprevisto":
        "Liquidez extraordinaria. Es el gasto imprevisto que obliga a vender activos. Se modela en la sección 12. Otras variables "
        "(canon, mantenimiento, tasa de cambio) entran como reglas o como variables de dependencia, no como procesos aleatorios "
        "propios: el canon se reajusta con la inflación simulada, el mantenimiento es un costo fijo con sensibilidad, y la tasa de "
        "cambio se prueba como variable explicativa del ETF en 7.5.",
    "Para cada proceso seguimos los 11 pasos del enunciado":
        "El estudio de cada proceso sigue el mismo orden: (1) fuente de datos, (2) exploración, (3) análisis temporal, (4) valores "
        "atípicos, (5) modelos candidatos, (6) estimación, (7) bondad de ajuste, (8) comparación, (9) selección, (10) dependencia "
        "con otras variables y (11) limitaciones. Las secciones 7.4, 7.6 y 7.7 llevan los pasos numerados igual; en las demás se "
        "recorren en el mismo orden. Donde un paso no aplica, se explica por qué. Todas las pruebas usan un nivel de significancia "
        "de 5 %.",
}

QUITAR = [
    "Los datos, cálculos, fórmulas y pruebas están en la hoja INF del libro.",
    " Los datos, cálculos, fórmulas y pruebas están en la hoja INF del libro.",
    " Los cálculos están en la hoja ETF.",
    " Los cálculos están en la hoja ETF",
    " Hoja: CDT.",
    " Hoja: INM.",
]


def _borrar_parrafos_que_empiezan(doc, inicio: str) -> int:
    n = 0
    for el in list(_cuerpo(doc)):
        if el.tag == qn("w:p") and _texto(el).strip().startswith(inicio):
            el.getparent().remove(el)
            n += 1
    return n


def _quitar_tabla_de_pasos(doc) -> None:
    """Quita la tabla que dice dónde está cada uno de los 11 pasos (y los rótulos que ya tenga)."""
    for t in doc.tables:
        if t.rows[0].cells[0].text.strip() == "Paso" and t.rows[1].cells[0].text.strip().startswith("1. Fuente de datos"):
            el = t._tbl
            sig = el.getnext()
            prev = el.getprevious()
            while prev is not None and prev.tag == qn("w:p") and (
                    _texto(prev).strip().startswith("Tabla ") or _texto(prev).strip().startswith("Dónde está cada uno")):
                p2 = prev.getprevious()
                prev.getparent().remove(prev)
                prev = p2
            if sig is not None and sig.tag == qn("w:p") and not _texto(sig).strip():
                sig.getparent().remove(sig)
            el.getparent().remove(el)
            return
    raise KeyError("tabla de los 11 pasos")


def _quitar_columna(t: Table, j: int) -> None:
    for fila in t.rows:
        celdas = fila._tr.findall(qn("w:tc"))
        fila._tr.remove(celdas[j])
    grid = t._tbl.find(qn("w:tblGrid"))
    if grid is not None:
        cols = grid.findall(qn("w:gridCol"))
        if len(cols) > j:
            grid.remove(cols[j])


def aplicar(doc) -> None:
    _quitar_tabla_de_pasos(doc)
    _borrar_parrafos_que_empiezan(doc, "Hoja del libro:")
    for inicio, texto in REESCRIBIR.items():
        _poner(_parrafo(doc, inicio), texto)
    for frase in QUITAR:
        reemplazar(doc, frase, "")
    reemplazar(doc, ". Hoja: INM (Figura 5).", " (Figura 5).")
    reemplazar(doc, "Definido por el equipo", "Parámetro de entrada")
    reemplazar(doc, "Se justifica en el punto 14", "Se justifica en la sección 14")
    reemplazar(doc, "Punto 13", "Sección 13")
    reemplazar(doc, "Documento del equipo, sección 12", "Sección 12")
    reemplazar(doc, "Decisión del equipo, contrastada con", "Decisión de modelación, contrastada con")
    reemplazar(doc, "Decisión del equipo", "Decisión de modelación")
    for t in doc.tables:
        enc = [c.text.strip() for c in t.rows[0].cells]
        if enc[:2] == ["Parámetro", "Valor base"]:
            _poner(t.rows[0].cells[1].paragraphs[0], "Valor de referencia")
        if len(enc) == 5 and enc[0] == "Proceso" and enc[3] == "Hoja":
            _quitar_columna(t, 3)

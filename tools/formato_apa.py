"""Formato APA 7 del informe: citas, lista de referencias, tablas, figuras, paginación y fuente.

Lo llama construir_informe.py al final, sobre el documento ya armado:
  - citas(doc): agrega (Autor, año) donde el texto usa una fuente.
  - referencias(w): escribe la lista de referencias.
  - aplicar(doc): numera tablas y figuras 1..n con número en negrita y título en cursiva arriba,
    pasa las aclaraciones a "Nota.", actualiza las menciones en el texto, repite el encabezado de
    las tablas, pone número de página y deja todo en Arial.
  - revisar(doc): lista citas sin referencia y referencias sin cita.
"""
from __future__ import annotations

import copy
import re

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph

# ---------------------------------------------------------------- referencias
# (clave de cita, referencia). La clave es lo que aparece en el texto, sin paréntesis.
REFERENCIAS = [
    ("Banco de la República, 2026a", "Banco de la República. (2026a). *Comunicados de las decisiones de política monetaria de la Junta Directiva, 2001 a 2026*. https://www.banrep.gov.co"),
    ("Banco de la República, 2026b", "Banco de la República. (2026b). *Sistema de consulta de estadísticas económicas (SUAMECA): tasa de política monetaria diaria, tasa de captación de CDT a 360 días, índice de precios de la vivienda usada (IPVU) e índice de precios de la vivienda nueva (IPVNBR)* [Conjunto de datos]. https://suameca.banrep.gov.co/estadisticas-economicas/"),
    ("Banks et al., 2010", "Banks, J., Carson, J. S., Nelson, B. L. y Nicol, D. M. (2010). *Discrete-event system simulation* (5.ª ed.). Pearson."),
    ("BlackRock Colombia, 2026", "BlackRock Colombia. (2026). *Fondo Bursátil iShares MSCI COLCAP: ficha del producto y archivo de datos del fondo* [Conjunto de datos]. https://www.blackrock.com/co"),
    ("Brant, 1990", "Brant, R. (1990). Assessing proportionality in the proportional odds model for ordinal logistic regression. *Biometrics, 46*(4), 1171-1178. https://doi.org/10.2307/2532457"),
    ("Burnham y Anderson, 2002", "Burnham, K. P. y Anderson, D. R. (2002). *Model selection and multimodel inference: A practical information-theoretic approach* (2.ª ed.). Springer. https://doi.org/10.1007/b97636"),
    ("Ciencuadras, 2025", "Ciencuadras. (2025). *I Informe Trimestral 2025: comportamiento de los inmuebles en Colombia*. https://www.ciencuadras.com/blog/wp-content/uploads/2025/05/I-Informe-Trimestral-2025.pdf"),
    ("Ciencuadras, 2026", "Ciencuadras. (2026). *Sube el arriendo en Colombia por escasez de vivienda usada*. https://www.ciencuadras.com/blog/guia-para-arrendar-inmuebles/sube-el-arriendo-en-colombia-por-escasez-de-vivienda-usada"),
    ("Concejo de Medellín, 2023", "Concejo de Medellín. (2023). *Acuerdo 093 de 2023, por medio del cual se expide el Estatuto Tributario del Distrito de Medellín*. https://www.medellin.gov.co/normograma/docs/astrea/docs/A_CONMED_0093_2023.htm"),
    ("Congreso de Colombia, 2003", "Congreso de Colombia. (2003, 10 de julio). *Ley 820 de 2003, por la cual se expide el régimen de arrendamiento de vivienda urbana*. http://www.secretariasenado.gov.co/senado/basedoc/ley_0820_2003.html"),
    ("Cont, 2001", "Cont, R. (2001). Empirical properties of asset returns: Stylized facts and statistical issues. *Quantitative Finance, 1*(2), 223-236. https://doi.org/10.1080/713665670"),
    ("CuantoMeCuesta, s. f.", "CuantoMeCuesta. (s. f.). *Gastos de comprar vivienda en Colombia*. https://cuantomecuesta.com/co/comprar-vivienda/"),
    ("DANE, 2019", "Departamento Administrativo Nacional de Estadística [DANE]. (2019). *Encuesta de Carga Financiera y Educación Financiera de los Hogares (IEFIC) 2017 y 2018* [Microdatos, catálogo 626]. Fuente: Departamento Administrativo Nacional de Estadística: www.dane.gov.co"),
    ("DANE, 2024", "Departamento Administrativo Nacional de Estadística [DANE]. (2024). *Encuesta Nacional de Calidad de Vida (ECV) 2023*. https://www.dane.gov.co/index.php/estadisticas-por-tema/salud/calidad-de-vida-ecv"),
    ("DANE, 2026a", "Departamento Administrativo Nacional de Estadística [DANE]. (2026a). *Índice de Precios al Consumidor: series históricas y anexos* [Conjunto de datos]. https://www.dane.gov.co/index.php/estadisticas-por-tema/precios-y-costos/indice-de-precios-al-consumidor-ipc"),
    ("DANE, 2026b", "Departamento Administrativo Nacional de Estadística [DANE]. (2026b). *Producto interno bruto trimestral a precios constantes* [Conjunto de datos]. Consultado en SUAMECA, https://suameca.banrep.gov.co/estadisticas-economicas/"),
    ("DIAN, 2025", "Dirección de Impuestos y Aduanas Nacionales [DIAN]. (2025, 15 de diciembre). *Resolución 000238 de 2025, por la cual se fija el valor de la Unidad de Valor Tributario (UVT) para 2026*. https://normograma.dian.gov.co/dian/compilacion/docs/resolucion_dian_0238_2025.htm"),
    ("El País, 2026", "El País. (2026, septiembre). *CDT en Colombia: estas son las tasas que ofrecen algunas entidades para septiembre de 2026*. https://www.elpais.com.co/economia/cdt-en-colombia-estas-son-las-tasas-que-ofrecen-algunas-entidades-para-septiembre-de-2026-3049.html"),
    ("Equipo del trabajo, 2026a", "Equipo del trabajo. (2026a). *Estrategias de inversión; Escenarios; Dependencia entre variables y estados económicos; Diseño experimental; Diagrama de eventos y diagrama de flujo; Análisis de estados económicos (Markov)* [Documentos internos no publicados]."),
    ("Equipo del trabajo, 2026b", "Equipo del trabajo. (2026b). *Tasas de cuentas de ahorro digitales en Colombia, diciembre de 2023 a octubre de 2026* [Conjunto de datos no publicado, recopilado de notas de El Tiempo, El Colombiano y Rankia; hoja TPM_Cuenta del Excel maestro]."),
    ("Fogafín, 2025", "Fondo de Garantías de Instituciones Financieras [Fogafín]. (2025). *ABC del proceso de pago del seguro de depósitos*. https://www.fogafin.gov.co/sites/default/files/2025-12/ABC%20del%20proceso%20de%20pago%20del%20seguro%20de%20dep%C3%B3sitos.pdf"),
    ("Hamilton, 1989", "Hamilton, J. D. (1989). A new approach to the economic analysis of nonstationary time series and the business cycle. *Econometrica, 57*(2), 357-384. https://doi.org/10.2307/1912559"),
    ("Harris et al., 2020", "Harris, C. R., Millman, K. J., van der Walt, S. J., Gommers, R., Virtanen, P., Cournapeau, D., Wieser, E., Taylor, J., Berg, S., Smith, N. J., Kern, R., Picus, M., Hoyer, S., van Kerkwijk, M. H., Brett, M., Haldane, A., Fernández del Río, J., Wiebe, M., Peterson, P., . . . Oliphant, T. E. (2020). Array programming with NumPy. *Nature, 585*(7825), 357-362. https://doi.org/10.1038/s41586-020-2649-2"),
    ("HaceCuentas, 2026", "HaceCuentas. (2026). *Calculadora de gastos notariales y de registro de una compraventa en 2026*. https://hacecuentas.com/co/calculadora-gastos-notariales-registro-compraventa-2026"),
    ("Jones y Faddy, 2003", "Jones, M. C. y Faddy, M. J. (2003). A skew extension of the t-distribution, with applications. *Journal of the Royal Statistical Society: Series B, 65*(1), 159-174. https://doi.org/10.1111/1467-9868.00378"),
    ("Kass y Raftery, 1995", "Kass, R. E. y Raftery, A. E. (1995). Bayes factors. *Journal of the American Statistical Association, 90*(430), 773-795. https://doi.org/10.1080/01621459.1995.10476572"),
    ("La Lonja de Propiedad Raíz de Medellín y Antioquia, 2025", "La Lonja de Propiedad Raíz de Medellín y Antioquia. (2025). *Mercado inmobiliario residencial* (Edición 275). https://lonja.org.co/wp-content/uploads/documentos-afiliados/275-Mercado-inmobiliario-residencial.pdf"),
    ("Law, 2015", "Law, A. M. (2015). *Simulation modeling and analysis* (5.ª ed.). McGraw-Hill."),
    ("Lilliefors, 1967", "Lilliefors, H. W. (1967). On the Kolmogorov-Smirnov test for normality with mean and variance unknown. *Journal of the American Statistical Association, 62*(318), 399-402. https://doi.org/10.1080/01621459.1967.10482916"),
    ("McCullagh, 1980", "McCullagh, P. (1980). Regression models for ordinal data. *Journal of the Royal Statistical Society: Series B, 42*(2), 109-142. https://doi.org/10.1111/j.2517-6161.1980.tb01109.x"),
    ("Metrocuadrado, 2024", "Metrocuadrado. (2024). *Gastos notariales de venta y compra del inmueble en 2024*. https://www.metrocuadrado.com/noticias/guia-de-compra/gastos-notariales-de-venta-y-compra-del-inmueble-en-2024-4709/"),
    ("Ministerio de Hacienda y Crédito Público, 2013", "Ministerio de Hacienda y Crédito Público. (2013). *Decreto 2418 de 2013, por el cual se reglamenta parcialmente el Estatuto Tributario*. https://normograma.dian.gov.co/dian/compilacion/docs/decreto_2418_2013.htm"),
    ("Ministerio de Hacienda y Crédito Público, 2016", "Ministerio de Hacienda y Crédito Público. (2016). *Decreto 1625 de 2016, Decreto Único Reglamentario en materia tributaria*. https://normograma.dian.gov.co/dian/compilacion/docs/decreto_1625_2016.htm"),
    ("Ministerio de Hacienda y Crédito Público, 2025", "Ministerio de Hacienda y Crédito Público. (2025, 28 de mayo). *Decreto 0572 de 2025, por el cual se modifican las tarifas de retención en la fuente*. https://normograma.dian.gov.co/dian/compilacion/docs/decreto_0572_2025.htm"),
    ("MSCI, 2023", "MSCI. (2023). *MSCI COLCAP Index methodology*. MSCI Inc."),
    ("Noticias Caracol, 2026", "Noticias Caracol. (2026, enero). *Esto le paga Davivienda por invertir 1 millón en un CDT en enero de 2026*. https://www.noticiascaracol.com/economia/esto-le-paga-davivienda-por-invertir-1-millon-en-un-cdt-en-enero-de-2026-nuevas-tasas-so35"),
    ("OneEstimate, 2026", "OneEstimate. (2026). *Precios de pintura y acabados por m² en Colombia 2026*. https://oneestimate.ai/es/blog/precio-m2-pintura-acabados-colombia-2026"),
    ("Palomma, 2026", "Palomma. (2026). *Observatorio del Arriendo: tiempos de arriendo, canon por tamaño y precios de venta en Medellín*. https://palomma.com/blog"),
    ("Seabold y Perktold, 2010", "Seabold, S. y Perktold, J. (2010). Statsmodels: Econometric and statistical modeling with Python. En *Proceedings of the 9th Python in Science Conference* (pp. 92-96). https://doi.org/10.25080/Majora-92bf1922-011"),
    ("Stephens, 1974", "Stephens, M. A. (1974). EDF statistics for goodness of fit and some comparisons. *Journal of the American Statistical Association, 69*(347), 730-737. https://doi.org/10.1080/01621459.1974.10480196"),
    ("Taylor, 1993", "Taylor, J. B. (1993). Discretion versus policy rules in practice. *Carnegie-Rochester Conference Series on Public Policy, 39*, 195-214. https://doi.org/10.1016/0167-2231(93)90009-L"),
    ("Virtanen et al., 2020", "Virtanen, P., Gommers, R., Oliphant, T. E., Haberland, M., Reddy, T., Cournapeau, D., Burovski, E., Peterson, P., Weckesser, W., Bright, J., van der Walt, S. J., Brett, M., Wilson, J., Millman, K. J., Mayorov, N., Nelson, A. R. J., Jones, E., Kern, R., Larson, E., . . . SciPy 1.0 Contributors. (2020). SciPy 1.0: Fundamental algorithms for scientific computing in Python. *Nature Methods, 17*(3), 261-272. https://doi.org/10.1038/s41592-019-0686-2"),
    ("Wilson, 1927", "Wilson, E. B. (1927). Probable inference, the law of succession, and statistical inference. *Journal of the American Statistical Association, 22*(158), 209-212. https://doi.org/10.1080/01621459.1927.10502953"),
    ("Yahoo Finance, 2026", "Yahoo Finance. (2026). *ICOLCAP.CL: datos históricos* [Conjunto de datos]. https://finance.yahoo.com"),
]

# (texto que ya está en el informe, texto con la cita). Cada par se aplica en todo el documento.
CITAS = [
    ("Para el canon usamos el boletín de mercado inmobiliario residencial de la Lonja de Propiedad Raíz de Medellín y Antioquia (edición 275, agosto de 2025).",
     "Para el canon usamos el boletín de mercado inmobiliario residencial de la Lonja de Propiedad Raíz de Medellín y Antioquia (2025), edición 275 de agosto de 2025."),
    ("de la tarifa por mil del Acuerdo 093 de 2023 para el estrato.", "de la tarifa por mil del Acuerdo 093 de 2023 para el estrato (Concejo de Medellín, 2023)."),
    ("(Ley 820 de 2003, art. 20).", "(Ley 820 de 2003, art. 20; Congreso de Colombia, 2003)."),
    ("Ciencuadras reporta un canon nacional", "Ciencuadras (2025) reporta un canon nacional"),
    ("una regla de Taylor, TPM", "una regla de Taylor (1993), TPM"),
    ("una prueba de Brant aproximada", "una prueba de Brant (1990) aproximada"),
    ("con la lógica de Lilliefors:", "con la lógica de Lilliefors (1967):"),
    ("las tablas usuales de KS y AD suponen", "las tablas usuales de KS y AD (Stephens, 1974) suponen"),
    ("en la escala de Kass y Raftery", "en la escala de Kass y Raftery (1995)"),
    ("es la t asimétrica de Jones y Faddy con", "es la t asimétrica de Jones y Faddy (2003) con"),
    ("las ponen unos pocos meses de choque.", "las ponen unos pocos meses de choque, algo habitual en los rendimientos de activos (Cont, 2001)."),
    ("del Observatorio del Arriendo de Palomma (2025 a 2026).", "del Observatorio del Arriendo de Palomma (2026), con datos de 2025 y 2026."),
    ("más que el 0,6 % que reporta Ciencuadras para Medellín.", "más que el 0,6 % que reporta Ciencuadras (2026) para Medellín."),
    ("frente a 0,6 % en Ciencuadras.", "frente a 0,6 % en Ciencuadras (2026)."),
    ("intervalo de Wilson, que no se sale de 0 y 1.", "intervalo de Wilson (1927), que no se sale de 0 y 1."),
    ("no se modela el riesgo de impago ni el tope de Fogafín.", "no se modela el riesgo de impago ni el tope del seguro de depósitos (Fogafín, 2025)."),
    ("Fogafín cubre hasta $50 millones por persona y por entidad.", "Fogafín cubre hasta $50 millones por persona y por entidad (Fogafín, 2025)."),
    ("Verificado (CuantoMeCuesta, HaceCuentas y Metrocuadrado, 2026)", "Verificado (CuantoMeCuesta, s. f.; HaceCuentas, 2026; Metrocuadrado, 2024)"),
    ("en la serie de Yahoo", "en la serie de Yahoo Finance (2026)"),
    ("Noticias Caracol (enero de 2026) y El País (septiembre de 2026)", "Noticias Caracol (2026) y El País (2026)"),
    ("DANE, Encuesta de Calidad de Vida 2023", "Encuesta Nacional de Calidad de Vida 2023 (DANE, 2024)"),
    ("la exención mensual de 18 millones (350 UVT)", "la exención mensual de 18 millones (350 UVT; DIAN, 2025)"),
    ("el índice MSCI COLCAP. Cotiza", "el índice MSCI COLCAP (MSCI, 2023). Cotiza"),
    ("Se constituyó el 6 de julio de 2011,", "Se constituyó el 6 de julio de 2011 (BlackRock Colombia, 2026),"),
    ("peso de Akaike 0,97", "peso de Akaike 0,97; Burnham y Anderson, 2002"),
    ("La Encuesta de Carga Financiera y Educación Financiera de los Hogares (IEFIC) del DANE, de 2017 y 2018,",
     "La Encuesta de Carga Financiera y Educación Financiera de los Hogares (IEFIC) del DANE (2019), de 2017 y 2018,"),
    ("202 decisiones, de las cuales", "202 decisiones (Banco de la República, 2026a), de las cuales"),
    ("La hoja PIB trae los 126 trimestres,", "La hoja PIB trae los 126 trimestres (DANE, 2026b),"),
    ("Modelo. Un logit ordenado con dos variables:", "Modelo. Un logit ordenado (McCullagh, 1980) con dos variables:"),
    ("La segunda es una cadena de Markov de tres estados económicos", "La segunda es una cadena de Markov de tres estados económicos, en la línea de los modelos de cambio de régimen de Hamilton (1989),"),
    ("El simulador está escrito en Python y no usa ninguna librería de simulación:",
     "El simulador está escrito en Python, con NumPy, SciPy y statsmodels para el cálculo y la estadística (Harris et al., 2020; Virtanen et al., 2020; Seabold y Perktold, 2010), y no usa ninguna librería de simulación:"),
    ("El diagrama de eventos y el diagrama de flujo los construyó el equipo antes de programar.",
     "El diagrama de eventos y el diagrama de flujo los construyó el equipo antes de programar (Equipo del trabajo, 2026a)."),
    ("Los promedios ponderados de captación de la Superintendencia Financiera dan un α más bajo, cercano a 0,82, porque incluyen cuentas tradicionales que pagan menos;",
     "Las cuentas de ahorro tradicionales pagan bastante menos que las digitales, así que un promedio de todo el sistema daría un α más bajo;"),
]


def _parrafos(doc):
    yield from doc.paragraphs
    for t in doc.tables:
        for f in t.rows:
            for c in f.cells:
                yield from c.paragraphs


def _poner(p: Paragraph, texto: str) -> None:
    runs = p.runs
    runs[0].text = texto
    for r in runs[1:]:
        r._r.getparent().remove(r._r)


def reemplazar(doc, viejo: str, nuevo: str) -> int:
    n = 0
    for p in _parrafos(doc):
        if viejo not in p.text:
            continue
        hecho = False
        for r in p.runs:
            if viejo in r.text:
                r.text = r.text.replace(viejo, nuevo)
                hecho = True
        if not hecho:
            _poner(p, p.text.replace(viejo, nuevo))
        n += 1
    return n


def citas(doc) -> list[str]:
    """Agrega las citas; devuelve los textos que no se encontraron (para revisar a mano)."""
    return [a[:70] for a, b in CITAS if reemplazar(doc, a, b) == 0]


def _cursiva_markdown(p: Paragraph, texto: str) -> None:
    """Escribe texto con *cursiva* en el párrafo."""
    for k, trozo in enumerate(texto.split("*")):
        if trozo:
            r = p.add_run(trozo)
            r.italic = bool(k % 2) or None


def referencias(w) -> None:
    w.h("Referencias")
    for _, ref in sorted(REFERENCIAS, key=lambda x: x[1].lower().replace("*", "")):
        p = w.doc.add_paragraph()
        _cursiva_markdown(p, ref)
        f = p.paragraph_format
        f.left_indent, f.first_line_indent = Pt(36), Pt(-36)   # sangría francesa de 1,27 cm


# ---------------------------------------------------------------- tablas y figuras
_PADRE = [None]  # cuerpo del documento, para crear Paragraph sueltos
LEYENDA = re.compile(r"^(Tabla|Figura) ([0-9A-C][0-9.]*-[a-z])\.\s*(.*)$", re.S)

# Títulos de las tablas del equipo que no tenían, por las dos primeras celdas del encabezado.
TITULOS = {
    "Parámetro|Valor base": "Parámetros configurables del inversionista",
    "Activo|Instrumento": "Universo de inversión",
    "Regla|Valor en el modelo|Tipo y fuente|Disponibilidad": "Reglas de la cuenta remunerada",
    "Regla|Valor en el modelo|Sustento": "Reglas del CDT",
    "Regla|Valor en el modelo|Tipo y fuente|ETF": "Reglas del ETF",
    "Regla|Valor en el modelo|Tipo y fuente|Inmueble": "Reglas del inmueble",
    "Avalúo / valor comercial|Avalúo": "Predial anual según el avalúo y el estrato",
    "Año|Arriendo efectivo (%)": "Reajuste de los arriendos frente al IPC, 2010 a 2025",
    "Fuente de incertidumbre|Qué cambia en el simulador": "Fuentes de incertidumbre del modelo",
    "Paso|Inflación": "Dónde está cada uno de los 11 pasos del estudio de cada proceso",
    "Proceso|Depende de": "Dependencias entre los procesos aleatorios",
    "Camino|Qué hace en el simulador": "Caminos por los que la inflación afecta al inversionista",
    "Período|Inflación a diciembre": "Inflación de diciembre por periodos",
    "Pregunta|Resultado": "Pruebas de hipótesis sobre la inflación anual",
    "Si el año pasado la inflación fue…": "Inflación esperada según la del año anterior",
    "Prueba|Resultado": "Validación del modelo de inflación",
    "Cambio en la reunión": "Cambios de la TPM por reunión de la Junta",
    "Tendencia de la inflación|P(baja)": "Probabilidades de la tabla de la Junta por tendencia de la inflación (alternativa)",
    "Concepto|Valor": "Estimación de α, la fracción de la TPM que paga la cuenta",
    "Fase de la TPM|Semanas": "Margen del CDT según la fase de la TPM",
    "Modelo|Especificación": "Modelos candidatos para el margen del CDT",
    "Prueba (nivel 5 %)|Estadístico": "Pruebas de hipótesis sobre el rendimiento mensual del ETF",
    "Modelo|Media mensual": "Qué reproduce cada modelo del ETF",
    "Serie|n": "Estadísticos de la variación anual del IPVU",
    "Modelo|Parámetros|AIC": "Modelos candidatos para la valorización",
    "Fuente|Qué aporta": "Fuentes consultadas para la vacancia",
    "Estado|p": "Vacancia implícita por estado de la economía",
    "Modelo|Parámetros|Log-verosimilitud por obs.": "Ajuste de la duración de la vacancia por máxima verosimilitud",
    "Estado|Media (meses)": "Distribución gamma de la duración de la vacancia por estado",
    "Proceso|Modelo para el simulador": "Resumen de los modelos elegidos",
}
NOTAS_TABLA = {
    "Pruebas de hipótesis sobre la inflación anual": "Datos del DANE (2026a). Nivel de significancia de 5 %.",
    "Reajuste de los arriendos frente al IPC, 2010 a 2025": "Datos del DANE (2026a).",
    "Margen del CDT según la fase de la TPM": "Datos semanales del Banco de la República (2026b), 2008 a 2026.",
    "Predial anual según el avalúo y el estrato": "Tarifas del Concejo de Medellín (2023).",
}
NOTA_RESULTADOS = "Elaboración propia con los resultados de la simulación (outputs/resultados)."


def _texto(el) -> str:
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def _tiene_imagen(el) -> bool:
    return el.tag == qn("w:p") and bool(list(el.iter(qn("w:drawing"))))


def _clave(tbl) -> list[str]:
    fila = tbl.find(qn("w:tr"))
    return [_texto(c).strip() for c in fila.findall(qn("w:tc"))]


def _titulo_tabla(tbl, previo: str) -> str | None:
    cab = _clave(tbl)
    if cab[:3] == ["Regla", "Valor en el modelo", "Tipo y fuente"]:
        todo = _texto(tbl)
        if "Precio de compra" in todo:
            return "Reglas del inmueble"
        if "Comisión de ad" in todo:
            return "Reglas del ETF"
        return "Reglas de la cuenta remunerada"
    for k in sorted(TITULOS, key=lambda k: -len(k.split("|"))):
        partes = k.split("|")
        if cab[:len(partes)] == partes:
            return TITULOS[k]
    return None


def _dividir(texto: str) -> tuple[str, str]:
    """Título (primera oración) y nota (el resto)."""
    m = re.search(r"\.\s+(?=[A-ZÁÉÍÓÚÑ¿(])", texto)
    if not m:
        return texto.strip().rstrip("."), ""
    return texto[:m.start()].strip(), texto[m.end():].strip()


def _parrafo_nuevo(antes_de, texto: str, negrita=False, cursiva=False, mantener=True, nota=False):
    p = OxmlElement("w:p")
    antes_de.addprevious(p)
    par = Paragraph(p, _PADRE[0])
    if nota:
        r = par.add_run("Nota. ")
        r.italic = True
        r.font.size = Pt(9)
        r2 = par.add_run(texto)
        r2.font.size = Pt(9)
    else:
        r = par.add_run(texto)
        r.bold = negrita or None
        r.italic = cursiva or None
    fmt = par.paragraph_format
    fmt.keep_with_next = mantener
    fmt.space_after = Pt(2 if not nota else 8)
    fmt.space_before = Pt(0)
    return p


def _repetir_encabezado(tbl) -> None:
    filas = tbl.findall(qn("w:tr"))
    if not filas:
        return
    trPr = filas[0].get_or_add_trPr()
    if trPr.find(qn("w:tblHeader")) is None:
        trPr.append(OxmlElement("w:tblHeader"))
    for f in filas:
        pr = f.get_or_add_trPr()
        if pr.find(qn("w:cantSplit")) is None:
            pr.append(OxmlElement("w:cantSplit"))


def _vecino(el, hacia_atras: bool):
    """Párrafo no vacío más cercano (sin saltarse imágenes ni tablas)."""
    x = el.getprevious() if hacia_atras else el.getnext()
    while x is not None and x.tag == qn("w:p") and not _texto(x).strip() and not _tiene_imagen(x):
        x = x.getprevious() if hacia_atras else x.getnext()
    return x


def _leyenda(x):
    if x is None or x.tag != qn("w:p"):
        return None
    return LEYENDA.match(_texto(x).strip())


def aplicar(doc, seccion_resultados_desde: str = "14. Diseño experimental") -> dict:
    """Numera, formatea y actualiza menciones. Devuelve el mapa de etiquetas viejas a nuevas."""
    body = doc.element.body
    _PADRE[0] = doc._body
    objetivos, en_resultados = [], False
    for el in list(body.iterchildren()):
        if el.tag == qn("w:p"):
            p = Paragraph(el, _PADRE[0])
            txt = _texto(el).strip()
            if p.style is not None and p.style.name.startswith("Heading"):
                if txt.startswith(seccion_resultados_desde):
                    en_resultados = True
                elif txt.startswith(("Discusión", "Referencias", "Anexo")):
                    en_resultados = False
            if _tiene_imagen(el):
                objetivos.append((el, en_resultados))
        elif el.tag == qn("w:tbl"):
            objetivos.append((el, en_resultados))
    mapa, n, sin_titulo = {}, {"Tabla": 0, "Figura": 0}, []
    for el, res in objetivos:
        if el.tag == qn("w:tbl"):
            previo = _vecino(el, True)
            m = _leyenda(previo)
            n["Tabla"] += 1
            nueva = f"Tabla {n['Tabla']}"
            if m and m.group(1) == "Tabla":
                mapa[("Tabla", m.group(2))] = nueva
                titulo, nota = _dividir(m.group(3))
                body.remove(previo)
            else:
                titulo, nota = _titulo_tabla(el, ""), ""
                if titulo is None:
                    titulo = "Sin título"
                    sin_titulo.append("tabla: " + " | ".join(_clave(el))[:80])
            nota = nota or NOTAS_TABLA.get(titulo, "") or (NOTA_RESULTADOS if res else "")
            _parrafo_nuevo(el, nueva, negrita=True)
            _parrafo_nuevo(el, titulo, cursiva=True)
            _repetir_encabezado(el)
            if nota:
                p = OxmlElement("w:p")
                el.addnext(p)
                nota_p = Paragraph(p, _PADRE[0])
                a_ = nota_p.add_run("Nota. ")
                a_.italic, a_.font.size = True, Pt(9)
                b_ = nota_p.add_run(nota)
                b_.font.size = Pt(9)
                nota_p.paragraph_format.space_after = Pt(10)
        else:
            sig = _vecino(el, False)
            m = _leyenda(sig)
            if not (m and m.group(1) == "Figura"):
                sin_titulo.append("imagen después de: " + _texto(_vecino(el, True))[:70])
                continue
            n["Figura"] += 1
            nueva = f"Figura {n['Figura']}"
            mapa[("Figura", m.group(2))] = nueva
            titulo, nota = _dividir(m.group(3))
            _parrafo_nuevo(el, nueva, negrita=True)
            _parrafo_nuevo(el, titulo, cursiva=True)
            sig_p = Paragraph(sig, _PADRE[0])
            if nota:
                Paragraph(el, _PADRE[0]).paragraph_format.keep_with_next = True
                for r in list(sig_p.runs):
                    r._r.getparent().remove(r._r)
                a_ = sig_p.add_run("Nota. ")
                a_.italic, a_.font.size = True, Pt(9)
                b_ = sig_p.add_run(nota)
                b_.font.size = Pt(9)
                sig_p.paragraph_format.space_after = Pt(10)
            else:
                body.remove(sig)
    _menciones(doc, mapa)
    return {"mapa": mapa, "sin_titulo": sin_titulo, "tablas": n["Tabla"], "figuras": n["Figura"]}


ETIQUETA = re.compile(r"\b(Tablas?|Figuras?)(\s+)((?:[0-9A-C][0-9.]*-[a-z])(?:(?:,\s*|\s+[ya]\s+)(?:[0-9A-C][0-9.]*-[a-z]))*)")


def _menciones(doc, mapa) -> None:
    def cambiar(m):
        tipo = "Tabla" if m.group(1).startswith("Tabla") else "Figura"
        cuerpo = re.sub(r"[0-9A-C][0-9.]*-[a-z]", lambda e: mapa.get((tipo, e.group(0)), f"{tipo} ?{e.group(0)}").split(" ")[1],
                        m.group(3))
        return m.group(1) + m.group(2) + cuerpo
    for p in _parrafos(doc):
        if not ETIQUETA.search(p.text):
            continue
        hecho = True
        for r in p.runs:
            if ETIQUETA.search(r.text):
                r.text = ETIQUETA.sub(cambiar, r.text)
        if ETIQUETA.search(p.text):
            hecho = False
        if not hecho:
            _poner(p, ETIQUETA.sub(cambiar, p.text))


# ---------------------------------------------------------------- página
def pagina(doc) -> None:
    """Número de página arriba a la derecha y una sola fuente (Arial) en todo el documento."""
    for parte in doc.part.package.parts:
        if parte.partname.endswith("theme1.xml"):
            xml = parte.blob.decode("utf-8")
            xml = re.sub(r'(<a:latin typeface=")[^"]*(")', r"\1Arial\2", xml)
            parte._blob = xml.encode("utf-8")
    for sec in doc.sections:
        sec.header.is_linked_to_previous = sec is not doc.sections[0]
    h = doc.sections[0].header
    p = h.paragraphs[0] if h.paragraphs else h.add_paragraph()
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run()
    for tipo, texto in (("begin", None), (None, " PAGE "), ("end", None)):
        if tipo:
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), tipo)
            run._r.append(fc)
        else:
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = texto
            run._r.append(it)
    for nombre in ("Normal", "Heading 1", "Heading 2", "Heading 3", "Title", "TOC 1", "TOC 2", "TOC 3"):
        try:
            st = doc.styles[nombre]
        except KeyError:
            continue
        rpr = st.element.get_or_add_rPr()
        fuentes = rpr.find(qn("w:rFonts"))
        if fuentes is None:
            fuentes = OxmlElement("w:rFonts")
            rpr.insert(0, fuentes)
        for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            fuentes.attrib.pop(qn(a), None)
        for a in ("w:ascii", "w:hAnsi", "w:cs"):
            fuentes.set(qn(a), "Arial")
        if nombre.startswith("Heading"):
            st.paragraph_format.keep_with_next = True


# ---------------------------------------------------------------- revisión
def revisar(doc) -> dict:
    """Citas del texto que no están en la lista y referencias que no se citan."""
    cuerpo = []
    en_refs = False
    for p in doc.paragraphs:
        if p.style.name.startswith("Heading") and p.text.strip() == "Referencias":
            en_refs = True
        elif p.style.name.startswith("Heading") and p.text.startswith("Anexo"):
            en_refs = False
        if not en_refs:
            cuerpo.append(p.text)
    for t in doc.tables:
        for f in t.rows:
            for c in f.cells:
                cuerpo.append(c.text)
    texto = "\n".join(cuerpo)
    sin_cita = []
    for clave, _ in REFERENCIAS:
        autor, anio = clave.rsplit(", ", 1)
        autor = autor.replace(" et al.", "")
        corto = autor.split(" y ")[0].split(",")[0]
        if not re.search(re.escape(corto.replace("La Lonja", "Lonja")) + r"[^()]{0,60}\(?" + re.escape(anio), texto, re.I) and clave not in texto:
            sin_cita.append(clave)
    return {"referencias_sin_cita": sin_cita}

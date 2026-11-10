"""Arma docs/informe_final.docx a partir del Word de los puntos 1 a 7 (que no se modifica)
y agrega los puntos 8 a 19, la discusión, las conclusiones, las limitaciones y los anexos.

  python tools/construir_informe.py [ruta del Word de los puntos 1 a 7]

Todos los números salen de outputs/resultados y de la hoja PARAMETROS; si se vuelve a
correr el experimento, basta volver a correr este script. Después, tools/finalizar_word.ps1
actualiza la tabla de contenido y exporta el PDF con Word.
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from docx.text.paragraph import Paragraph

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / "src"))
sys.path.insert(0, str(Path(__file__).parent))
from investment_sim.config.parameters import cargar  # noqa: E402
from investment_sim.models import inflation, markov  # noqa: E402
import formato_apa  # noqa: E402
import limpieza_texto  # noqa: E402
from informe_1_a_7 import actualizar  # noqa: E402
from parametros_def import FILAS  # noqa: E402

ORIGEN = RAIZ.parent / "insumos_equipo" / "07_consolidado_puntos_1_a_7" / "Trabajo_Final_Simulacion_puntos_1_a_7.docx"
DESTINO = RAIZ / "docs" / "informe_final.docx"
RES = RAIZ / "outputs" / "resultados"
FIG = RAIZ / "docs" / "figuras"
DIAG = RAIZ / "docs" / "diagramas"
AZUL = "1F4E78"
EST = ["E1", "E2", "E3", "E4"]
NOM = {"E1": "E1 Seguridad", "E2": "E2 Moderada", "E3": "E3 Arriesgada", "E4": "E4 Dinámica"}
ESC = {"base": "base", "adverso": "adverso", "estres": "estrés"}

# ---------------------------------------------------------------- datos
P = cargar("base")
PA = cargar("adverso")
R = pd.read_csv(RES / "resumen.csv").set_index(["escenario", "estrategia"])
PAR = pd.read_csv(RES / "comparaciones_pareadas.csv")
NR = pd.read_csv(RES / "nominal_vs_real.csv").set_index(["escenario", "estrategia"])
CN = pd.read_csv(RES / "capital_necesario.csv")
CM = pd.read_csv(RES / "capital_meta.csv")
TOR = pd.read_csv(RES / "tornado.csv")
SM = pd.read_csv(RES / "sensibilidad_modelos.csv")
VAL = pd.read_csv(RES / "validacion.csv")
PREC = pd.read_csv(RES / "precision_replicas.csv")
CORR = json.loads((RES / "corrida.json").read_text(encoding="utf-8"))
TRAZA = pd.read_csv(RES / "traza_E4_replica0.csv")
INTER = pd.read_csv(RES / "interaccion.csv")
N = CORR["replicas"]


def c(x: float, dec: int = 0) -> str:
    """Número con formato colombiano, redondeo comercial (0,25 → 0,3) y signo menos tipográfico."""
    from decimal import ROUND_HALF_UP, Decimal
    q = Decimal(repr(round(float(x), dec + 6))).quantize(Decimal(1).scaleb(-dec), rounding=ROUND_HALF_UP)
    txt = f"{q:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return txt.replace("-", "−")


def sg(x: float, dec: int = 1) -> str:
    """Con signo explícito: +0,3 o −0,5."""
    t = c(x, dec)
    return t if t.startswith("−") else "+" + t


def principales(e: str = "E2", k: int = 5) -> str:
    """Los k factores del tornado con más efecto sobre P(meta), en orden, en minúscula."""
    nombres = {"Rendimiento medio del ETF": "el rendimiento medio del ETF", "Nivel de la inflación": "el nivel de la inflación",
               "Valorización": "la valorización del inmueble", "Frecuencia de necesidades": "la frecuencia de los gastos imprevistos",
               "Monto mediano de las necesidades": "su monto", "Memoria de la inflación": "la memoria de la inflación",
               "Volatilidad del ETF": "la volatilidad del ETF", "Canon inicial": "el canon inicial",
               "TPM inicial": "la TPM inicial", "Descuento por venta urgente": "el descuento de la venta urgente"}
    t = TOR[TOR.estrategia == e].sort_values("rango_p_meta", ascending=False).head(k)
    lista = [next((v for kk, v in nombres.items() if x.startswith(kk)), x.split(" (")[0].lower()) for x in t.factor]
    return ", ".join(lista[:-1]) + " y " + lista[-1]


def pc(x: float, dec: int = 1) -> str:
    return f"{c(x * 100, dec)} %"


def mm(x: float) -> str:
    return f"{c(x / 1e6)} millones"


def r(esc, e, col):
    return float(R.loc[(esc, e), col])


def par(esc, a, b, var):
    q = PAR[(PAR.escenario == esc) & (PAR.a == a) & (PAR.b == b) & (PAR.variable == var)].iloc[0]
    return q


def holm(esc, a, b, var, alfa=0.05) -> bool:
    """¿La diferencia sigue siendo significativa con la corrección de Holm sobre todas las comparaciones?"""
    orden = PAR.sort_values("p_valor").reset_index()
    m = len(PAR)
    for k, x in orden.iterrows():
        if x.p_valor > alfa / (m - k):
            return False
        if (x.escenario, x.a, x.b, x.variable) == (esc, a, b, var):
            return True
    return False


def frase_e4() -> str:
    """Cómo queda E4 frente a E2 en patrimonio, según la significancia con Holm."""
    v = par("base", "E4", "E2", "W_real")
    if holm("base", "E4", "E2", "W_real"):
        return f"termina en promedio {mm(-v.diferencia)} por debajo en el escenario base"
    return (f"termina en promedio {mm(-v.diferencia)} por debajo en el escenario base, una diferencia que deja de ser "
            "significativa con la corrección de Holm, así que en patrimonio las dos son prácticamente iguales")


def cap(esc, e, meta, col="capital_80"):
    v = CN[(CN.escenario == esc) & (CN.estrategia == e) & (CN.meta == meta)][col].iloc[0]
    return v


def mejor(esc, col, mayor=True):
    s = R.loc[esc][col]
    return s.idxmax() if mayor else s.idxmin()


def sm(modelo_inicio, e):
    g = SM[SM.modelo.str.startswith(modelo_inicio) & (SM.estrategia == e)].iloc[0]
    return (g.alt_p_meta - g.ref_p_meta) * 100


def rango_sm(modelo_inicio, ests=("E2", "E3", "E4")):
    v = [abs(sm(modelo_inicio, e)) for e in ests]
    return f"{c(min(v))} y {c(max(v))}"


def alt(modelo_inicio, e):
    return float(SM[SM.modelo.str.startswith(modelo_inicio) & (SM.estrategia == e)].alt_p_meta.iloc[0])


def tor(factor_inicio, e):
    return float(TOR[TOR.factor.str.startswith(factor_inicio) & (TOR.estrategia == e)].rango_p_meta.iloc[0])


def pm(esc, e, capital, meta=1e9):
    return float(CM[(CM.escenario == esc) & (CM.estrategia == e) & (CM.capital == capital) & (CM.meta == meta)].p_meta.iloc[0])


TASA_REAL = float(VAL[VAL.medida.str.startswith("Tasa real")].simulado.iloc[0])


# ---------------------------------------------------------------- utilidades de Word
class Informe:
    def __init__(self, doc: Document):
        self.doc = doc
        self.fig = {}
        self.tab = {}

    # Escritura al final del documento
    def h(self, texto: str, nivel: int = 1):
        return self.doc.add_heading(texto, level=nivel)

    def p(self, *partes, centro=False, tam=None):
        """Párrafo; un texto que empieza con '**' va en negrita (sin los asteriscos)."""
        par_ = self.doc.add_paragraph()
        for t in partes:
            negrita = t.startswith("**")
            run = par_.add_run(t[2:] if negrita else t)
            run.bold = negrita or None
            if tam:
                run.font.size = Pt(tam)
        if centro:
            par_.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return par_

    def vinetas(self, items):
        for i in items:
            if isinstance(i, tuple):
                self.p("• ", f"**{i[0]}", f" {i[1]}")
            else:
                self.p(f"• {i}")

    def leyenda(self, tipo: str, seccion: str, texto: str):
        clave = (tipo, seccion)
        k = self.fig.get(clave, 0)
        self.fig[clave] = k + 1
        etiqueta = f"{tipo} {seccion}-{'abcdefghijklmnopqrstuvwxyz'[k]}. "
        par_ = self.doc.add_paragraph()
        a = par_.add_run(etiqueta)
        a.bold = True
        a.font.size = Pt(9.5)
        b = par_.add_run(texto)
        b.font.size = Pt(9.5)
        return etiqueta.strip(" .")

    def figura(self, archivo: Path, seccion: str, texto: str, ancho: float = 6.3):
        par_ = self.doc.add_paragraph()
        par_.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par_.add_run().add_picture(str(archivo), width=Inches(ancho))
        return self.leyenda("Figura", seccion, texto)

    def tabla(self, seccion: str, titulo: str, encabezado: list[str], filas: list[list], anchos: list[float] | None = None,
              tam: float = 9):
        etiqueta = self.leyenda("Tabla", seccion, titulo)
        t = self.doc.add_table(rows=1, cols=len(encabezado))
        t.style = self.doc.tables[0].style
        for j, txt in enumerate(encabezado):
            celda = t.rows[0].cells[j]
            _sombra(celda, AZUL)
            run = celda.paragraphs[0].add_run(str(txt))
            run.bold = True
            run.font.size = Pt(tam)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            celda.paragraphs[0].paragraph_format.space_after = Pt(0)
        for fila in filas:
            celdas = t.add_row().cells
            for j, v in enumerate(fila):
                run = celdas[j].paragraphs[0].add_run("" if v is None else str(v))
                run.font.size = Pt(tam)
                celdas[j].paragraphs[0].paragraph_format.space_after = Pt(0)
        if anchos:
            # Word respeta los anchos solo con diseño fijo y la cuadrícula de la tabla al día.
            t.autofit = False
            grid = t._tbl.tblGrid
            for j, a in enumerate(anchos):
                if j < len(grid.gridCol_lst):
                    grid.gridCol_lst[j].w = Inches(a)
            for fila in t.rows:
                for j, a in enumerate(anchos):
                    fila.cells[j].width = Inches(a)
        self.doc.add_paragraph()
        return etiqueta

    def salto(self):
        self.doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def _sombra(celda, color):
    tcPr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)


def buscar(doc, inicio: str, encabezado: bool | None = None) -> Paragraph:
    for x in doc.paragraphs:
        if x.text.strip().startswith(inicio) and (encabezado is None or x.style.name.startswith("Heading") == encabezado):
            return x
    raise KeyError(inicio)


def poner(x: Paragraph, texto: str) -> None:
    """Cambia el texto de un párrafo y conserva el formato de su primer run."""
    runs = x.runs
    runs[0].text = texto
    for run in runs[1:]:
        run._r.getparent().remove(run._r)


def antes(doc, ancla: Paragraph, partes: list[str], estilo: str = "Normal") -> Paragraph:
    nuevo = OxmlElement("w:p")
    ancla._p.addprevious(nuevo)
    x = Paragraph(nuevo, ancla._parent)
    x.style = doc.styles[estilo]
    for t in partes:
        negrita = t.startswith("**")
        run = x.add_run(t[2:] if negrita else t)
        if negrita:
            run.bold = True
    return x


# ---------------------------------------------------------------- partes que se insertan en el Word original
def portada_y_resumen(doc):
    poner(doc.paragraphs[1], "Trabajo final de Simulación. Informe completo: planteamiento, procesos aleatorios, modelo de "
                             "eventos discretos, estrategias, diseño experimental, verificación, validación y resultados")
    poner(buscar(doc, "Integrantes:"), "Integrantes: Daniel Giraldo Arenas, Allison Correa y Samuel Alexis Vásquez Muñoz")
    fecha = buscar(doc, "Noviembre de 2026")
    for t in ("Universidad Pontificia Bolivariana", "Curso: Simulación", "Profesor: John Fernando Vargas Buitrago"):
        antes(doc, fecha, [t])
    poner(fecha, "10 de noviembre de 2026")
    for inicio in ("Archivo de soporte", "Este documento cubre los puntos 1 a 7"):
        x = buscar(doc, inicio)
        x._p.getparent().remove(x._p)


def notas_7(doc):
    antes(doc, buscar(doc, "7.4.10 Dependencia", True), [
        "**Uso en el simulador. ",
        "El simulador usa M5 como referencia, con los residuos observados sorteados por bootstrap. El margen por fase de la TPM "
        "queda como alternativa y su efecto sobre las conclusiones se mide en la sección 16.3 (cambio en la probabilidad de meta "
        "de un punto o menos)."])


# ---------------------------------------------------------------- secciones nuevas
def seccion_8(w: Informe):
    Pm = markov.matriz(P)
    pi = markov.estacionaria(Pm)
    dur = 1 / (1 - np.diag(Pm))
    w.salto()
    w.h("8. Dependencia entre variables y estados económicos")
    w.p("No suponemos que las variables sean independientes. El modelo tiene dos estructuras de dependencia. La primera es una "
        "cadena causal: la inflación mueve la decisión de la Junta, la TPM fija la tasa de la cuenta (α × TPM) y la de cada CDT "
        "que se abre, y la inflación reajusta el canon y los gastos del inmueble. La segunda es una cadena de Markov de tres "
        "estados económicos que cambia a la vez la volatilidad del ETF, la vacancia y la frecuencia de los gastos imprevistos, y "
        "que la estrategia E4 usa para decidir.")
    w.h("8.1 Estados y matriz de transición", 2)
    w.p("Clasificamos cada trimestre de 2006 a 2026 según el crecimiento anual del PIB real (DANE): expansión si pasa de 4 %, "
        "crisis si es menor que 1 % y normal en otro caso. De 82 trimestres, 30 fueron de expansión, 40 normales y 12 de crisis. "
        "La matriz sale de contar las transiciones de cada estado a cada otro y dividir por el total de la fila.")
    w.p("Una alternativa frecuente usa cinco estados: expansión, normalidad, desaceleración, crisis y recuperación. "
        "Se usaron tres. Con 81 transiciones, una matriz de 5 × 5 tiene 25 casillas y varias quedarían con uno o dos datos, sobre "
        "todo alrededor de las crisis, que en el periodo son solo tres episodios. Lo necesario es una estructura de dependencia, "
        "y con tres estados las probabilidades se pueden estimar con los datos que hay.")
    filas = [[s] + [c(Pm[i, j], 3) for j in range(3)] + [c(dur[i], 1), pc(pi[i], 0)] for i, s in enumerate(markov.ESTADOS)]
    w.tabla("8", "Matriz de transición trimestral y lo que implica.", ["Desde \\ hacia", "Expansión", "Normal", "Crisis",
            "Duración media (trimestres)", "Tiempo a largo plazo"], filas)
    w.p("Los estados son persistentes, y en veinte años la economía nunca pasó directo de expansión a crisis ni al revés. La "
        "simulación empieza en Normal, que es el estado del segundo trimestre de 2026 (crecimiento de 3,5 %), y cada tres meses "
        "sortea el estado siguiente con la fila del estado vigente.")
    w.p("La matriz hay que leerla con cuidado. Con tan pocos datos los intervalos son anchos: con bootstrap, la probabilidad de "
        "seguir en crisis va de 0,25 a 0,89 y la de pasar de normal a crisis de 0,02 a 0,17. Parte de la persistencia es "
        "mecánica, porque el crecimiento anual de un trimestre comparte tres trimestres con el del anterior. La duración media de una crisis "
        f"simulada ({c(float(VAL[VAL.medida.str.startswith('Duración media de una crisis')].simulado.iloc[0]), 1)} trimestres) es "
        "parecida a la observada (4), pero en diez años una réplica puede encadenar crisis: la racha más larga de cada réplica "
        "tiene mediana de seis trimestres, más que cualquiera de las tres crisis reales. Lo que sí comprobamos es que la matriz no cambia entre 2006-2015 y "
        "2016-2026 (G = 4,29; p = 0,37). La probabilidad de crisis entra al análisis de sensibilidad.")
    w.h("8.2 Qué cambia con el estado", 2)
    filas = [
        ["ETF: volatilidad mensual", pc(P.f("etf.vol_E"), 1), pc(P.f("etf.vol_N"), 1), pc(P.f("etf.vol_C"), 1), "MARKOV_ETF"],
        ["Vacancia: probabilidad de no renovar", pc(P.f("vac.p_E"), 0), pc(P.f("vac.p_N"), 0), pc(P.f("vac.p_C"), 0), "Supuesto"],
        ["Vacancia: duración media (meses)", c(P.f("vac.media_E"), 1), c(P.f("vac.media_N"), 1), c(P.f("vac.media_C"), 1), "7.7"],
        ["Gastos imprevistos: uno cada ... meses", c(1 / P.f("liq.tasa_E")), c(1 / P.f("liq.tasa_N")), c(1 / P.f("liq.tasa_C")), "Supuesto"],
        ["E4: reparto objetivo", "Más ETF", "Intermedio", "Más CDT y cuenta", "Sección 11"],
    ]
    w.tabla("8", "Efecto del estado económico sobre los demás procesos.", ["Variable", "Expansión", "Normal", "Crisis", "Fuente"], filas)
    w.p("Así, cuando llega una crisis el ETF se vuelve más riesgoso, arrendar cuesta más y los imprevistos llegan el doble de "
        "seguido, todo en la misma réplica. Dos cosas no cambian con el estado porque los datos no lo sostienen: el rendimiento "
        "medio del ETF (Kruskal-Wallis, p = 0,39) y la valorización del inmueble (Welch, p = 0,275). Preferimos no inventar esas "
        "diferencias, y esa decisión explica buena parte del resultado de E4 (sección 18).")


def seccion_9(w: Informe):
    w.h("9. Modelo de simulación de eventos discretos")
    w.p("El simulador está escrito en Python y no usa ninguna librería de simulación: el reloj, la lista de eventos y los "
        "manejadores son nuestros, para que cada regla sea visible. El tiempo se mide en meses, empieza en 0 y termina en H = "
        f"{P.i('inv.horizonte')}. La lista de eventos futuros es un montículo (heapq) ordenado por tiempo, luego por la prioridad "
        "del diagrama y luego por orden de llegada; así dos eventos en el mismo instante siempre se procesan en el mismo orden. En "
        "cada paso el motor saca el evento más próximo, adelanta el reloj, ejecuta el manejador (que cambia el estado y programa "
        "eventos futuros) y vuelve a empezar. Un evento cancelado, como el próximo arriendo de un inmueble que se vendió, se marca "
        "y se descarta al salir de la lista. Es el esquema de próximo evento que describen Banks et al. (2010) y Law (2015).")
    w.p("El modelo es híbrido. Las necesidades de liquidez, el fin de una vacancia, la llegada del dinero de una venta urgente y "
        "el inicio de la crisis forzada ocurren en tiempos continuos, entre un mes y otro. El rendimiento del ETF, los intereses de "
        "la cuenta, el índice de precios y la regla de liquidez mínima se actualizan en un evento periódico de cierre de mes (el "
        "evento 1), porque sus datos son mensuales. Los demás eventos periódicos (Junta, inflación, estado, vencimientos, "
        "aniversarios del arriendo) entran a la misma lista con su propio calendario.")
    w.h("9.1 Eventos", 2)
    filas = [
        ["0", "Inicio", "Compra los activos con los pesos de la estrategia y programa el primer evento de cada proceso", "0"],
        ["2", "Cambio de estado económico", "Sortea el estado del trimestre siguiente con la matriz", "Cada 3 meses"],
        ["3", "Inflación anual", "Sortea la inflación del próximo diciembre; la ruta mensual es lineal entre diciembres", "Cada 12 meses"],
        ["4", "Reunión de la Junta", "Decide bajar, mantener o subir la TPM y el tamaño del cambio", "8 al año"],
        ["9", "Valorización y gastos anuales", "Valoriza el inmueble y cobra predial, mantenimiento y pintura", "Cada 12 meses"],
        ["14", "Revisión anual (solo E4)", "Compara el reparto con el objetivo del estado y rebalancea si sale de la banda", "Cada 12 meses"],
        ["5", "Vencimiento de CDT", "Paga intereses con retención y abre un CDT nuevo a la tasa del momento", "Cada 12 meses"],
        ["7", "Fin de contrato de arriendo", "Sortea si el inquilino renueva; si no, empieza una vacancia de duración gamma", "Cada 12 meses"],
        ["8", "Fin de vacancia", "Entra un inquilino nuevo con el canon de mercado", "Aleatorio"],
        ["6", "Pago de arriendo", "Abona el canon menos la comisión de la inmobiliaria", "Mensual si está ocupado"],
        ["10", "Reparación extraordinaria", "Programado en el código pero apagado (frecuencia 0, sin datos)", "Poisson"],
        ["11", "Necesidad de liquidez", "Sortea el monto y llama a la cobertura (12); programa la siguiente", "Poisson según el estado"],
        ["12", "Cobertura de liquidez", "Obtiene el efectivo: cuenta, ETF, venta anticipada de CDT, venta urgente del inmueble", "Al instante, la llama otro evento"],
        ["13", "Venta del inmueble", "Recibe el producto de la venta urgente, con descuento y costos", "Entre 4 y 7 meses después de decidirla (uniforme)"],
        ["1", "Cierre de mes", "Rendimiento del ETF, intereses de la cuenta, gastos fijos, liquidez mínima, registro del patrimonio", "Mensual"],
        ["16 y 17", "Inicio y fin de la crisis forzada", "Fija el estado en Crisis y lo libera (solo estrés)", "Aleatorio"],
        ["15", "Fin", "Valora el inmueble neto de costos de venta y calcula los indicadores", "H"],
    ]
    w.tabla("9", "Eventos del modelo, en el orden de prioridad que se usa en los empates. El número es el del diagrama de eventos.",
            ["N.º", "Evento", "Qué hace", "Cuándo ocurre"], filas, [0.5, 1.6, 3.2, 1.3])
    w.p("Los eventos 3, 5, 7, 9 y 14 caen el mismo mes cada año. El orden de prioridad hace que la inflación se conozca antes de "
        "que se reúna la Junta, que la Junta decida antes de que venza el CDT y que la revisión de E4 llegue antes del vencimiento, "
        "de modo que E4 decide cuánto renovar sin vender el CDT antes de tiempo. La compra o venta de activos y el pago de "
        "dividendos no son eventos separados: las compras ocurren en el inicio y en la revisión de E4, las ventas en la cobertura "
        "de liquidez, y los dividendos del ETF ya están dentro de su rendimiento total (sección 7.5).")
    w.h("9.2 Variables de estado", 2)
    filas = [
        ["Reloj", "t (meses)"], ["Efectivo", "Saldo de la cuenta"], ["Renta fija", "Lista de CDT abiertos: capital, tasa, vencimiento"],
        ["Renta variable", "Valor del ETF y su costo de compra"], ["Inmueble", "Valor, canon vigente, ocupado o vacante, venta en curso"],
        ["Deuda", "Obligaciones que no se pudieron cubrir"], ["Economía", "Estado, TPM, inflación anual, índice de precios"],
        ["Acumuladores", "Arriendos, gastos, impuestos, comisiones, intereses perdidos, necesidades y cómo se cubrieron"],
        ["Patrimonio", "Nominal y real cada mes, máximo previo (para la caída máxima), primer mes en la meta"],
        ["Liquidez", "Episodios de iliquidez, su duración y el mes en que empezó el actual"],
    ]
    w.tabla("9", "Variables de estado.", ["Grupo", "Variables"], filas, [1.6, 5.0])
    w.h("9.3 Generadores de números aleatorios", 2)
    w.p("Cada réplica r tiene una semilla (semilla base + r). De ella salen 13 generadores independientes con "
        "numpy.random.SeedSequence: estado, inflación, Junta, CDT, ETF, valorización, vacancia, canon, reparaciones, tiempo y "
        "monto de las necesidades, crisis forzada y venta del inmueble. Que cada proceso tenga su propio generador importa por dos "
        "razones. Si una estrategia vende el inmueble y deja de sortear vacancias, los demás procesos no se corren y la economía "
        "sigue idéntica. Y dos estrategias con la misma semilla ven exactamente la misma inflación, la misma Junta y el mismo "
        "ETF, que es lo que permite compararlas réplica a réplica (sección 14).")
    w.h("9.4 Modelos de referencia del simulador", 2)
    w.p("Para algunos procesos se estudió más de un modelo. El de referencia es el que usan los resultados; el otro queda en el "
        "código solo para el análisis de sensibilidad (sección 16.3).")
    filas = [
        ["Estados económicos", "Cadena de Markov de tres estados (umbral 1 % y 4 %)", "No"],
        ["Junta", "Logit ordenado con tendencia y brecha de Taylor (7.2.7)", "Tabla por tendencia (7.2.3)"],
        ["Margen del CDT", "M5 con residuos observados (7.4)", "Normal por fase de la TPM"],
        ["ETF", "t de Student (ν = 5) con volatilidad por estado (7.5.11)", "Jones-Faddy sin estados (7.5.9)"],
        ["Valorización", "AR(1) del IPVU (7.6)", "Normal por estado con el IPVNBR de Medellín (vivienda nueva)"],
        ["Reajuste del canon", "Modelo DANE con tope en el IPC previo (5.4.3)", "Sube todo el IPC"],
    ]
    w.tabla("9", "Modelo de referencia y alternativa de cada proceso.", ["Proceso", "Referencia", "Alternativa"], filas)


def seccion_10(w: Informe):
    w.h("10. Diagrama de eventos")
    w.p("El diagrama de eventos y el diagrama de flujo se elaboraron antes de programar. El motor sigue su numeración: "
        "cada manejador lleva el número de su evento, y la prioridad de los empates es la de la página 1 del diagrama de flujo. "
        "Los originales en PDF están en docs/diagramas del repositorio. Sobre el diagrama original se hicieron dos agregados que "
        "no cambian su lógica: los eventos 16 y 17 de la crisis forzada del escenario de estrés, y que la cobertura de liquidez "
        "(12) se ejecute en el mismo instante del evento que la necesita, en vez de programarse en la lista. El texto de las cajas "
        "describe los modelos que usa el simulador; la estructura y las flechas son las del diseño original.")
    w.p("El diagrama es grande y en una página carta solo se ve su estructura. La Figura 10-a lo muestra completo como vista "
        "general. Para leer el texto de las cajas se entregan con el informe los PDF vectoriales, que se pueden ampliar sin "
        "perder calidad: diagrama_eventos.pdf (una hoja) y diagrama_de_flujo_completo.pdf (cinco hojas), en la carpeta "
        "docs/diagramas del repositorio.")
    if (DIAG / "diagrama_eventos.png").exists():
        w.figura(DIAG / "diagrama_eventos.png", "10", "Diagrama de eventos, vista general. El detalle se lee en "
                                                      "docs/diagramas/diagrama_eventos.pdf.")


def seccion_11(w: Informe):
    w.h("11. Estrategias de inversión")
    w.p("Las reglas se fijaron antes de correr el experimento y no se ajustaron después de ver los resultados. Las cuatro "
        "estrategias arrancan con el mismo capital y en las mismas condiciones; solo cambia cómo reparten y cómo deciden.")
    filas = []
    for e in EST:
        pes = P.lista(f"{e}.pesos")
        filas.append([NOM[e]] + [pc(x, 0) if x else "0 %" for x in pes])
    w.tabla("11", "Reparto inicial del capital.", ["Estrategia", "Cuenta", "CDT 12 meses", "ETF iCOLCAP", "Inmueble"], filas)
    w.p("E1 es para quien prioriza no perder y tener el dinero a mano: 80 % en cuenta y CDT, sin inmueble. E2 reparte entre los "
        "cuatro activos para que ninguno domine. E3 apuesta por crecer: 85 % en inmueble y ETF, con poco efectivo. E1, E2 y E3 "
        "compran y mantienen; solo venden para cubrir liquidez y renuevan cada CDT por el mismo capital.")
    w.p(f"E4 es la estrategia dinámica. Empieza igual que E2 y cada {P.i('E4.cada')} meses compara su portafolio financiero (cuenta, "
        "CDT y ETF; el inmueble no se toca) con un objetivo que depende del estado de la economía. Si algún activo se aleja del "
        f"objetivo más de {c(P.f('E4.banda') * 100)} puntos, rebalancea. La revisión coincide con el vencimiento del CDT, así que el "
        "monto que se renueva se fija ahí; si hace falta comprar ETF con plata del CDT, la compra espera a que el CDT venza.")
    filas = [[s] + [pc(x, 1) for x in P.lista(f"E4.obj_{L}")] for s, L in (("Expansión", "E"), ("Normal", "N"), ("Crisis", "C"))]
    w.tabla("11", "Objetivo de E4 para el portafolio financiero según el estado.", ["Estado al revisar", "Cuenta", "CDT", "ETF"], filas)
    w.p("Reglas comunes:")
    w.vinetas([
        ("Liquidez mínima.", f"Al cierre de cada mes la cuenta debe tener al menos {pc(P.f('inv.liquidez_min'), 0)} del patrimonio. "
                             "Si no, se repone con la cobertura de la sección 12."),
        ("Inmueble.", f"Vale el porcentaje del capital de cada estrategia; la compra cuesta {pc(P.f('inv.capital') * 0 + P.f('inm.costo_compra'), 1)} "
                      f"del valor y se arrienda desde el primer mes con un canon de {pc(P.f('inm.canon_inicial'), 1)} del valor."),
        ("Costos.", f"Comisión de {pc(P.f('etf.comision'), 1)} al comprar o vender ETF, retención de {pc(P.f('imp.retencion'), 0)} sobre los intereses del CDT y de "
                    f"{pc(P.f('imp.retencion_cuenta'), 0)} sobre los de la cuenta, 4x1000 sobre los retiros que pasan la exención mensual de {mm(P.f('imp.gmf_exencion'))} (350 UVT) y "
                    f"comisión de la inmobiliaria de {pc(P.f('inm.comision_inmobiliaria'), 1)} del canon."),
    ])


def seccion_12(w: Informe):
    w.h("12. Evento extraordinario de liquidez")
    pct_liq = P.f("liq.mediana_pct")
    p90_liq = pct_liq * float(np.exp(1.2816 * P.f("liq.sigma")))
    w.p("Los gastos imprevistos (una enfermedad, una reparación de la casa propia, una ayuda a la familia) llegan como un proceso "
        f"de Poisson: en expansión y normal, en promedio uno cada {c(1 / P.f('liq.tasa_N'))} meses; en crisis, uno cada "
        f"{c(1 / P.f('liq.tasa_C'))}. El monto es lognormal con mediana de {pc(pct_liq, 1)} del capital inicial, en pesos de hoy, y σ = "
        f"{c(P.f('liq.sigma'), 1)} en el logaritmo, y se indexa con el índice de precios simulado. Al ser una fracción del capital, "
        f"el imprevisto pesa lo mismo sin importar cuánto se invierte. Con esos valores, una de cada diez necesidades pasa de "
        f"{pc(p90_liq, 1)} del capital inicial. En el horizonte de referencia el inversionista enfrenta en promedio "
        f"{c(r('base', 'E1', 'n_necesidades_media'), 1)} necesidades en el escenario base.")
    w.p("Estos valores son un supuesto, porque no hay datos de monto, y se buscó con qué contrastarlos. La Encuesta de Carga Financiera y Educación Financiera de los Hogares (IEFIC) del DANE, de 2017 y 2018, "
        "pregunta si en el último mes los gastos del hogar superaron sus ingresos y cómo se cubren esos meses. Entre el 10 % de "
        "hogares de Medellín con más ingreso, un mes con gastos mayores que los ingresos cubierto con ahorros o venta de activos "
        "aparece una vez cada 16 a 23 meses. No la usamos para calibrar, por dos razones: cuenta déficits de cualquier tamaño, "
        "también los pequeños, y la pregunta de cómo se cubren es sobre lo que el hogar suele hacer, no sobre ese mes. Sí muestra "
        "que una necesidad cada dos años no es exagerada. Para el monto no hay ninguna fuente. Como el supuesto pesa "
        "mucho, la sección 16.3 muestra los resultados sin necesidades extraordinarias. (Fuente: Departamento Administrativo "
        "Nacional de Estadística: www.dane.gov.co. Los microdatos no se publican con este trabajo; solo el resumen.)")
    w.p("Cuando llega una necesidad, o cuando la cuenta queda por debajo del mínimo al cierre de mes, el simulador consigue el "
        "efectivo en este orden y registra lo que cuesta cada paso:")
    w.vinetas([
        ("Cuenta.", "Sin costo, salvo el 4x1000 si se pasa la exención."),
        ("Venta de ETF.", f"Comisión de {pc(P.f('etf.comision'), 1)} y 4x1000; si el ETF está en pérdida, la pérdida queda realizada."),
        ("Venta anticipada de CDT.", "El CDT se vende al valor K(1 + r₀)^T / (1 + r_mkt + c)^(T − τ), donde r_mkt es la tasa de "
                                     f"mercado del momento y c = {c(P.f('cdt.castigo') * 100)} punto es el castigo. Se pierde parte de los intereses."),
        ("Venta urgente del inmueble.", f"Descuento de {pc(P.f('inm.descuento_urgente'), 0)} y costos de venta de {pc(P.f('inm.costo_venta'), 2)}; "
                                        f"el dinero llega entre {P.i('inm.venta_min')} y {P.i('inm.venta_max')} meses después."),
    ])
    w.p("Si nada alcanza, lo que falta queda como deuda y se paga con el primer ingreso que llegue. Llamamos episodio de iliquidez "
        "al periodo en que el inversionista no puede pagar sin vender el inmueble de urgencia o sin quedar debiendo; el episodio "
        "termina cuando la deuda vuelve a cero y la cuenta recupera el mínimo.")
    w.p("Este componente separa tener patrimonio de tener liquidez. En el escenario base, al final del horizonte E1 tiene "
        f"{mm(r('base', 'E1', 'liquido_final_real_media'))} reales en activos líquidos y E3 solo {mm(r('base', 'E3', 'liquido_final_real_media'))}, "
        f"aunque E3 tiene más patrimonio ({mm(r('base', 'E3', 'W_real_media'))} contra {mm(r('base', 'E1', 'W_real_media'))} en promedio). "
        "La sección 18 muestra qué pasa con esa diferencia cuando la economía se pone difícil.")


def seccion_13(w: Informe):
    w.h("13. Escenarios")
    w.p("Los tres escenarios usan el mismo modelo y las mismas dependencias; cambian parámetros, y el de estrés agrega un evento. "
        "El estado económico sigue moviendo la volatilidad del ETF, la vacancia, la frecuencia de las necesidades de liquidez y "
        "las decisiones de E4, y la inflación sigue moviendo la TPM, así que un cambio en un escenario se transmite a las demás "
        "variables. La valorización y la TPM no dependen del estado (secciones 7.6 y 7.2), salvo en la crisis forzada del estrés, "
        "que fija sus propios valores.")
    nivel_b = inflation.nivel_largo_plazo(P.f("inf.c"), P.f("inf.phi"))
    nivel_a = inflation.nivel_largo_plazo(PA.f("inf.c"), PA.f("inf.phi"))
    filas = [
        ["Estados económicos", "Matriz estimada", f"Normal → Crisis {c(PA.f('markov.N_C'), 2)}; Crisis → Normal {c(PA.f('markov.C_N'), 2)}",
         "Estado fijo en Crisis durante la crisis forzada"],
        ["Inflación", f"Nivel normal {c(nivel_b, 1)} %", f"Nivel normal {c(nivel_a, 1)} %", "Base"],
        ["TPM (Junta)", "Logit", "Logit (sube más por la inflación)", "Bajadas de 0,5 a 2 puntos por reunión"],
        ["Cuenta", f"{c(P.f('cta.alfa'), 3)} × TPM", f"{c(PA.f('cta.alfa'), 2)} × TPM", "Base"],
        ["ETF", f"μ = {pc(P.f('etf.mu'), 2)} mensual", "μ = 0", "Caída de 25 % al empezar la crisis"],
        ["Valorización", "AR(1) del IPVU, media 7,55 %", "Media 2 puntos menos", "−5 % anual durante la crisis"],
        ["Vacancia", "5 / 10 / 30 %; 1 / 1,5 / 6 meses", "Probabilidad y duración × 1,5", "30 % y 9 meses durante la crisis"],
        ["Necesidades de liquidez", "1 cada 24 meses (12 en crisis)", "Frecuencia × 1,5", "1 cada 12 meses y monto × 2 durante la crisis"],
    ]
    w.tabla("13", "Qué cambia en cada escenario.", ["Elemento", "Base", "Adverso", "Estrés"], filas)
    w.p("En el escenario de estrés la crisis empieza en un mes uniforme entre 1 y H − 12 y dura entre 12 y 24 meses (uniforme), "
        "más que la crisis media de los datos. Mientras dura, el estado queda fijo en Crisis y rigen los valores de la columna "
        "Estrés; al terminar, la economía pasa a Normal y todo vuelve a los valores base. El inicio y la duración se sortean con "
        "su propio generador, así que en una misma réplica las cuatro estrategias enfrentan la misma crisis en el mismo momento.")


def seccion_14(w: Informe):
    w.h("14. Diseño experimental")
    w.p(f"Simulamos 12 configuraciones (4 estrategias por 3 escenarios) con {c(N)} réplicas cada una. El indicador principal es la "
        "probabilidad de alcanzar la meta, y exigimos que su error sea de ±3 puntos con 95 % de confianza en el peor caso (p = 0,5): "
        "n = 1,96² × 0,25 / 0,03² = 1.068.")
    rn = CORR["replicas_necesarias_W_real"]
    w.p("Para el patrimonio real final hicimos un piloto de 200 réplicas con otra semilla. Pidiendo un error de 2 % del promedio, "
        "las réplicas necesarias son " + ", ".join(f"{c(rn[e])} para {e}" for e in EST) + ", todas por debajo de 1.068, así que "
        "ese número alcanza para los dos indicadores. Para comprobarlo corrimos E2 en el escenario base con 2.400 réplicas: la "
        "probabilidad de meta y su intervalo se estabilizan como dice la teoría (Figura 14-a).")
    w.tabla("14", "Precisión de la probabilidad de meta de E2 según el número de réplicas.",
            ["Réplicas", "P(meta)", "IC 95 %", "Semiancho", "Semiancho teórico máximo"],
            [[c(x.replicas), pc(x.p_meta), f"{pc(x.lo)} a {pc(x.hi)}", f"{c(x.semiancho * 100, 1)} puntos",
              f"{c(x.semiancho_teorico_max * 100, 1)} puntos"] for x in PREC.itertuples()])
    w.figura(FIG / "f08_precision.png", "14", "Probabilidad de meta de E2 (escenario base) e intervalo de 95 % con 100 a 2.400 réplicas.", 4.6)
    w.h("14.1 Semillas, independencia y condiciones iniciales", 2)
    w.p(f"La corrida parte de la semilla base {CORR['semilla_base']}, que queda guardada con los resultados junto con la versión del "
        f"modelo ({CORR['version_modelo']}) y los parámetros. La réplica r usa la semilla base + r, así que cualquier réplica se puede "
        "repetir sola, y el resultado es idéntico con uno o con varios procesos en paralelo (hay una prueba que lo verifica). Ninguna "
        "réplica usa resultados de otra. Todas parten del mismo capital, la misma meta, el estado Normal, TPM de "
        f"{c(P.f('tpm.inicial'), 2)} % e inflación de {c(P.f('inf.inicial'), 2)} %.")
    w.h("14.2 Números aleatorios comunes", 2)
    v = par("base", "E4", "E2", "W_real")
    w.p("Dentro de un escenario, la réplica r usa la misma semilla para las cuatro estrategias. Como cada proceso tiene su propio "
        "generador, las cuatro enfrentan la misma economía, y al restar sus resultados réplica a réplica se elimina la suerte. La "
        "varianza de la diferencia es Var(A) + Var(B) − 2Cov(A, B), y la covarianza es grande. Para la diferencia de patrimonio "
        f"real entre E4 y E2, la varianza baja {pc(v.reduccion_varianza, 0)} frente a comparar corridas independientes: el mismo "
        f"intervalo exigiría unas {c(1 / (1 - v.reduccion_varianza))} veces más réplicas sin esta técnica.")
    w.h("14.3 Intervalos de confianza", 2)
    w.vinetas([
        ("Medias", "(patrimonio, rentabilidad, meses de iliquidez): t de Student."),
        ("Proporciones", "(meta, pérdida real, iliquidez, venta urgente): intervalo de Wilson, que no se sale de 0 y 1."),
        ("Percentiles y CVaR", ": bootstrap con 1.000 remuestreos."),
        ("Comparaciones entre estrategias", ": media de las diferencias pareadas con intervalo t."),
    ])
    sig = PAR[PAR.significativa]
    m = len(PAR)
    orden = np.sort(PAR.p_valor.to_numpy())
    holm = 0
    for k, pv in enumerate(orden):
        if pv > 0.05 / (m - k):
            break
        holm += 1
    w.p(f"En total hacemos {c(m)} comparaciones (6 pares de estrategias, 4 indicadores y 3 escenarios). Con tantas, alguna "
        f"saldría significativa por azar. De las {c(len(sig))} diferencias significativas al 5 %, {c(holm)} lo siguen siendo con la "
        "corrección de Holm, que controla la probabilidad de cometer al menos un falso positivo en todo el conjunto; las "
        "conclusiones de la sección 18 se apoyan en estas.")


def seccion_15(w: Informe):
    w.h("15. Indicadores de desempeño")
    w.p("Cada réplica produce una fila con 66 columnas . Se reportan los indicadores mínimos de desempeño y tres más que responden "
        "preguntas del estudio; el resto queda en los archivos de resultados.")
    filas = [
        ["Patrimonio nominal y real final", "Inmueble neto de costos de venta; real = nominal / índice de precios"],
        ["Rentabilidad real", "Tasa anual compuesta del patrimonio real"],
        ["Variabilidad", "Desviación estándar y percentiles 5, 25, 50, 75 y 95"],
        ["Pérdida máxima", "Mayor caída del patrimonio real desde un máximo previo"],
        ["Pérdida real", "Probabilidad de terminar con patrimonio real menor que el capital"],
        ["Meta", "Probabilidad de que el patrimonio real final alcance la meta"],
        ["Momento de la meta", "Primer mes en que el patrimonio real alcanza la meta"],
        ["Iliquidez", "Probabilidad de algún episodio, número, duración y meses en iliquidez"],
        ["Agregado: CVaR 5 %", "Promedio del 5 % de réplicas con peor patrimonio final: qué tan malo es el peor caso"],
        ["Agregado: venta urgente", "Probabilidad de tener que vender el inmueble con descuento"],
        ["Agregado: costos totales", "Comisiones, impuestos, castigos y descuentos pagados en diez años, para juzgar a E4"],
    ]
    w.tabla("15", "Indicadores.", ["Indicador", "Definición"], filas, [2.2, 4.4])


def seccion_16(w: Informe):
    w.h("16. Verificación, validación y sensibilidad")
    w.h("16.1 Verificación", 2)
    w.p("La verificación responde si el programa hace lo que dice el modelo conceptual. Escribimos 59 pruebas automáticas (pytest) "
        "que corren en cada cambio del código en GitHub Actions. La mayoría usan un mundo quieto: rendimientos, inflación, costos "
        "y eventos aleatorios en cero, donde el resultado correcto se puede calcular a mano.")
    filas = [
        ["Eventos y reloj", "9", "La lista ordena por tiempo, prioridad y llegada; un evento cancelado no se procesa; el reloj nunca "
                               "retrocede y termina en H; hay exactamente H cierres de mes; los eventos periódicos caen en su calendario"],
        ["Cálculos financieros", "17", "Mundo quieto conserva el patrimonio en las cuatro estrategias; cuenta sola crece al interés "
                                       "compuesto; CDT paga TPM más margen; 4x1000 solo sobre lo que pasa la exención; fórmula de venta "
                                       "anticipada; tope del canon; costos de compra y venta; subir la constante del CDT sube el margen"],
        ["Generadores y parámetros", "12", "Filas de la matriz suman 1; frecuencias simuladas igualan la matriz; volatilidad del ETF por estado; "
                              "Jones-Faddy reproduce media y asimetría; inflación converge a su nivel; media de la vacancia; Poisson de liquidez; valores fuera de dominio se rechazan antes de simular"],
        ["Estrategias y liquidez", "14", "Pesos iniciales; E4 se mueve hacia el objetivo del estado; orden de cobertura; sin liquidez "
                                         "hay iliquidez y deuda; caída extrema del ETF; crisis persistente; estrés fija y libera la crisis"],
        ["Reproducibilidad", "7", "Misma semilla, mismo resultado; flujos por proceso independientes; las estrategias ven la misma "
                                  "economía; el resultado no depende del número de procesos; los números comunes reducen la varianza; Wilson queda entre 0 y 1"],
    ]
    w.tabla("16", "Pruebas automáticas por grupo.", ["Grupo", "Pruebas", "Qué verifican"], filas, [1.4, 0.6, 4.6])
    w.p("Las pruebas encontraron un error real durante el desarrollo: los residuos del margen del CDT se calculaban con la "
        "constante a que elegía el usuario, de modo que un cambio en a se cancelaba con el corrimiento de los residuos y no movía "
        "el margen. Lo notamos porque el tornado mostraba efecto cero para ese factor. Ahora los residuos salen de los valores "
        "estimados y una prueba lo vigila.")
    w.p("Además de las pruebas, el simulador registra cada evento de una réplica con su hora, número, "
        f"detalle, estado económico, cuenta y patrimonio. La réplica 0 de E4 en el escenario base tiene {c(len(TRAZA))} eventos sin "
        "contar arriendos ni cierres de mes; el Anexo B muestra los primeros y la Figura 16-a su trayectoria.")
    w.figura(FIG / "f09_replica_E4.png", "16", "Una réplica de E4 en el escenario base: composición del patrimonio, TPM e inflación, y estado económico.")
    w.h("16.2 Validación", 2)
    w.p("Validar es comprobar que el modelo se parece al sistema real. Corrimos 500 réplicas con los parámetros de referencia y "
        "comparamos lo simulado con los datos. La Tabla 16-b separa dos tipos de comparación. Las marcadas como \"reproduce "
        "insumos\" usan los mismos datos con que se estimó el proceso: muestran que el generador está bien programado, pero no "
        "son una prueba independiente del modelo. Las marcadas como \"contraste\" miden algo que el modelo no usó al "
        "estimarse, como el nivel que alcanza la TPM, la tasa real, las colas del ETF o la vacancia de otra fuente. La "
        "asimetría, la curtosis y el peor mes del ETF se comparan con la misma muestra con que se estimó, pero son rasgos que "
        "el modelo no ajustó: por eso los contamos como contraste, aunque más débil que uno con datos nuevos.")
    filas = []
    for x in VAL.itertuples():
        def fmt(v):
            try:
                return c(float(v), 3)
            except (TypeError, ValueError):
                return str(v).replace(".", ",")
        filas.append([x.proceso, x.medida, x.tipo.capitalize(), fmt(x.simulado), fmt(x.observado), x.referencia])
    w.tabla("16", "Simulado frente a observado. Corrida aparte de 500 réplicas de E3 con los parámetros de referencia; por eso "
                  "la fracción de tiempo desocupado puede diferir en unas décimas de la de la corrida principal.",
            ["Proceso", "Medida", "Tipo", "Simulado", "Observado", "Referencia"], filas,
            [0.9, 2.1, 0.9, 0.8, 0.8, 1.4], tam=8)
    w.p("Hay además tres validaciones fuera de muestra en la sección 7: la inflación y el IPVU pronosticados un año adelante "
        "caen en la banda de 90 % en 8 de 10 y 10 de 10 años de 2016 a 2025, y las probabilidades del logit de la Junta "
        "coinciden con las frecuencias observadas por tercios (Tabla 7.2-b). Lo que no hicimos es una validación del portafolio "
        "completo, por ejemplo correr una estrategia con los datos reales de 2011 a 2026 y compararla con lo que habría pasado; "
        "con una sola trayectoria histórica esa comparación no permite juzgar una distribución. La autocorrelación de la "
        "inflación sale más baja en la simulación porque se mide con 10 años por réplica, y con muestras tan cortas ese "
        "estimador se sesga hacia cero; con la serie larga el modelo reproduce 0,41.")
    tasa_real = VAL[VAL.medida.str.startswith("Tasa real")].iloc[0]
    vac = VAL[VAL.proceso == "Vacancia"].iloc[0]
    w.p("La mayoría de los procesos quedan cerca de sus datos: la mediana de la inflación, la media y la desviación del ETF, la "
        "proporción del tiempo en cada estado y la valorización. Tres diferencias merecen comentario.")
    tardia = VAL[VAL.medida.str.startswith("Tasa real media en los años")].iloc[0]
    taylor = re.search(r"pide (\d+,\d) puntos", str(tardia.lectura)).group(1)
    w.vinetas([
        ("Tasa real alta.", f"La TPM menos la inflación promedia {c(float(tasa_real.simulado), 1)} puntos en los diez años simulados y "
                            f"{c(float(tasa_real.observado), 1)} en la historia. La diferencia viene sobre todo del punto de partida: hoy la "
                            "TPM está en 12,25 % con inflación de 6,3 %, una tasa real de 6 puntos, y la Junta simulada la baja de a poco. "
                            f"En los años 6 a 10 la tasa real ya es {c(float(tardia.simulado), 1)} y sigue bajando hacia lo que pide la "
                            f"regla de Taylor con la inflación simulada ({taylor} puntos). Esto favorece a la renta fija, es decir a E1, y "
                            "lo tratamos como limitación."),
        ("ETF sin asimetría.", "La t de Student es simétrica y los datos tienen asimetría negativa (−1,12); el peor mes simulado en "
                               "diez años es más suave que marzo de 2020. La Jones-Faddy sí la reproduce y está disponible como "
                               f"alternativa; usarla cambia la probabilidad de meta entre {rango_sm('ETF')} puntos."),
        ("Vacancia conservadora.", f"El apartamento pasa {pc(r('base', 'E2', 'tiempo_vacante_media'), 1)} del tiempo desocupado frente a 0,6 % en "
                                   "Ciencuadras. Con la duración de la Lonja (121 días) en vez de la de Palomma, la probabilidad de "
                                   f"meta cambia menos de un punto."),
    ])
    w.h("16.3 Análisis de sensibilidad", 2)
    w.p(f"Movimos {c(TOR.factor.nunique())} factores, uno a la vez, a un valor bajo y uno alto, con 400 réplicas y las mismas "
        "semillas, y medimos el cambio en la probabilidad de meta. Los factores incluyen la inflación, la renta "
        "variable, las tasas, la valorización y la vacancia, y además las necesidades de liquidez, la probabilidad de crisis, la liquidez "
        "mínima, el descuento de la venta urgente, el castigo del CDT y la administración del edificio.")
    w.figura(FIG / "f06_tornado.png", "16", "Tornado de la probabilidad de meta para E2 y E3 (escenario base, 400 réplicas por punto).")
    top = TOR[TOR.estrategia == "E2"].sort_values("rango_p_meta", ascending=False).head(5)
    w.p("Para E2, los cinco factores con más efecto son: " + "; ".join(
        f"{x.factor.split(' (')[0][0].lower() + x.factor.split(' (')[0][1:]} ({c(x.bajo_p_meta * 100)} % a {c(x.alto_p_meta * 100)} %)" for x in top.itertuples())
        + ". La vacancia, el margen del CDT, la tasa de la cuenta y la probabilidad de crisis casi no mueven el resultado. La "
        "volatilidad del ETF tiene un efecto positivo pequeño: como la meta está por encima de la mediana, más dispersión deja más "
        "réplicas por encima de ella.")
    w.p("El tornado mueve un factor a la vez y supone que los efectos se suman. Para revisarlo cruzamos los dos factores que más "
        "pesan entre los que no son del inmueble, el nivel de la inflación y la frecuencia de las necesidades, en un diseño 2 × 2 "
        "con 400 réplicas por celda (Tabla 16-c). La interacción es el cambio en el efecto de las necesidades cuando la inflación "
        "pasa de baja a alta.")
    w.tabla("16", "Interacción entre el nivel de la inflación y la frecuencia de las necesidades: probabilidad de meta (%).",
            ["Estrategia", "Inflación 3,5 %, necesidades × 0,5", "3,5 %, × 1,5", "6,5 %, × 0,5", "6,5 %, × 1,5", "Interacción (puntos)"],
            [[x.estrategia, c(x.inf_baja_liq_baja * 100), c(x.inf_baja_liq_alta * 100), c(x.inf_alta_liq_baja * 100),
              c(x.inf_alta_liq_alta * 100), sg(x.interaccion * 100)] for x in INTER.itertuples()],
            [0.9, 1.3, 1.0, 1.0, 1.0, 1.2])
    x2 = INTER.set_index("estrategia").loc["E2"]
    suma = x2.inf_baja_liq_baja + (x2.inf_alta_liq_baja - x2.inf_baja_liq_baja) + (x2.inf_baja_liq_alta - x2.inf_baja_liq_baja)
    w.p(f"Los dos factores no se suman. Para E2, la inflación alta quita {c((x2.inf_baja_liq_baja - x2.inf_alta_liq_baja) * 100)} "
        f"puntos con pocas necesidades, y las necesidades frecuentes quitan {c((x2.inf_baja_liq_baja - x2.inf_baja_liq_alta) * 100)} "
        f"con inflación baja. Si los efectos se sumaran, con los dos juntos la probabilidad quedaría en {c(suma * 100)} %, que no "
        f"es posible; la simulada es {c(x2.inf_alta_liq_alta * 100)} %. Cuando la inflación es alta la meta ya es casi inalcanzable "
        "y las necesidades tienen poco que quitar. Por eso el tornado sirve para ordenar los factores uno por uno, pero no para "
        "sumar efectos de escenarios que combinan varios cambios; para eso están los escenarios adverso y de estrés.")
    w.p("También medimos qué pasa si cada proceso se simula con su modelo alternativo (Tabla 16-d).")
    filas = []
    for m, g in SM.groupby("modelo", sort=False):
        g = g.set_index("estrategia")
        filas.append([m] + [sg((g.loc[e, 'alt_p_meta'] - g.loc[e, 'ref_p_meta']) * 100) for e in EST])
    w.tabla("16", "Cambio en la probabilidad de meta (puntos) al usar el modelo alternativo.", ["Modelo alternativo"] + EST, filas,
            [3.4, 0.8, 0.8, 0.8, 0.8])
    nombres = {"CDT": "el CDT por fase", "Canon": "el canon que sube todo el IPC", "ETF": "la Jones-Faddy",
               "Junta: logit": "el logit de la Junta estimado solo con 2017 a 2026", "Vacancia": "la vacancia de 4 meses de la Lonja"}
    efecto = {k: max(abs(sm(k, e)) for e in EST) for k in nombres}
    lim = int(np.ceil(max(efecto.values())))
    w.p(f"Cinco alternativas mueven {c(lim)} puntos o menos y no cambian la conclusión principal: "
        + ", ".join(nombres[k] for k in nombres) + ". Tres sí cambian la conclusión. Con la tabla "
        "de la Junta la TPM llega al tope de 15 % en cerca de la mitad de las réplicas y la renta fija se vuelve muy rentable: E1 "
        f"gana {c(sm('Junta: tabla', 'E1'))} puntos y pasa a ser la mejor estrategia. Con el índice de vivienda nueva de Medellín "
        f"(IPVNBR) las estrategias con inmueble ganan entre {rango_sm('Valorización')} puntos. Y sin necesidades extraordinarias de "
        f"liquidez la probabilidad de meta de E2 a E4 sube entre {rango_sm('Sin necesidades')} puntos. Las dos primeras alternativas "
        "no reproducen mejor la historia: la tabla lleva la TPM a niveles que nunca tuvo y el IPVNBR mide vivienda nueva, no usada. "
        "La tercera muestra cuánto pesa un supuesto que no tiene datos.")


def seccion_17(w: Informe):
    w.h("17. Experimento capital-meta")
    caps_cm = sorted(CM.capital.unique())
    metas_cm = sorted(CM.meta.unique())
    lista_metas = ", ".join(c(m / 1e6) for m in metas_cm)
    w.p(f"Se corrieron las cuatro estrategias con {c(len(caps_cm))} capitales, de {c(caps_cm[0] / 1e6)} a {c(caps_cm[-1] / 1e6)} millones, y se "
        f"midió la probabilidad de llegar a {c(len(metas_cm))} metas reales ({lista_metas} millones), con 500 réplicas por celda, "
        "en los escenarios base y de estrés. Las dos mallas son múltiplos del capital y de la meta de la configuración de "
        "referencia, y en la aplicación se recalculan para cualquier otra. El inmueble cambia en proporción al capital.")
    w.figura(FIG / "f05_capital_meta.png", "17", "Probabilidad de alcanzar la meta según el capital inicial. Izquierda: E2 por capital y meta (escenario base). Derecha: "
                                                 "probabilidad de llegar a 1.000 millones según el capital, base (continua) y estrés (punteada).")
    filas = []
    for esc in ("base", "estres"):
        for m in (700e6, 800e6, 1000e6, 1200e6):
            fila = [ESC[esc].capitalize(), c(m / 1e6)]
            for e in EST:
                v = cap(esc, e, m)
                fila.append(c(v / 1e6) if v == v else "más de 1.000")
            filas.append(fila)
    w.tabla("17", "Capital inicial (millones) para tener 80 % de probabilidad de alcanzar cada meta real.",
            ["Escenario", "Meta (M)"] + EST, filas)
    saltos = [pm("base", e, k + 1e8) - pm("base", e, k) for e in ("E2", "E3") for k in (7e8, 8e8, 9e8)]
    w.p(f"Con el capital y la meta de referencia, llegar a la meta es poco probable para todas "
        f"({pc(pm('base', 'E3', 7e8), 0)} con E3 en esta corrida de 500 réplicas). La curva es empinada cerca del capital de referencia: "
        f"cada 100 millones adicionales suben la probabilidad entre {c(min(saltos) * 100)} y {c(max(saltos) * 100)} puntos con E2 y "
        f"E3. Con E1 hace falta más capital para cualquier meta: con {c(max(caps_cm) / 1e6)} millones llega a la meta de 1.000 millones reales en "
        f"{pc(pm('base', 'E1', 1e9), 0)} de las réplicas, frente a {pc(pm('base', 'E2', 1e9), 0)} con E2. Para 80 % de probabilidad "
        f"con E2 se necesitan unos {mm(cap('base', 'E2', 1e9))}. Esa cifra sale de interpolar una malla de 100 millones con 500 "
        "réplicas por punto, así que tiene su propio error: con los extremos del intervalo de 95 % de cada punto, el capital "
        f"necesario queda entre {mm(cap('base', 'E2', 1e9, 'capital_80_lo'))} y {mm(cap('base', 'E2', 1e9, 'capital_80_hi'))}. "
        "En el escenario de estrés ninguna estrategia llega a 80 % para esa meta dentro de la malla.")


def seccion_18(w: Informe):
    w.h("18. Resultados")
    filas = []
    for esc in ESC:
        for e in EST:
            x = R.loc[(esc, e)]
            filas.append([ESC[esc].capitalize(), e, f"{pc(x.p_meta)} ({c(x.p_meta_lo * 100, 1)} a {c(x.p_meta_hi * 100, 1)})",
                          c(x.W_real_p5 / 1e6), c(x.W_real_p25 / 1e6), c(x.W_real_p50 / 1e6), c(x.W_real_p75 / 1e6),
                          c(x.W_real_p95 / 1e6), c(x.cvar5_W_real / 1e6), pc(x.rent_real_media), pc(x.p_perdida_real, 0),
                          pc(x.caida_max_p50, 0), pc(x.p_iliquidez), pc(x.p_venta_urgente)])
    w.tabla("18", f"Indicadores principales ({c(N)} réplicas por fila). Patrimonio real final en millones de pesos de hoy.",
            ["Escenario", "Estr.", "P(meta) e IC 95 %", "P5", "P25", "P50", "P75", "P95", "CVaR 5 %", "Rent. real", "P(pérdida real)",
             "Caída máx. P50", "P(iliquidez)", "P(venta urgente)"], filas,
            [0.5, 0.33, 0.9, 0.36, 0.36, 0.36, 0.38, 0.4, 0.44, 0.45, 0.5, 0.45, 0.5, 0.5], tam=7)
    filas2 = []
    for esc in ESC:
        for e in EST:
            x = R.loc[(esc, e)]
            t_meta = "n. a." if x.t_meta_p50_anios != x.t_meta_p50_anios else c(x.t_meta_p50_anios, 1)
            filas2.append([ESC[esc].capitalize(), e, c(x.W_nominal_media / 1e6), c(x.W_nominal_p50 / 1e6), c(x.W_real_sd / 1e6),
                           pc(x.caida_max_p95, 0), t_meta, c(x.episodios_iliquidez_media, 3), c(x.duracion_iliquidez_media, 1),
                           c(x.meses_iliquidez_media, 2), pc(x.participacion_arriendo_media, 0)])
    w.tabla("18", "Otros indicadores. Patrimonio en millones de pesos corrientes (nominal) y de hoy (real); años hasta la meta entre las réplicas que la alcanzan.",
            ["Escenario", "Estr.", "Nominal medio", "Nominal P50", "Desv. real", "Pérdida máx. P95", "Años a la meta (P50)",
             "Episodios por réplica", "Meses por episodio", "Meses en iliquidez", "Arriendo en el retorno"], filas2,
            [0.55, 0.35, 0.6, 0.6, 0.55, 0.55, 0.6, 0.6, 0.55, 0.55, 0.6], tam=7)
    w.figura(FIG / "f01_abanico.png", "18", "Patrimonio real mes a mes: mediana (línea) y percentiles 25 a 75 (banda). La línea gris es la meta.")
    w.figura(FIG / "f02_cajas.png", "18", "Patrimonio real final: caja del percentil 25 al 75 y bigotes del 5 al 95. Punteada: capital inicial; rayada: meta.")
    w.figura(FIG / "f03_prob_meta.png", "18", "Probabilidad de alcanzar la meta con su intervalo de 95 %.", 5.5)
    w.figura(FIG / "f04_iliquidez.png", "18", "Porcentaje de réplicas en iliquidez cada mes.")
    w.h("18.1 Respuestas a las preguntas del estudio", 2)

    e32 = par("base", "E3", "E2", "W_real")
    w.p("**1. ¿Cómo cambia la distribución del patrimonio final entre estrategias? ",
        f"En el escenario base E2, E3 y E4 tienen medianas parecidas ({mm(r('base', 'E2', 'W_real_p50'))}, {mm(r('base', 'E3', 'W_real_p50'))} "
        f"y {mm(r('base', 'E4', 'W_real_p50'))}) y E1 queda unos {c((r('base', 'E2', 'W_real_p50') - r('base', 'E1', 'W_real_p50')) / 1e6)} "
        f"millones abajo ({mm(r('base', 'E1', 'W_real_p50'))}). Lo que cambia es la forma. E3 tiene la cola derecha más larga "
        f"(percentil 95 de {mm(r('base', 'E3', 'W_real_p95'))}) y también una cola izquierda peor que E2: su CVaR 5 % es "
        f"{mm(r('base', 'E3', 'cvar5_W_real'))} frente a {mm(r('base', 'E2', 'cvar5_W_real'))}. Las diferencias pareadas entre "
        f"estrategias son distintas de cero: E3 supera a E2 en {mm(e32.diferencia)} en promedio (IC 95 % de {c(e32.lo / 1e6)} a "
        f"{c(e32.hi / 1e6)} millones).")
    w.p("**2. ¿Qué efecto tiene la inflación sobre las conclusiones? ",
        "Es la que más cambia la respuesta. Si la meta se midiera en pesos corrientes, las estrategias con inmueble la alcanzarían "
        f"en al menos {c(np.floor(min(NR.loc[('base', e), 'p_meta_nominal'] for e in ('E2', 'E3', 'E4')) * 100))} % de las réplicas; en pesos de hoy, "
        f"en menos de {pc(r('base', 'E3', 'p_meta') + 0.01, 0)}. En el escenario adverso, con inflación alrededor de 6,5 %, los precios "
        f"se duplican en la mediana ({pc(NR.loc[('adverso', 'E2'), 'inflacion_acumulada_p50'], 0)} de inflación acumulada) y la "
        "probabilidad de meta cae a menos de 1 % para todas. En el tornado, pasar la inflación de largo plazo de 3,5 % a 6,5 % mueve "
        f"la probabilidad de meta de E2 en {c(tor('Nivel de la inflación', 'E2') * 100)} puntos. La inflación también afecta más a "
        f"E1: pierde poder adquisitivo en {pc(r('base', 'E1', 'p_perdida_real'), 0)} de las réplicas del escenario base.")
    w.p("**3. ¿Qué estrategias presentan mayor exposición a problemas de liquidez? ",
        f"E3, con diferencia. En el escenario base {pc(r('base', 'E3', 'p_iliquidez'))} de sus réplicas tienen algún episodio, contra "
        f"{pc(r('base', 'E2', 'p_iliquidez'))} de E2. En el adverso la cifra sube a {pc(r('adverso', 'E3', 'p_iliquidez'))} para E3, "
        f"{pc(r('adverso', 'E2', 'p_iliquidez'))} para E2 y {pc(r('adverso', 'E4', 'p_iliquidez'))} para E4, y en "
        f"{pc(r('adverso', 'E3', 'p_venta_urgente'))} de las réplicas E3 vende el apartamento con descuento. E1 casi nunca tiene "
        "problemas porque 80 % de su capital está en cuenta y CDT. La exposición depende de cuánto efectivo y ETF quedan después de "
        f"comprar el inmueble: E3 empieza con {pc(1 - P.lista('E3.pesos')[3], 0)} del capital en activos líquidos y E2 con "
        f"{pc(1 - P.lista('E2.pesos')[3], 0)}.")
    w.p("**4. ¿Cuál es la probabilidad de alcanzar diferentes metas para distintos capitales iniciales? ",
        "La sección 17 lo responde con la matriz completa. En resumen, con E2 en el escenario base la probabilidad de llegar a 1.000 "
        f"millones pasa de {pc(pm('base', 'E2', 5e8))} con 500 millones a {pc(pm('base', 'E2', 9e8), 0)} con 900 millones, y 80 % de "
        f"probabilidad pide unos {mm(cap('base', 'E2', 1e9))}. Para una meta de 800 millones bastan {mm(cap('base', 'E2', 800e6))}.")
    m_adv, m_est = mejor("adverso", "W_real_p50"), mejor("estres", "W_real_p50")
    w.p("**5. ¿Cómo cambia el desempeño bajo escenarios adversos y de estrés? ",
        f"En el adverso todas pierden poder adquisitivo (rentabilidad real media entre {pc(min(r('adverso', e, 'rent_real_media') for e in EST))} "
        f"y {pc(max(r('adverso', e, 'rent_real_media') for e in EST))} al año) y el orden se invierte. E3, la de mayor probabilidad de "
        f"meta en el base, termina con la mediana más baja ({mm(r('adverso', 'E3', 'W_real_p50'))}) y la peor caída máxima. "
        f"{NOM[m_adv]} termina con la más alta ({mm(r('adverso', m_adv, 'W_real_p50'))}): casi todo su capital está en cuenta y CDT, "
        "cuyas tasas suben con la inflación, y no tiene inmueble que vender con descuento. En estrés la crisis forzada baja la "
        f"probabilidad de meta a menos de {pc(max(r('estres', e, 'p_meta') for e in EST) + 0.01, 0)} y la probabilidad de terminar con "
        f"menos poder adquisitivo que al inicio queda entre {pc(min(r('estres', e, 'p_perdida_real') for e in EST), 0)} y "
        f"{pc(max(r('estres', e, 'p_perdida_real') for e in EST), 0)}. También ahí {NOM[m_est]} conserva el patrimonio mediano más alto "
        f"({mm(r('estres', m_est, 'W_real_p50'))}).")
    w.p("**6. ¿Qué variables y supuestos explican principalmente el riesgo? ",
        f"En orden, {principales()} (sección 16.3). Entre las decisiones de modelación, la regla de la Junta es la que más pesa, seguida del "
        "índice de vivienda que se use para la valorización. La vacancia y la probabilidad de crisis mueven poco el resultado a "
        "diez años.")
    w.p("**7. ¿Qué diferencias aparecen entre analizar rentabilidad esperada y la distribución completa? ",
        f"Por probabilidad de meta, E3 es la mejor en el base ({pc(r('base', 'E3', 'p_meta'), 0)} contra {pc(r('base', 'E2', 'p_meta'), 0)} "
        f"de E2), y su rentabilidad real media es casi la misma de E2 ({pc(r('base', 'E3', 'rent_real_media'))} y "
        f"{pc(r('base', 'E2', 'rent_real_media'))} al año). Mirando la distribución, E3 tiene peor CVaR, más iliquidez, más ventas "
        f"urgentes y una caída máxima mediana mayor ({pc(r('base', 'E3', 'caida_max_p50'), 0)} contra {pc(r('base', 'E2', 'caida_max_p50'), 0)}). "
        "En el escenario adverso su cola izquierda es la peor de las cuatro. Quien solo mire el promedio o la probabilidad de meta "
        "elegiría E3 sin ver que en los años malos es la que más sufre.")
    v = par("base", "E4", "E2", "W_real")
    w.p("**8. ¿Qué efecto tienen las reglas dinámicas de decisión? ",
        f"En este modelo, ninguno favorable. E4 termina {mm(-v.diferencia)} por debajo de E2 en promedio en el escenario base (IC 95 % "
        f"de {c(-v.hi / 1e6, 1)} a {c(-v.lo / 1e6, 1)} millones" + ("" if holm("base", "E4", "E2", "W_real") else
        "; con la corrección de Holm la diferencia no es significativa") + "), con la misma probabilidad de meta y la misma "
        "iliquidez. Paga más "
        f"costos ({mm(r('base', 'E4', 'costos_totales_media'))} contra {mm(r('base', 'E2', 'costos_totales_media'))}) porque rebalancea "
        f"en promedio {c(r('base', 'E4', 'n_rebalanceos_media'), 1)} veces. La regla saca ETF cuando la economía entra en crisis, "
        "pero el modelo no le da al ETF un rendimiento medio más bajo en crisis (los datos no lo sostienen), solo más volatilidad, "
        "y la revisión es anual mientras una crisis dura unos cuatro trimestres. Frente a E3, E4 tiene "
        f"{c(-par('estres', 'E4', 'E3', 'algun_episodio').diferencia * 100, 1)} puntos menos de probabilidad de iliquidez en el estrés, "
        "pero eso viene más de su reparto inicial que de la regla.")
    w.p("**9. ¿Cuáles son las principales limitaciones del modelo y cuándo no deberían usarse sus resultados? ",
        "Las detallamos en la sección de limitaciones. Las más importantes: la tasa real simulada es más alta que la histórica, "
        "sobre todo en los primeros años; la comparación entre E1 y las estrategias con inmueble depende de la regla de la Junta; "
        "las necesidades de liquidez son un supuesto sin datos que pesa mucho; y no modelamos el impuesto de renta sobre intereses "
        "y arriendos ni la ganancia ocasional al vender el inmueble.")


def seccion_19(w: Informe):
    w.h("19. Entregables y cómo reproducir el estudio")
    w.vinetas([
        ("Informe técnico:", "este documento (docs/informe_final.docx y su PDF)."),
        ("Diagrama de eventos:", "sección 10 y docs/diagramas en PDF."),
        ("Simulador:", "repositorio con el motor (src/investment_sim), la aplicación web (app/), las pruebas (tests/), los "
                       "scripts de estimación y de corrida (tools/) y los resultados (outputs/resultados)."),
        ("Excel maestro:", "data/master_datos.xlsx, con los parámetros que se pueden cambiar, el modelo elegido para cada "
                           "proceso, los datos y las pruebas de la sección 7, la validación y los resultados."),
        ("Registro de decisiones:", "docs/decision_log.md."),
    ])
    w.p("Para reproducir todo: instalar las dependencias (pip install -r requirements-dev.txt), correr pytest para las pruebas, "
        "python tools/correr_experimentos.py 12 para la corrida completa (unos 10 minutos con 12 procesos) y streamlit run "
        "app/main.py para abrir la aplicación. Con la semilla base y los parámetros de referencia se obtienen exactamente las "
        "tablas de este informe.")
    w.p("La aplicación permite cambiar sin tocar código el capital, el horizonte, la meta, la liquidez mínima, las necesidades "
        "de liquidez, los pesos de las estrategias y las reglas de E4. Los procesos aleatorios quedan fijos con el modelo elegido "
        "en la sección 7. Tiene siete páginas: inicio, inversionista, modelos, estrategias, experimento, resultados, y sensibilidad "
        "y capital-meta. Los resultados, las réplicas y los parámetros usados se descargan en CSV, Excel y JSON.")


def discusion_y_cierre(w: Informe, refs_equipo: list[str]):
    w.h("Discusión")
    w.p("La pregunta del trabajo no era qué activo rinde más, y los resultados muestran por qué. Las cuatro estrategias crecen en "
        f"pesos corrientes: la mediana nominal de E2 en el escenario base es {mm(r('base', 'E2', 'W_nominal_p50'))}. Lo que decide la "
        "meta es cuánto de ese crecimiento sobrevive a la inflación y cuánto se pierde vendiendo a destiempo.")
    w.p("El inmueble es la pieza que más separa a las estrategias. En el escenario base se valoriza en promedio "
        f"{pc(np.exp(P.lista('val.ar1')[0]) - 1, 1)} nominal al año, y el arriendo neto es "
        f"{pc(r('base', 'E2', 'participacion_arriendo_media'), 0)} de los ingresos corrientes (arriendo más intereses) de E2 y "
        f"{pc(r('base', 'E3', 'participacion_arriendo_media'), 0)} de los de E3. Por eso E2, E3 y E4 superan a E1 en el escenario base. Pero el inmueble no se puede vender por partes ni "
        "rápido. Cuando llegan varios imprevistos seguidos, una estrategia con poco efectivo vende ETF, luego CDT con castigo y al "
        "final el apartamento con 10 % de descuento y costos de casi 5 %. En el escenario adverso eso le pasa a E3 en "
        f"{pc(r('adverso', 'E3', 'p_venta_urgente'), 0)} de las réplicas, y esa venta explica que su cola izquierda sea la peor. La "
        "relación es directa: menos efectivo inicial, más ventas forzadas cuando la economía empeora.")
    w.p("E1 muestra el otro lado. Nunca vende de urgencia y es la que mejor aguanta los escenarios adverso y de estrés, porque "
        "sus tasas suben cuando sube la inflación. Pero en el escenario base pierde poder adquisitivo en "
        f"{pc(r('base', 'E1', 'p_perdida_real'), 0)} de las réplicas, porque la TPM sigue a la inflación con retraso. Su resultado "
        "mejora mucho si la tasa real se mantiene alta, como pasa con la tabla de la Junta de 7.2. E1 le sirve a quien le importa "
        "más no vender nunca y protegerse de un mal escenario que llegar a la meta.")
    w.p("Sobre E4, el resultado no dice que las reglas dinámicas no sirvan. Dice que esta regla, con revisión anual y en un modelo "
        "donde el estado no cambia el rendimiento medio del ETF, no tiene de dónde sacar ventaja y sí paga costos. Una regla que "
        "reaccione a la liquidez disponible o que proteja el inmueble de una venta urgente atacaría el riesgo que de verdad pesa.")
    w.p("Todo esto es evidencia de simulación, no un pronóstico. Las probabilidades describen lo que pasaría si el futuro se "
        "pareciera a los procesos que estimamos con datos de 2001 a 2026. Los intervalos de confianza miden el error de simular "
        "con un número finito de réplicas; no miden el error de haber elegido mal un modelo, que puede ser más grande, y por eso "
        "hicimos la sensibilidad a los modelos.")

    w.h("Conclusiones")
    w.vinetas([
        f"Con la configuración de referencia, la meta tiene a lo sumo {pc(r('base', mejor('base', 'p_meta'), 'p_meta'), 0)} "
        f"de probabilidad en el escenario base. Para 80 % con E2 hace falta un capital de unos {mm(cap('base', 'E2', 1e9))}, es decir "
        f"{c(cap('base', 'E2', 1e9) / P.f('inv.capital'), 2)} veces el de referencia.",
        "Las estrategias con inmueble (E2, E3, E4) dan más patrimonio que E1 en el escenario base. E2 tiene una probabilidad de "
        f"iliquidez parecida a la de E1 ({pc(r('base', 'E2', 'p_iliquidez'), 0)} frente a {pc(r('base', 'E1', 'p_iliquidez'), 0)}), "
        f"aunque termina con menos dinero líquido ({mm(r('base', 'E2', 'liquido_final_real_media'))} frente a "
        f"{mm(r('base', 'E1', 'liquido_final_real_media'))} reales). E3 tiene la mayor probabilidad de meta a cambio de más riesgo de iliquidez y de ventas urgentes.",
        "La inflación es la variable que más pesa en las conclusiones. Medir la meta en pesos corrientes daría una imagen falsa: "
        f"con E2 la probabilidad nominal es {pc(NR.loc[('base', 'E2'), 'p_meta_nominal'], 0)} y la real {pc(r('base', 'E2', 'p_meta'), 0)}.",
        "En los escenarios adverso y de estrés el orden se invierte: E1 termina con el patrimonio mediano más alto y E3 con el más bajo.",
        "La regla dinámica de E4 no mejora a E2 en este modelo; cuesta más y no gana rentabilidad.",
        "Dos supuestos pueden cambiar la recomendación: la regla de la Junta (con la tabla de 7.2, E1 pasa a ser la mejor) y las "
        "necesidades extraordinarias de liquidez, que no tienen datos y bajan mucho la probabilidad de meta de todas las estrategias.",
    ])

    w.h("Limitaciones")
    real = VAL[VAL.medida.str.startswith("Tasa real media en los años")].iloc[0]
    w.vinetas([
        ("Tasa real alta.", f"La TPM simulada menos la inflación promedia {c(TASA_REAL, 1)} puntos en los diez años, contra 1,2 en la "
                            f"historia, y {c(float(real.simulado), 1)} en los años 6 a 10. Viene de partir de una tasa real de 6 puntos. "
                            "Favorece la renta fija, es decir a E1."),
        ("Necesidades de liquidez.", "Frecuencia y monto son un supuesto. La IEFIC del DANE solo sirve de referencia para la "
                                     "frecuencia, y para el monto no hay datos. Es uno de los factores que más mueve el resultado."),
        ("Rendimiento del ETF por estado.", "Solo la volatilidad depende del estado. Si el rendimiento medio también dependiera, la "
                                            "regla de E4 podría servir."),
        ("ETF con poca historia.", "La muestra tiene 15 años y casi todo el rendimiento positivo está en 2024 a 2026. La media que "
                                   "usamos depende mucho de esos años, y la t no reproduce la asimetría de los datos."),
        ("Impuestos.", "Modelamos la retención en la fuente (4 % en el CDT y 7 % en la cuenta) y el 4x1000, y tratamos la retención "
                       "como costo final aunque es un anticipo. No modelamos el impuesto de renta sobre intereses y arriendos ni la "
                       "ganancia ocasional al vender el inmueble, que reducirían el patrimonio de todas las estrategias."),
        ("Inmueble.", "Es un solo apartamento indivisible. Su valorización viene del IPVU, que agrega nueve ciudades y suaviza los "
                      "precios porque usa avalúos, y el apartamento no tiene riesgo propio (un daño grande, un barrio que se "
                      "deteriora)."),
        ("Pocos datos de crisis.", "La matriz de Markov se estimó con 12 trimestres de crisis y tres episodios. La duración media "
                                   "de una crisis simulada se parece a la observada, pero en diez años una réplica puede "
                                   "encadenar crisis más largas que las tres reales."),
        ("Vacancia.", "Más alta que la reportada por el mercado, a propósito."),
    ])
    w.p("Los resultados no deberían usarse para recomendar una inversión concreta a una persona, para horizontes mucho más largos "
        "que los datos (más de 15 años), para economías con inflación muy distinta a la colombiana de 2001 a 2026, ni para elegir un "
        "CDT o un fondo particular. Sirven para comparar estrategias bajo los mismos supuestos y para ver qué supuestos pesan.")

    formato_apa.referencias(w)


def anexos(w: Informe):
    w.salto()
    w.h("Anexo A. Parámetros del simulador")
    w.p("Valores de referencia del simulador. Los del inversionista, el experimento, las estrategias y las necesidades de "
        "liquidez se pueden cambiar; los demás están estimados con datos y son fijos. Adverso y Estrés vacíos significan el mismo "
        "valor de la base.")
    filas = []
    for f in (f for f in FILAS if not f[1].startswith("rep.")):
        def v(x):
            if x is None:
                return ""
            if isinstance(x, float):
                return c(x, 4) if abs(x) < 10 else c(x, 2)
            if isinstance(x, int):
                return str(x) if x > 10_000_000 else c(x)
            return str(x)
        filas.append([f[0], f[3], v(f[4]), v(f[5]), v(f[6]), f[7], f[8]])
    w.tabla("A", "Parámetros.", ["ID", "Parámetro", "Base", "Adverso", "Estrés", "Unidad", "Estado"], filas,
            [0.55, 2.05, 0.85, 0.7, 0.7, 1.0, 0.85], tam=7)
    w.h("Anexo B. Traza de eventos de una réplica")
    w.p("Primeros eventos de la réplica 0 de E4 en el escenario base (semilla 20261107), sin los cierres de mes ni los arriendos. "
        "La traza completa está en outputs/resultados/traza_E4_replica0.csv.")
    t = TRAZA[~TRAZA.evento.isin(["Cierre de mes", "Pago de arriendo"])].head(40)
    w.tabla("B", "Traza de eventos.", ["Mes", "N.º", "Evento", "Qué pasó", "Estado", "Cuenta (M)", "Patrimonio (M)"],
            [[c(x.t, 2), x.n, x.evento, str(x.detalle).translate(str.maketrans(".,", ",.")), x.estado, c(x.cuenta / 1e6, 1), c(x.W / 1e6, 1)]
             for x in t.itertuples()], [0.5, 0.4, 1.3, 2.5, 0.7, 0.7, 0.8], tam=7)


def main(origen: Path) -> None:
    doc = Document(str(origen))
    portada_y_resumen(doc)
    notas_7(doc)
    w = Informe(doc)
    refs_equipo = actualizar(w, P, c, pc)
    seccion_8(w)
    seccion_9(w)
    seccion_10(w)
    seccion_11(w)
    seccion_12(w)
    seccion_13(w)
    seccion_14(w)
    seccion_15(w)
    seccion_16(w)
    seccion_17(w)
    seccion_18(w)
    seccion_19(w)
    discusion_y_cierre(w, refs_equipo)
    anexos(w)
    limpieza_texto.aplicar(doc)
    faltan = formato_apa.citas(doc)
    info = formato_apa.aplicar(doc)
    formato_apa.notas_sueltas(doc)
    print("llamadas agregadas:", formato_apa.llamadas(doc))
    formato_apa.signos(doc)
    formato_apa.saltos(doc)
    formato_apa.estilo(doc)
    formato_apa.pagina(doc)
    doc.core_properties.title = "Simulación de estrategias de inversión bajo incertidumbre"
    doc.core_properties.author = "Daniel Giraldo Arenas, Allison Correa y Samuel Alexis Vásquez Muñoz"
    rev = formato_apa.revisar(doc)
    print("citas no aplicadas:", faltan)
    print(f"{info['tablas']} tablas, {info['figuras']} figuras; sin título: {info['sin_titulo']}")
    print("referencias sin cita:", rev["referencias_sin_cita"])
    DESTINO.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(DESTINO))
    print("ok", DESTINO)


if __name__ == "__main__":
    main(Path(sys.argv[1]) if len(sys.argv) > 1 else ORIGEN)

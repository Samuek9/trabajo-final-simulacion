"""Arma docs/informe_final.docx a partir del Word de los puntos 1 a 7 (que no se modifica)
y agrega los puntos 8 a 19, la discusión, las conclusiones, las limitaciones y los anexos.

  python tools/construir_informe.py [ruta del Word de los puntos 1 a 7]

Todos los números salen de outputs/resultados y de la hoja PARAMETROS; si se vuelve a
correr el experimento, basta volver a correr este script. Después, tools/finalizar_word.ps1
actualiza la tabla de contenido y exporta el PDF con Word.
"""
from __future__ import annotations

import json
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
N = CORR["replicas"]


def c(x: float, dec: int = 0) -> str:
    """Número con formato colombiano."""
    return f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pc(x: float, dec: int = 1) -> str:
    return f"{c(x * 100, dec)} %"


def mm(x: float) -> str:
    return f"{c(x / 1e6)} millones"


def r(esc, e, col):
    return float(R.loc[(esc, e), col])


def par(esc, a, b, var):
    q = PAR[(PAR.escenario == esc) & (PAR.a == a) & (PAR.b == b) & (PAR.variable == var)].iloc[0]
    return q


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
    poner(buscar(doc, "Archivo de soporte"), "Archivos de soporte: data/master_datos.xlsx (Excel maestro con parámetros, fuentes, "
          "datos, validación y resultados), el repositorio del simulador (código, datos, pruebas y resultados de la corrida) y "
          "docs/decision_log.md (decisiones de modelación).")
    poner(buscar(doc, "Este documento cubre los puntos 1 a 7"),
          "Este documento cubre los puntos 1 a 19 del enunciado. Las secciones 1 a 7 plantean el problema y estudian cada proceso "
          "aleatorio con datos; las secciones 8 a 15 describen el modelo de eventos discretos, las estrategias, los escenarios y el "
          "diseño del experimento; las secciones 16 a 18 muestran la verificación, la validación, la sensibilidad y los resultados. "
          "Todo número que aparece aquí sale de los datos citados o de la corrida del simulador; nada se estimó a ojo.")

    ancla = buscar(doc, "1. Contexto", True)
    e_mejor = mejor("base", "p_meta")
    textos = [
        ["Simulamos diez años de un inversionista de Medellín con 700 millones de pesos que quiere terminar con al menos 1.000 "
         "millones en pesos de hoy. El modelo es de eventos discretos: un reloj que salta de evento en evento (reuniones de la "
         "Junta, vencimientos de CDT, arriendos, vacancias, gastos imprevistos, cambios del ciclo económico) y un estado que cada "
         f"evento actualiza. Corrimos {c(N)} réplicas de cada una de las cuatro estrategias en tres escenarios, con la misma economía "
         "para las cuatro estrategias dentro de cada réplica."],
        [f"En el escenario base ninguna estrategia vuelve probable la meta. {NOM[e_mejor]} la alcanza en {pc(r('base', e_mejor, 'p_meta'))} "
         f"de las réplicas; E2 en {pc(r('base', 'E2', 'p_meta'))}, E4 en {pc(r('base', 'E4', 'p_meta'))} y E1 en "
         f"{pc(r('base', 'E1', 'p_meta'))}. El obstáculo es la inflación, no la rentabilidad nominal: medida en pesos corrientes, "
         f"E2 pasa de 1.000 millones en {pc(NR.loc[('base', 'E2'), 'p_meta_nominal'], 0)} de las réplicas. La inflación acumulada "
         f"mediana de los diez años es {pc(NR.loc[('base', 'E2'), 'inflacion_acumulada_p50'], 0)}."],
        [f"Más patrimonio esperado viene con menos liquidez. E3 tiene el patrimonio real mediano más alto ({mm(r('base', 'E3', 'W_real_p50'))}), "
         f"pero en el escenario adverso {pc(r('adverso', 'E3', 'p_iliquidez'), 0)} de sus réplicas tienen al menos un episodio de "
         f"iliquidez y {pc(r('adverso', 'E3', 'p_venta_urgente'), 0)} terminan vendiendo el apartamento de urgencia. E1 casi nunca se "
         f"queda sin efectivo, pero en {pc(r('base', 'E1', 'p_perdida_real'), 0)} de las réplicas pierde poder adquisitivo."],
        [f"La estrategia dinámica E4 no mejora a E2, de la que parte: termina en promedio {mm(-par('base', 'E4', 'E2', 'W_real').diferencia)} "
         "por debajo en el escenario base, y la diferencia en la probabilidad de meta no es distinta de cero. En el modelo el estado "
         "de la economía cambia la volatilidad del ETF pero no su rendimiento medio, así que moverse a CDT en crisis reduce el riesgo "
         "sin ganar rentabilidad, y cada ajuste cuesta comisión y 4x1000."],
        [f"Con 700 millones la meta es una apuesta. Para tener 80 % de probabilidad de llegar a 1.000 millones reales hacen falta cerca de "
         f"{mm(cap('base', 'E2', 1e9))} con E2 o E3 en el escenario base. Las variables que más mueven el resultado son el nivel de la "
         "inflación, la valorización del inmueble, el rendimiento medio del ETF y el monto y la frecuencia de los gastos imprevistos. "
         "Dos elecciones de modelo también cambian la conclusión (la regla de la Junta y el modelo de valorización) y se discuten en "
         "la sección de limitaciones."],
    ]
    antes(doc, ancla, ["Resumen ejecutivo"], "Heading 1")
    for t in textos:
        antes(doc, ancla, t)
    salto = antes(doc, ancla, [""])
    salto.add_run().add_break(WD_BREAK.PAGE)


def notas_7(doc):
    pi = markov.estacionaria(markov.matriz(P))
    notas = {
        "7.2.4 Dependencia y pruebas": (
            "Uso en el simulador. ",
            "La tabla anterior es la versión estimada con 46 reuniones. Al simular diez años con ella la TPM tiende a subir hasta "
            "el tope de 15 % en cerca de la mitad de las réplicas, algo que no ocurre en la historia. Por eso el simulador usa como "
            "referencia un logit ordenado estimado con las 266 decisiones de la Junta entre 2001 y 2026, con dos variables: la "
            "tendencia de la inflación en tres meses y la brecha frente a una regla de Taylor, TPM − (1 + π + 0,5(π − 3)). La "
            f"tabla queda como alternativa en la aplicación. El detalle está en la hoja TPM_Logit y en la sección 16.2."),
        "7.4.10 Dependencia": (
            "Uso en el simulador. ",
            "El simulador usa M5 como referencia, con los residuos observados sorteados por bootstrap. El margen por fase de la TPM "
            "queda como alternativa y su efecto sobre las conclusiones se mide en la sección 16.3 (cambio en la "
            "probabilidad de meta, un punto o menos)."),
        "7.5.10 Dependencia con otras variables": (
            "Uso en el simulador. ",
            "Aunque la Jones-Faddy ajusta mejor la distribución sin estados, el simulador usa como referencia la t de Student con "
            f"ν = {c(P.f('etf.nu'))} y volatilidad mensual que depende del estado económico ({pc(P.f('etf.vol_E'), 1)} en expansión y "
            f"normal, {pc(P.f('etf.vol_C'), 1)} en crisis), estimada por el equipo con la hoja MARKOV_ETF. La razón es el punto 8: "
            "el enunciado pide que los estados muevan el ETF, y la Jones-Faddy no tiene estados. La Jones-Faddy queda como "
            "alternativa; cambiarla mueve la probabilidad de meta en menos de dos puntos (sección 16.3)."),
        "7.6.10 Dependencia": (
            "Uso en el simulador. ",
            "El simulador usa como referencia la valorización normal por estado del análisis del equipo (IPVNBR de Medellín): media "
            f"logarítmica anual de {c(P.f('val.mu_E'), 3)} en expansión y {c(P.f('val.mu_N'), 3)} en normal y crisis. El AR(1) del IPVU "
            "nacional queda como alternativa. Esta elección sí pesa: con el AR(1) la probabilidad de meta de E2 a E4 baja entre "
            f"{rango_sm('Valorización')} puntos (sección 16.3), porque su media de largo plazo (7,55 %) es menor que la de Medellín."),
        "7.7.10 Dependencia": (
            "Uso en el simulador. ",
            "La validación de la sección 16.2 muestra que con estos valores el apartamento pasa desocupado cerca de 3 % del tiempo, "
            "más que el 0,6 % que reporta Ciencuadras para Medellín. Lo dejamos así a propósito: preferimos un modelo que exagere "
            "la vacancia a uno que la ignore."),
        "Referencias de las secciones 5.4, 7.6 y 7.7": (
            "Uso en el simulador. ",
            "El simulador usa la clasificación del equipo (sección 8), que difiere de la de esta sección en dos cosas: el periodo "
            "(2006 a 2026, 82 trimestres) y el umbral de crisis (crecimiento menor que 1 %, para incluir el estancamiento). Con 0 % "
            "solo quedaban 4 trimestres de crisis en ese periodo. A largo plazo la cadena pasa "
            f"{pc(pi[0], 0)} del tiempo en expansión, {pc(pi[1], 0)} en normal y {pc(pi[2], 0)} en crisis."),
    }
    for inicio, (lead, texto) in notas.items():
        antes(doc, buscar(doc, inicio, True), ["**" + lead, texto])


# ---------------------------------------------------------------- secciones nuevas
def seccion_8(w: Informe):
    Pm = markov.matriz(P)
    pi = markov.estacionaria(Pm)
    dur = 1 / (1 - np.diag(Pm))
    w.salto()
    w.h("8. Dependencia entre variables y estados económicos")
    w.p("No suponemos que las variables sean independientes. El modelo tiene dos estructuras de dependencia. La primera es una "
        "cadena causal: la inflación mueve la decisión de la Junta, la TPM fija la tasa de la cuenta (α × TPM) y la tasa a la que "
        "se abre cada CDT, y la inflación también reajusta el canon y las cuotas del inmueble. La segunda es una cadena de Markov "
        "de tres estados económicos que cambia al mismo tiempo la volatilidad del ETF, la valorización del inmueble, la vacancia y "
        "la frecuencia de los gastos imprevistos.")
    w.h("8.1 Estados y matriz de transición", 2)
    w.p("Clasificamos cada trimestre de 2006 a 2026 según el crecimiento anual del PIB real (DANE): expansión si pasa de 4 %, "
        "crisis si es menor que 1 % y normal en otro caso. Usamos tres estados y no cinco porque Colombia tuvo pocas crisis en el "
        "periodo; con cinco, algunos estados tendrían dos o tres datos. De 82 trimestres, 30 fueron de expansión, 40 normales y 12 "
        "de crisis. La matriz sale de contar las transiciones de cada estado a cada otro y dividir por el total de la fila.")
    filas = [[s] + [c(Pm[i, j], 3) for j in range(3)] + [c(dur[i], 1), pc(pi[i], 0)] for i, s in enumerate(markov.ESTADOS)]
    w.tabla("8", "Matriz de transición trimestral y lo que implica.", ["Desde \\ hacia", "Expansión", "Normal", "Crisis",
            "Duración media (trimestres)", "Tiempo a largo plazo"], filas)
    w.p("Los estados son persistentes y la economía nunca pasó directo de expansión a crisis ni de crisis a expansión. La "
        "simulación empieza en Normal, que es el estado del segundo trimestre de 2026 (crecimiento de 3,5 %), y cada tres meses "
        "sortea el estado siguiente con la fila del estado vigente.")
    w.h("8.2 Qué cambia con el estado", 2)
    filas = [
        ["ETF: volatilidad mensual", pc(P.f("etf.vol_E"), 1), pc(P.f("etf.vol_N"), 1), pc(P.f("etf.vol_C"), 1), "MARKOV_ETF"],
        ["Inmueble: valorización log anual", c(P.f("val.mu_E"), 3), c(P.f("val.mu_N"), 3), c(P.f("val.mu_C"), 3), "MARKOV_Vivienda"],
        ["Vacancia: probabilidad de no renovar", pc(P.f("vac.p_E"), 0), pc(P.f("vac.p_N"), 0), pc(P.f("vac.p_C"), 0), "Supuesto"],
        ["Vacancia: duración media (meses)", c(P.f("vac.media_E"), 1), c(P.f("vac.media_N"), 1), c(P.f("vac.media_C"), 1), "7.7"],
        ["Gastos imprevistos: uno cada ... meses", c(1 / P.f("liq.tasa_E")), c(1 / P.f("liq.tasa_N")), c(1 / P.f("liq.tasa_C")), "Supuesto"],
    ]
    w.tabla("8", "Efecto del estado económico sobre los demás procesos.", ["Variable", "Expansión", "Normal", "Crisis", "Fuente"], filas)
    w.p("Así, cuando llega una crisis el ETF se vuelve más riesgoso, la vivienda deja de valorizarse como en la expansión, "
        "arrendar cuesta más y los imprevistos llegan el doble de seguido, todo en la misma réplica. Una cosa no cambia con el "
        "estado: el rendimiento medio del ETF. Con 182 meses de datos la diferencia de medias entre estados no fue significativa, "
        "y preferimos no inventarla. Esa decisión explica buena parte del resultado de E4 (sección 18).")


def seccion_9(w: Informe):
    w.h("9. Modelo de simulación de eventos discretos")
    w.p("El simulador está escrito en Python y no usa ninguna librería de simulación: el reloj, la lista de eventos y los "
        "manejadores son nuestros, para que cada regla sea visible. El tiempo se mide en meses, empieza en 0 y termina en H = "
        f"{P.i('inv.horizonte')}. La lista de eventos futuros es un montículo (heapq) ordenado por tiempo, luego por la prioridad "
        "del diagrama y luego por orden de llegada; así dos eventos en el mismo instante siempre se procesan en el mismo orden. En "
        "cada paso el motor saca el evento más próximo, adelanta el reloj, ejecuta el manejador (que cambia el estado y programa "
        "eventos futuros) y vuelve a empezar. Un evento cancelado, como el próximo arriendo de un inmueble que se vendió, se marca "
        "y se descarta al salir de la lista.")
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
        ["10", "Reparación extraordinaria", "Apagado (sin datos)", "Poisson"],
        ["11", "Necesidad de liquidez", "Sortea el monto y llama a la cobertura (12); programa la siguiente", "Poisson según el estado"],
        ["12", "Cobertura de liquidez", "Obtiene el efectivo: cuenta, ETF, venta anticipada de CDT, venta urgente del inmueble", "Al instante, la llama otro evento"],
        ["13", "Venta del inmueble", "Recibe el producto de la venta urgente, con descuento y costos", "3 a 6 meses después de decidirla"],
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
    w.p("Para algunos procesos el estudio de la sección 7 y el análisis del equipo propusieron modelos distintos. El simulador "
        "trae los dos; el de referencia es el que se usa en los resultados y el otro se puede elegir en la aplicación.")
    filas = [
        ["Estados económicos", "Markov del equipo (umbral 1 % y 4 %)", "No"],
        ["Junta", "Logit ordenado con tendencia y brecha de Taylor", "Tabla por tendencia (7.2)"],
        ["Margen del CDT", "M5 con residuos observados (7.4)", "Normal por fase de la TPM"],
        ["ETF", "t de Student con volatilidad por estado (equipo)", "Jones-Faddy sin estados (7.5)"],
        ["Valorización", "Normal por estado, IPVNBR Medellín (equipo)", "AR(1) del IPVU (7.6)"],
        ["Reajuste del canon", "Modelo DANE con tope en el IPC previo (5.4.3)", "Sube todo el IPC (diagrama del equipo)"],
    ]
    w.tabla("9", "Modelo de referencia y alternativa de cada proceso.", ["Proceso", "Referencia", "Alternativa"], filas)


def seccion_10(w: Informe):
    w.h("10. Diagrama de eventos")
    w.p("El diagrama de eventos y el diagrama de flujo los construyó el equipo antes de programar. El motor sigue su numeración: "
        "cada manejador lleva el número de su evento, y la prioridad de los empates es la de la página 1 del diagrama de flujo. "
        "Los originales en PDF están en docs/diagramas del repositorio y en la página Metodología de la aplicación. Sobre el "
        "diagrama original hicimos dos agregados que no cambian su lógica: los eventos 16 y 17 de la crisis forzada del escenario "
        "de estrés, y que la cobertura de liquidez (12) se ejecute en el mismo instante del evento que la necesita, en vez de "
        "programarse en la lista.")
    if (DIAG / "diagrama_eventos.png").exists():
        w.figura(DIAG / "diagrama_eventos.png", "10", "Diagrama de eventos del equipo. La imagen tiene resolución para ampliarla en el PDF; "
                                                      "el original vectorial está en docs/diagramas/diagrama_eventos.pdf.")
    for k in range(1, 6):
        f = DIAG / f"flujo_pagina_{k}.png"
        if f.exists():
            w.figura(f, "10", f"Diagrama de flujo, página {k} de 5.")


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
    w.p("Los gastos imprevistos (una enfermedad, una reparación de la casa propia, una ayuda a la familia) llegan como un proceso "
        f"de Poisson: en expansión y normal, en promedio uno cada {c(1 / P.f('liq.tasa_N'))} meses; en crisis, uno cada "
        f"{c(1 / P.f('liq.tasa_C'))}. El monto es lognormal con mediana de {mm(P.f('liq.mediana'))} en pesos de hoy y σ = "
        f"{c(P.f('liq.sigma'), 1)} en el logaritmo, y se indexa con el índice de precios simulado. Con esos valores, una de cada "
        "diez necesidades pasa de 55 millones. En diez años el inversionista enfrenta en promedio "
        f"{c(r('base', 'E1', 'n_necesidades_media'), 1)} necesidades en el escenario base.")
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
        "Como el estado sigue moviendo el ETF, la valorización, la vacancia y la liquidez, y la inflación sigue moviendo la TPM, un "
        "cambio en un escenario se transmite a las demás variables.")
    nivel_b = inflation.nivel_largo_plazo(P.f("inf.c"), P.f("inf.phi"))
    nivel_a = inflation.nivel_largo_plazo(PA.f("inf.c"), PA.f("inf.phi"))
    filas = [
        ["Estados económicos", "Matriz estimada", f"Normal → Crisis {c(PA.f('markov.N_C'), 2)}; Crisis → Normal {c(PA.f('markov.C_N'), 2)}",
         "Estado fijo en Crisis durante la crisis forzada"],
        ["Inflación", f"Nivel normal {c(nivel_b, 1)} %", f"Nivel normal {c(nivel_a, 1)} %", "Base"],
        ["TPM (Junta)", "Logit", "Logit (sube más por la inflación)", "Bajadas de 0,5 a 2 puntos por reunión"],
        ["Cuenta", f"{c(P.f('cta.alfa'), 3)} × TPM", f"{c(PA.f('cta.alfa'), 2)} × TPM", "Base"],
        ["ETF", f"μ = {pc(P.f('etf.mu'), 2)} mensual", "μ = 0", "Caída de 25 % al empezar la crisis"],
        ["Valorización", "Según el estado", "2 puntos menos en cada estado", "−5 % anual durante la crisis"],
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


def seccion_15(w: Informe):
    w.h("15. Indicadores de desempeño")
    w.p("Cada réplica produce una fila con 66 columnas (hoja DICCIONARIO del Excel). Estos son los indicadores del enunciado y "
        "los que agregamos:")
    filas = [
        ["Patrimonio nominal y real final", "Inmueble neto de costos de venta; real = nominal / índice de precios"],
        ["Rentabilidad real", "Tasa anual compuesta del patrimonio real"],
        ["Variabilidad", "Desviación estándar y percentiles 5, 25, 50, 75 y 95"],
        ["Pérdida máxima", "Mayor caída del patrimonio real desde un máximo previo"],
        ["Pérdida real", "Probabilidad de terminar con patrimonio real menor que el capital"],
        ["Meta", "Probabilidad de que el patrimonio real final alcance la meta"],
        ["Momento de la meta", "Primer mes en que el patrimonio real alcanza la meta"],
        ["Iliquidez", "Probabilidad de algún episodio, número, duración y meses en iliquidez"],
        ["Agregado: CVaR 5 %", "Promedio del 5 % de réplicas con peor patrimonio final"],
        ["Agregado: venta urgente", "Probabilidad de tener que vender el inmueble con descuento"],
        ["Agregado: costos totales", "Comisiones, impuestos, castigos y descuentos pagados en diez años"],
        ["Agregado: peso del arriendo", "Arriendo neto sobre la ganancia total"],
        ["Agregado: liquidez final", "Cuenta, CDT y ETF al final, en pesos de hoy"],
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
    w.p("Además de las pruebas, la página Trazabilidad de la aplicación corre una réplica y muestra cada evento con su hora, número, "
        f"detalle, estado económico, cuenta y patrimonio. La réplica 0 de E4 en el escenario base tiene {c(len(TRAZA))} eventos sin "
        "contar arriendos ni cierres de mes; el Anexo B muestra los primeros y la Figura 16-a su trayectoria.")
    w.figura(FIG / "f09_replica_E4.png", "16", "Una réplica de E4 en el escenario base: composición del patrimonio, TPM e inflación, y estado económico.")
    w.h("16.2 Validación", 2)
    w.p("Validar es comprobar que el modelo se parece al sistema real. Corrimos 500 réplicas con los parámetros de referencia y "
        "comparamos lo simulado con los datos de cada proceso.")
    filas = []
    for x in VAL.itertuples():
        def fmt(v):
            try:
                return c(float(v), 3)
            except (TypeError, ValueError):
                return str(v).replace(".", ",")
        filas.append([x.proceso, x.medida, fmt(x.simulado), fmt(x.observado), x.referencia])
    w.tabla("16", "Simulado frente a observado.", ["Proceso", "Medida", "Simulado", "Observado", "Referencia"], filas,
            [1.0, 2.4, 0.9, 0.9, 1.6], tam=8.5)
    tasa_real = VAL[VAL.medida.str.startswith("Tasa real")].iloc[0]
    vac = VAL[VAL.proceso == "Vacancia"].iloc[0]
    w.p("La mayoría de los procesos quedan cerca de sus datos: la mediana de la inflación, la media y la desviación del ETF, la "
        "proporción del tiempo en cada estado y la valorización. Tres diferencias merecen comentario.")
    w.vinetas([
        ("Tasa real alta.", f"La TPM menos la inflación promedia {c(float(tasa_real.simulado), 1)} puntos en la simulación y "
                            f"{c(float(tasa_real.observado), 1)} en la historia. El modelo parte de una tasa real de unos 6 puntos (TPM de "
                            "12,25 % con inflación de 6,3 %) y la Junta simulada la baja despacio. Esto favorece a la renta fija, es decir "
                            "a E1, y lo tratamos como limitación."),
        ("ETF sin asimetría.", "La t de Student es simétrica y los datos tienen asimetría negativa (−1,12); el peor mes simulado en "
                               "diez años es más suave que marzo de 2020. La Jones-Faddy sí la reproduce y está disponible como alternativa."),
        ("Vacancia conservadora.", f"El apartamento pasa {pc(float(vac.simulado), 1)} del tiempo desocupado frente a 0,6 % en Ciencuadras."),
    ])
    w.h("16.3 Análisis de sensibilidad", 2)
    w.p("Movimos 14 factores, uno a la vez, a un valor bajo y uno alto, con 400 réplicas y las mismas semillas, y medimos el "
        "cambio en la probabilidad de meta. Los factores cubren lo que pide el enunciado (inflación, renta variable, tasas, "
        "valorización y vacancia) y además la liquidez y la probabilidad de crisis.")
    w.figura(FIG / "f06_tornado.png", "16", "Tornado de la probabilidad de meta para E2 y E3 (escenario base, 400 réplicas por punto).")
    top = TOR[TOR.estrategia == "E2"].sort_values("rango_p_meta", ascending=False).head(5)
    w.p("Para E2, los cinco factores con más efecto son: " + "; ".join(
        f"{x.factor.split(' (')[0].lower()} ({c(x.bajo_p_meta * 100)} % a {c(x.alto_p_meta * 100)} %)" for x in top.itertuples())
        + ". La vacancia, la volatilidad del ETF y el margen del CDT casi no mueven el resultado. El riesgo de no llegar a la meta "
        "lo explican el poder adquisitivo (inflación) y el crecimiento de los dos activos de riesgo, más el tamaño de los imprevistos.")
    w.p("También medimos qué pasa si cada proceso se simula con su modelo alternativo (Tabla 16-c).")
    filas = []
    for m, g in SM.groupby("modelo", sort=False):
        g = g.set_index("estrategia")
        filas.append([m] + [f"{(g.loc[e, 'alt_p_meta'] - g.loc[e, 'ref_p_meta']) * 100:+.1f}".replace(".", ",") for e in EST])
    w.tabla("16", "Cambio en la probabilidad de meta (puntos) al usar el modelo alternativo.", ["Modelo alternativo"] + EST, filas,
            [3.4, 0.8, 0.8, 0.8, 0.8])
    w.p("Tres alternativas mueven menos de tres puntos. Dos sí cambian la conclusión. Con la tabla de la Junta la TPM sube hasta el "
        f"tope y la renta fija se vuelve muy rentable: E1 gana {c(sm('Junta', 'E1'))} puntos y pasa a ser la mejor estrategia. Con el "
        f"AR(1) del IPVU las estrategias con inmueble pierden entre {rango_sm('Valorización')} puntos. Ninguna de las dos alternativas reproduce mejor la historia que la "
        "referencia (la primera lleva la TPM a niveles que nunca tuvo; la segunda usa precios nacionales para un apartamento en "
        "Medellín), pero muestran que la comparación entre E1 y las estrategias con inmueble depende de esos dos supuestos.")


def seccion_17(w: Informe):
    w.h("17. Experimento capital-meta")
    w.p("Corrimos las cuatro estrategias con capitales de 400 a 1.000 millones y medimos la probabilidad de llegar a metas reales "
        "de 700, 800, 1.000, 1.200 y 1.500 millones, con 500 réplicas por celda, en los escenarios base y de estrés. El inmueble "
        "cambia en proporción al capital.")
    w.figura(FIG / "f05_capital_meta.png", "17", "Izquierda: probabilidad de meta de E2 por capital y meta (escenario base). Derecha: "
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
    w.p(f"Con el capital actual de 700 millones, la meta de 1.000 millones es casi una moneda al aire para E2 y E3 "
        f"({pc(r('base', 'E2', 'p_meta'), 0)} y {pc(r('base', 'E3', 'p_meta'), 0)}). La curva es empinada entre 600 y 900 millones: cada 100 "
        f"millones adicionales suben la probabilidad entre {c(min(pm('base', e, k + 1e8) - pm('base', e, k) for e in ('E2', 'E3') for k in (6e8, 7e8, 8e8)) * 100)} "
        f"y {c(max(pm('base', e, k + 1e8) - pm('base', e, k) for e in ('E2', 'E3') for k in (6e8, 7e8, 8e8)) * 100)} puntos. Con E1 hace falta más capital para cualquier meta: para "
        f"1.000 millones con 80 % se necesitan {mm(cap('base', 'E1', 1e9))}, frente a {mm(cap('base', 'E2', 1e9))} con E2. En el "
        "escenario de estrés ninguna estrategia llega a 80 % para la meta de 1.000 millones con menos de 1.000 millones de capital, "
        "es decir, habría que empezar casi con la meta.")


def seccion_18(w: Informe):
    w.h("18. Resultados")
    filas = []
    for esc in ESC:
        for e in EST:
            x = R.loc[(esc, e)]
            filas.append([ESC[esc].capitalize(), e, f"{pc(x.p_meta)} ({c(x.p_meta_lo * 100, 1)} a {c(x.p_meta_hi * 100, 1)})",
                          c(x.W_real_p5 / 1e6), c(x.W_real_p50 / 1e6), c(x.W_real_p95 / 1e6), c(x.cvar5_W_real / 1e6),
                          pc(x.rent_real_media), pc(x.p_perdida_real, 0), pc(x.caida_max_p50, 0), pc(x.p_iliquidez), pc(x.p_venta_urgente)])
    w.tabla("18", f"Indicadores principales ({c(N)} réplicas por fila). Patrimonio en millones de pesos de hoy.",
            ["Escenario", "Estr.", "P(meta) e IC 95 %", "P5", "P50", "P95", "CVaR 5 %", "Rent. real", "P(pérdida real)",
             "Caída máx. P50", "P(iliquidez)", "P(venta urgente)"], filas,
            [0.55, 0.35, 0.95, 0.4, 0.4, 0.45, 0.48, 0.48, 0.55, 0.48, 0.55, 0.55], tam=7.5)
    w.figura(FIG / "f01_abanico.png", "18", "Patrimonio real mes a mes: mediana (línea) y percentiles 25 a 75 (banda). La línea gris es la meta.")
    w.figura(FIG / "f02_cajas.png", "18", "Patrimonio real final: caja del percentil 25 al 75 y bigotes del 5 al 95. Punteada: capital inicial; rayada: meta.")
    w.figura(FIG / "f03_prob_meta.png", "18", "Probabilidad de alcanzar la meta con su intervalo de 95 %.", 5.5)
    w.figura(FIG / "f04_iliquidez.png", "18", "Porcentaje de réplicas en iliquidez cada mes.")
    w.h("18.1 Respuestas a las preguntas del estudio", 2)

    w.p("**1. ¿Cómo cambia la distribución del patrimonio final entre estrategias? ",
        f"En el escenario base E2, E3 y E4 tienen medianas parecidas ({mm(r('base', 'E2', 'W_real_p50'))}, {mm(r('base', 'E3', 'W_real_p50'))} "
        f"y {mm(r('base', 'E4', 'W_real_p50'))}) y E1 queda unos 120 millones abajo ({mm(r('base', 'E1', 'W_real_p50'))}). Lo que cambia es la "
        f"forma. E3 tiene la cola derecha más larga (percentil 95 de {mm(r('base', 'E3', 'W_real_p95'))}) y también una cola izquierda peor "
        f"que E2: su CVaR 5 % es {mm(r('base', 'E3', 'cvar5_W_real'))} frente a {mm(r('base', 'E2', 'cvar5_W_real'))} de E2. Las diferencias "
        f"pareadas son todas distintas de cero: E3 supera a E2 en {mm(par('base', 'E3', 'E2', 'W_real').diferencia)} en promedio "
        f"(IC 95 % de {c(par('base', 'E3', 'E2', 'W_real').lo / 1e6)} a {c(par('base', 'E3', 'E2', 'W_real').hi / 1e6)} millones).")
    w.p("**2. ¿Qué efecto tiene la inflación sobre las conclusiones? ",
        "Es la variable que más cambia la respuesta. Si la meta se midiera en pesos corrientes, todas las estrategias con inmueble "
        f"la alcanzarían en más de {pc(min(NR.loc[('base', e), 'p_meta_nominal'] for e in ('E2', 'E3', 'E4')), 0)} de las réplicas; en pesos "
        f"de hoy, en menos de {pc(r('base', 'E3', 'p_meta') + 0.01, 0)}. En el escenario adverso, con inflación alrededor de 6,5 %, los precios "
        f"se duplican en la mediana ({pc(NR.loc[('adverso', 'E2'), 'inflacion_acumulada_p50'], 0)} de inflación acumulada) y la probabilidad "
        "de meta cae a menos de 1 % para todas. En el tornado, pasar la inflación de largo plazo de 3,5 % a 6,5 % mueve la probabilidad "
        "de meta de E2 en más de 40 puntos. La inflación también ordena las estrategias: E1, la más expuesta, pierde poder adquisitivo en "
        f"{pc(r('base', 'E1', 'p_perdida_real'), 0)} de las réplicas del escenario base.")
    w.p("**3. ¿Qué estrategias presentan mayor exposición a problemas de liquidez? ",
        f"E3, con diferencia. En el escenario base {pc(r('base', 'E3', 'p_iliquidez'))} de sus réplicas tienen algún episodio, contra "
        f"{pc(r('base', 'E2', 'p_iliquidez'))} de E2. En el adverso la cifra sube a {pc(r('adverso', 'E3', 'p_iliquidez'))} para E3, "
        f"{pc(r('adverso', 'E2', 'p_iliquidez'))} para E2 y {pc(r('adverso', 'E4', 'p_iliquidez'))} para E4, y en {pc(r('adverso', 'E3', 'p_venta_urgente'))} "
        "de las réplicas E3 vende el apartamento con descuento. E1 casi nunca tiene problemas porque 80 % de su capital está en cuenta y "
        f"CDT. La exposición depende de cuánto efectivo y ETF quedan después de comprar el inmueble: E3 empieza con {pc(1 - P.lista('E3.pesos')[3], 0)} del capital "
        f"en activos líquidos, E2 con {pc(1 - P.lista('E2.pesos')[3], 0)}.")
    w.p("**4. ¿Cuál es la probabilidad de alcanzar diferentes metas para distintos capitales iniciales? ",
        "La sección 17 lo responde con la matriz completa. En resumen, con E2 en el escenario base la probabilidad de llegar a 1.000 "
        f"millones pasa de {pc(pm('base', 'E2', 5e8))} con 500 millones a {pc(pm('base', 'E2', 9e8), 0)} con 900 millones, y 80 % de probabilidad pide unos "
        f"{mm(cap('base', 'E2', 1e9))}. Para una meta de 800 millones bastan {mm(cap('base', 'E2', 800e6))}.")
    w.p("**5. ¿Cómo cambia el desempeño bajo escenarios adversos y de estrés? ",
        f"En el adverso todas pierden poder adquisitivo en la mediana (rentabilidad real media entre {pc(min(r('adverso', e, 'rent_real_media') for e in EST))} "
        f"y {pc(max(r('adverso', e, 'rent_real_media') for e in EST))} al año) y el orden se invierte: E3, la mejor en el base, termina con "
        f"la mediana más baja ({mm(r('adverso', 'E3', 'W_real_p50'))}) y la peor caída máxima. En estrés la crisis forzada baja la "
        f"probabilidad de meta a menos de {pc(max(r('estres', e, 'p_meta') for e in EST) + 0.01, 0)} y la probabilidad de terminar con menos "
        f"poder adquisitivo que al inicio queda entre {pc(min(r('estres', e, 'p_perdida_real') for e in EST), 0)} y "
        f"{pc(max(r('estres', e, 'p_perdida_real') for e in EST), 0)}. E2 es la que mejor resiste el estrés en patrimonio mediano.")
    w.p("**6. ¿Qué variables y supuestos explican principalmente el riesgo? ",
        "El nivel de la inflación, la valorización del inmueble, el rendimiento medio del ETF y el monto y la frecuencia de los gastos "
        "imprevistos (sección 16.3). Entre los supuestos de modelación, la regla de la Junta y el modelo de valorización son los que "
        "más pesan. La vacancia y la volatilidad del ETF, que intuitivamente parecerían importantes, mueven poco el resultado a diez años.")
    w.p("**7. ¿Qué diferencias aparecen entre analizar rentabilidad esperada y la distribución completa? ",
        f"Por rentabilidad real media, E3 es la mejor en el base ({pc(r('base', 'E3', 'rent_real_media'))} anual contra {pc(r('base', 'E2', 'rent_real_media'))} "
        f"de E2). Mirando la distribución, E3 tiene peor CVaR, más iliquidez, más ventas urgentes y una caída máxima mediana mayor "
        f"({pc(r('base', 'E3', 'caida_max_p50'), 0)} contra {pc(r('base', 'E2', 'caida_max_p50'), 0)}). En el escenario adverso su ventaja "
        "desaparece y su cola izquierda es la peor de las cuatro. Un inversionista que solo mire el promedio elegiría E3 sin ver que en "
        "los años malos es la que más sufre.")
    v = par("base", "E4", "E2", "W_real")
    w.p("**8. ¿Qué efecto tienen las reglas dinámicas de decisión? ",
        f"En este modelo, ninguno favorable. E4 termina {mm(-v.diferencia)} por debajo de E2 en promedio en el escenario base (IC 95 % de "
        f"{c(-v.hi / 1e6, 1)} a {c(-v.lo / 1e6, 1)} millones), con la misma probabilidad de meta y la misma iliquidez. Paga más costos "
        f"({mm(r('base', 'E4', 'costos_totales_media'))} contra {mm(r('base', 'E2', 'costos_totales_media'))}) porque rebalancea en promedio "
        f"{c(r('base', 'E4', 'n_rebalanceos_media'), 1)} veces. La regla saca ETF cuando la economía entra en crisis, pero el modelo no "
        "le da al ETF un rendimiento medio más bajo en crisis (los datos no lo sostienen), solo más volatilidad, y la revisión es "
        "anual mientras una crisis dura unos 4 trimestres. La regla reduce algo el riesgo en estrés frente a E3 "
        f"({c(-par('estres', 'E4', 'E3', 'algun_episodio').diferencia * 100, 1)} puntos menos de probabilidad de iliquidez), pero eso se debe más a su reparto inicial que a la regla.")
    w.p("**9. ¿Cuáles son las principales limitaciones del modelo y cuándo no deberían usarse sus resultados? ",
        "Las detallamos en la sección de limitaciones. Las más importantes: la tasa real simulada es más alta que la histórica, el "
        "resultado entre E1 y las estrategias con inmueble depende del modelo de la Junta y del de valorización, las necesidades de "
        "liquidez son un supuesto sin datos, y no modelamos impuestos sobre el arriendo ni sobre la ganancia al vender.")


def seccion_19(w: Informe):
    w.h("19. Entregables y cómo reproducir el estudio")
    w.vinetas([
        ("Informe técnico:", "este documento (docs/informe_final.docx y su PDF)."),
        ("Diagrama de eventos:", "sección 10 y docs/diagramas en PDF."),
        ("Simulador:", "repositorio con el motor (src/investment_sim), la aplicación web (app/), las pruebas (tests/), los "
                       "scripts de estimación y de corrida (tools/) y los resultados (outputs/resultados)."),
        ("Excel maestro:", "data/master_datos.xlsx con 20 hojas: parámetros, fuentes, diccionario, distribuciones, decisiones, "
                           "los datos de los puntos 7 y 8, validación y resultados."),
        ("Registro de decisiones:", "docs/decision_log.md."),
    ])
    w.p("Para reproducir todo: instalar las dependencias (pip install -r requirements-dev.txt), correr pytest para las pruebas, "
        "python tools/correr_experimentos.py 8 para la corrida completa (unos 5 minutos con 8 procesos) y streamlit run "
        "app/main.py para abrir la aplicación. Con la semilla base y los parámetros de referencia se obtienen exactamente las "
        "tablas de este informe.")
    w.p("La aplicación permite cambiar sin tocar código el capital, el horizonte, la meta, la liquidez mínima, los pesos de las "
        "estrategias, las reglas de E4, cada parámetro de los modelos y el modelo de cada proceso. Debajo de cada control muestra el "
        "valor de referencia, la fuente y si el parámetro es estimado, supuesto o pendiente. Tiene nueve páginas: inicio, "
        "inversionista, modelos, estrategias, experimento, resultados, sensibilidad y capital-meta, trazabilidad de una réplica y "
        "metodología. Los resultados, las réplicas y los parámetros usados se descargan en CSV, Excel y JSON.")


def discusion_y_cierre(w: Informe):
    w.h("Discusión")
    w.p("La pregunta del trabajo no era qué activo rinde más. Los resultados muestran por qué. Las cuatro estrategias crecen en "
        f"pesos corrientes: la mediana nominal de E2 en el escenario base es {mm(r('base', 'E2', 'W_nominal_p50'))}. Lo que decide la meta "
        "es cuánto de ese crecimiento sobrevive a la inflación y cuánto se pierde vendiendo a destiempo.")
    w.p("El inmueble es la pieza que más separa a las estrategias. En el escenario base aporta valorización nominal de "
        f"{pc(np.exp(P.f('val.mu_N')) - 1, 1)} a {pc(np.exp(P.f('val.mu_E')) - 1, 1)} por año y un arriendo neto que explica {pc(r('base', 'E2', 'participacion_arriendo_media'), 0)} de la ganancia de E2 y "
        f"{pc(r('base', 'E3', 'participacion_arriendo_media'), 0)} de la de E3. Por eso E2, E3 y E4 superan a E1. Pero el inmueble no "
        "se puede vender por partes ni rápido. Cuando llegan varios imprevistos seguidos, una estrategia con poco efectivo vende ETF, "
        "luego CDT con castigo, y al final el apartamento con 10 % de descuento y costos de casi 5 %. En el escenario adverso eso le "
        f"pasa a E3 en {pc(r('adverso', 'E3', 'p_venta_urgente'), 0)} de las réplicas, y esa venta es la razón de que su cola izquierda "
        "sea la peor. La relación causa y efecto es directa: menos efectivo inicial, más ventas forzadas cuando la economía empeora.")
    w.p("E1 muestra el otro lado. Nunca vende de urgencia, pero su rendimiento depende de la TPM, y la TPM sigue a la inflación con "
        "retraso. Cuando la inflación sube, E1 pierde poder adquisitivo antes de que la Junta reaccione. Su resultado mejora mucho si "
        "la tasa real se mantiene alta, como ocurre con la tabla de 7.2, y empeora con una tasa real como la histórica. Una conclusión "
        "honesta es que E1 conviene solo a quien le importa más no vender nunca que llegar a la meta.")
    w.p("Sobre E4, el resultado no dice que las reglas dinámicas no sirvan. Dice que esta regla, con revisión anual y en un modelo "
        "donde el estado no cambia el rendimiento medio del ETF, no tiene de dónde sacar ventaja y sí paga costos. Una regla que "
        "reaccione a la liquidez disponible o que proteja el inmueble de una venta urgente atacaría el riesgo que de verdad pesa.")
    w.p("Todo esto es evidencia de simulación, no un pronóstico. Las probabilidades describen lo que pasaría si el futuro se "
        "pareciera a los procesos que estimamos con datos de 2001 a 2026. Los intervalos de confianza miden el error de simular con un "
        "número finito de réplicas; no miden el error de haber elegido mal un modelo, que es más grande y por eso hicimos la "
        "sensibilidad a los modelos.")

    w.h("Conclusiones")
    w.vinetas([
        f"Con 700 millones y diez años, la meta de 1.000 millones reales tiene a lo sumo {pc(r('base', mejor('base', 'p_meta'), 'p_meta'), 0)} de "
        f"probabilidad en el escenario base. Para 80 % hacen falta unos {mm(cap('base', 'E2', 1e9))} de capital.",
        "Las estrategias con inmueble (E2, E3, E4) dan más patrimonio que E1 en el escenario base, y E2 lo hace con casi la misma "
        "liquidez que E1. E3 gana algo más en promedio a cambio de bastante más riesgo de iliquidez y de ventas urgentes.",
        "La inflación es la variable que más pesa en las conclusiones. Medir la meta en pesos corrientes daría una imagen falsa: "
        f"con E2 la probabilidad nominal es {pc(NR.loc[('base', 'E2'), 'p_meta_nominal'], 0)} y la real {pc(r('base', 'E2', 'p_meta'), 0)}.",
        "En escenarios adversos el orden cambia: la estrategia más agresiva pasa de primera a última y la moderada resiste mejor.",
        "La regla dinámica de E4 no mejora a E2 en este modelo; cuesta más y no gana rentabilidad.",
        "Los resultados son sensibles a dos decisiones de modelación (la regla de la Junta y la valorización), que el equipo debe "
        "tener presentes al defender cualquier recomendación.",
    ])

    w.h("Limitaciones")
    w.vinetas([
        ("Tasa real alta.", f"La TPM simulada menos la inflación promedia {c(TASA_REAL, 1)} puntos, contra 1,2 en la historia. Esto favorece "
                            "la renta fija. El logit ordenado corrige la deriva de la tabla de 7.2 pero no del todo."),
        ("Necesidades de liquidez.", "Frecuencia y monto son un supuesto del equipo. No encontramos datos de gastos imprevistos de "
                                     "hogares con este patrimonio, y es uno de los factores que más mueve el resultado."),
        ("Rendimiento del ETF por estado.", "Solo la volatilidad depende del estado. Si el rendimiento medio también dependiera, la "
                                            "regla de E4 podría servir."),
        ("Impuestos.", "Modelamos retención sobre intereses y 4x1000. No modelamos impuesto de renta sobre arriendos ni ganancia "
                       "ocasional al vender el inmueble, que reducirían el patrimonio de E2, E3 y E4."),
        ("Inmueble.", "Es un solo apartamento indivisible; su valorización viene del índice de vivienda nueva de Medellín aplicado a "
                      "un apartamento usado. Las reparaciones extraordinarias están apagadas porque no tenemos datos."),
        ("Pocos datos de crisis.", "La matriz de Markov se estimó con 12 trimestres de crisis y solo tres episodios."),
        ("Vacancia.", "Más alta que la reportada por el mercado, a propósito."),
    ])
    w.p("Los resultados no deberían usarse para recomendar una inversión concreta a una persona, para horizontes mucho más largos "
        "que los datos (más de 15 años), para economías con inflación muy distinta a la colombiana de 2001 a 2026, ni para elegir un "
        "CDT o un fondo particular. Sirven para comparar estrategias bajo los mismos supuestos y para ver qué supuestos pesan.")

    w.h("Referencias adicionales")
    for ref in [
        "Banco de la República. (2026). Comunicados de las decisiones de política monetaria de la Junta Directiva, 2001 a 2026. https://www.banrep.gov.co.",
        "Banks, J., Carson, J. S., Nelson, B. L. y Nicol, D. M. (2010). Discrete-event system simulation (5.ª ed.). Pearson.",
        "Harris, C. R. et al. (2020). Array programming with NumPy. Nature, 585, 357-362.",
        "Jones, M. C. y Faddy, M. J. (2003). A skew extension of the t-distribution, with applications. Journal of the Royal Statistical Society B, 65(1), 159-174.",
        "Law, A. M. (2015). Simulation modeling and analysis (5.ª ed.). McGraw-Hill.",
        "McCullagh, P. (1980). Regression models for ordinal data. Journal of the Royal Statistical Society B, 42(2), 109-142.",
        "Seabold, S. y Perktold, J. (2010). Statsmodels: econometric and statistical modeling with Python. Proceedings of the 9th Python in Science Conference.",
        "Taylor, J. B. (1993). Discretion versus policy rules in practice. Carnegie-Rochester Conference Series on Public Policy, 39, 195-214.",
        "Virtanen, P. et al. (2020). SciPy 1.0: fundamental algorithms for scientific computing in Python. Nature Methods, 17, 261-272.",
        "Wilson, E. B. (1927). Probable inference, the law of succession, and statistical inference. Journal of the American Statistical Association, 22(158), 209-212.",
        "Equipo del trabajo. (2026). Estrategias de inversión; Escenarios; Dependencia entre variables y estados económicos; Diseño experimental; Diagrama de eventos y diagrama de flujo; Análisis de estados económicos (Markov). Documentos internos.",
    ]:
        w.p(ref)


def anexos(w: Informe):
    w.salto()
    w.h("Anexo A. Parámetros del simulador")
    w.p("Copia de la hoja PARAMETROS del Excel maestro. Adverso y Estrés vacíos significan el mismo valor de la base.")
    filas = []
    for f in FILAS:
        def v(x):
            if x is None:
                return ""
            if isinstance(x, float):
                return c(x, 4) if abs(x) < 10 else c(x, 2)
            if isinstance(x, int):
                return c(x)
            return str(x)
        filas.append([f[0], f[3], v(f[4]), v(f[5]), v(f[6]), f[7], f[8]])
    w.tabla("A", "Parámetros.", ["ID", "Parámetro", "Base", "Adverso", "Estrés", "Unidad", "Estado"], filas,
            [0.5, 2.6, 0.9, 0.7, 0.7, 0.9, 0.8], tam=7)
    w.h("Anexo B. Traza de eventos de una réplica")
    w.p("Primeros eventos de la réplica 0 de E4 en el escenario base (semilla 20261107), sin los cierres de mes ni los arriendos. "
        "La traza completa está en outputs/resultados/traza_E4_replica0.csv.")
    t = TRAZA[~TRAZA.evento.isin(["Cierre de mes", "Pago de arriendo"])].head(40)
    w.tabla("B", "Traza de eventos.", ["Mes", "N.º", "Evento", "Qué pasó", "Estado", "Cuenta (M)", "Patrimonio (M)"],
            [[c(x.t, 2), x.n, x.evento, str(x.detalle).replace(".", ","), x.estado, c(x.cuenta / 1e6, 1), c(x.W / 1e6, 1)]
             for x in t.itertuples()], [0.5, 0.4, 1.3, 2.5, 0.7, 0.7, 0.8], tam=7)


def main(origen: Path) -> None:
    doc = Document(str(origen))
    portada_y_resumen(doc)
    notas_7(doc)
    w = Informe(doc)
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
    discusion_y_cierre(w)
    anexos(w)
    DESTINO.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(DESTINO))
    print("ok", DESTINO)


if __name__ == "__main__":
    main(Path(sys.argv[1]) if len(sys.argv) > 1 else ORIGEN)

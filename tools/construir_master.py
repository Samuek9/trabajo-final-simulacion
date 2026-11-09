"""Hojas nuevas del Excel maestro (las que no vienen del consolidado ni del Markov del equipo).

  python tools/construir_master.py SALIDA.xlsx

Después tools/unir_master.ps1 junta este libro con las hojas que traen gráficos
(consolidado de los puntos 1 a 7 y análisis de Markov) usando Excel, porque
openpyxl borra los gráficos al guardar.
"""
from __future__ import annotations

import gzip
import json
import sys
from collections import Counter
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(Path(__file__).parent))
from hoja_parametros import escribir as escribir_parametros  # noqa: E402
from parametros_def import FILAS  # noqa: E402

RES = RAIZ / "outputs" / "resultados"
AZUL = "1F4E78"
ID = {f[1]: f[0] for f in FILAS}


def ids(*prefijos: str) -> str:
    return ", ".join(f[0] for f in FILAS if f[1].startswith(prefijos))


def titulo(ws, texto: str, nota: str | None = None) -> int:
    ws["A1"] = texto
    ws["A1"].font = Font(bold=True, size=14, color=AZUL)
    if nota:
        ws["A2"] = nota
        ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells("A2:H2")
        ws.row_dimensions[2].height = 45
    return 4


def tabla(ws, fila: int, encabezado: list[str], filas: list[list], anchos: list[int] | None = None,
          subtitulo: str | None = None) -> int:
    if subtitulo:
        ws.cell(fila, 1, subtitulo).font = Font(bold=True, color=AZUL, size=12)
        fila += 1
    for j, c in enumerate(encabezado, 1):
        cel = ws.cell(fila, j, c)
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor=AZUL)
        cel.alignment = Alignment(wrap_text=True, vertical="center")
    for i, f in enumerate(filas, fila + 1):
        for j, v in enumerate(f, 1):
            if isinstance(v, float) and v != v:
                v = None
            cel = ws.cell(i, j, v)
            cel.alignment = Alignment(wrap_text=isinstance(v, str) and len(v) > 30, vertical="top")
    if anchos:
        for j, a in enumerate(anchos, 1):
            ws.column_dimensions[get_column_letter(j)].width = a
    return fila + len(filas) + 2


def leeme(ws) -> None:
    corrida = json.loads((RES / "corrida.json").read_text(encoding="utf-8"))
    f = titulo(ws, "Excel maestro del trabajo final de Simulación",
               f"Reúne en un solo libro los datos, parámetros y resultados de los puntos 1 a 24. Generado el {date.today():%d/%m/%Y}. "
               "El simulador lee solo la hoja PARAMETROS; las demás documentan de dónde sale cada número. "
               "Las hojas INF, TPM_Cuenta, CDT, ETF, INM y PIB vienen del consolidado de los puntos 1 a 7; las MARKOV, "
               "del análisis de estados económicos del equipo, copiadas sin cambios.")
    hojas = [
        ["LEEME", "Esta hoja", ""],
        ["PARAMETROS", f"Los {len(FILAS)} parámetros del simulador con valor base, adverso y de estrés, unidad, estado y fuente", "Puntos 7 a 14"],
        ["FUENTES", "Fuentes de datos y documentos del equipo, con cuántos parámetros sale de cada una", "Todos"],
        ["DICCIONARIO", "Qué significa cada columna de los resultados por réplica", "Punto 15"],
        ["DISTRIBUCIONES", "Modelo de cada proceso aleatorio, parámetros que usa y cómo se validó", "Puntos 7 y 16"],
        ["DECISIONES", "Modelos alternativos, supuestos y pendientes que el equipo debe revisar", "Todos"],
        ["MARKOV", "Matriz de transición y resultado del análisis de estados (equipo)", "Punto 8"],
        ["MARKOV_PIB, MARKOV_ETF, MARKOV_Vivienda, MARKOV_Fuentes", "Datos del análisis de estados (equipo)", "Punto 8"],
        ["INF", "Inflación: datos DANE y ajuste del AR(1) en logaritmos", "7.1"],
        ["TPM_Cuenta", "TPM, tabla de decisiones por tendencia y tasa de la cuenta", "7.2 y 7.3"],
        ["TPM_Logit", "Las 202 decisiones de la Junta 2001-2026 (calendario de 8 reuniones) y el logit ordenado que se usa como referencia", "7.2"],
        ["CDT", "Margen de los CDT a 360 días sobre la TPM, modelo M5", "7.4"],
        ["ETF", "Rendimientos mensuales del iCOLCAP y ajuste de distribuciones", "7.5"],
        ["INM", "Valorización, canon, vacancia y gastos del inmueble", "7.6 a 7.8"],
        ["PIB", "Crecimiento del PIB y clasificación en estados", "Punto 8"],
        ["VALIDACION", "Simulado contra observado para cada proceso", "Punto 16"],
        ["RESULTADOS", "Indicadores de las 12 configuraciones, capital necesario, precisión y comparaciones pareadas", "Puntos 15 a 18"],
    ]
    f = tabla(ws, f, ["Hoja", "Contenido", "Punto del enunciado"], hojas, [34, 95, 20], "Hojas")
    archivos = []
    for p in sorted((RAIZ / "data" / "raw").glob("*.csv")):
        archivos.append([f"data/raw/{p.name}", sum(1 for _ in p.open(encoding="utf-8")) - 1, "Serie de entrada"])
    with gzip.open(RES / "replicas.csv.gz", "rt", encoding="utf-8") as g:
        archivos.append(["outputs/resultados/replicas.csv.gz", sum(1 for _ in g) - 1, "Una fila por réplica, estrategia y escenario"])
    for nombre, desc in (("bandas.csv", "Percentiles mes a mes del patrimonio real y la cuenta"),
                         ("capital_meta.csv", "Probabilidad de meta por capital, meta y estrategia"),
                         ("tornado.csv", "Sensibilidad uno a uno de 14 factores"),
                         ("traza_E4_replica0.csv", "Todos los eventos de una réplica de ejemplo")):
        archivos.append([f"outputs/resultados/{nombre}", len(pd.read_csv(RES / nombre)), desc])
    f = tabla(ws, f, ["Archivo del repositorio", "Filas", "Contenido"], archivos, None,
              "Series que no caben bien en Excel (están completas en el repositorio, sin recortar)")
    estados = [["Configurable", "Lo elige el usuario en la aplicación"], ["Verificado", "Dato tomado de una fuente oficial"],
               ["Estimado", "Sale de ajustar un modelo a los datos"], ["Calibrado", "Ajustado para reproducir una referencia"],
               ["Supuesto", "Decisión del equipo sin dato directo"], ["Pendiente", "Sin datos; el valor está apagado o es provisional"]]
    f = tabla(ws, f, ["Estado", "Significado"], estados, None, "Estados de un parámetro")
    ws.cell(f, 1, f"Corrida del informe: {corrida['replicas']} réplicas, semilla {corrida['semilla_base']}, "
                  f"modelo versión {corrida['version_modelo']}, {corrida['fecha']}.")


FUENTES = [
    ["F01", "IPC y variación anual", "DANE", "dane.gov.co", "2001-2025", "Inflación (7.1), canon"],
    ["F02", "PIB trimestral a precios constantes", "DANE", "dane.gov.co", "2006-2026", "Estados económicos (8)"],
    ["F03", "Tasa de política monetaria diaria", "Banco de la República", "banrep.gov.co", "2001-2026", "TPM (7.2)"],
    ["F04", "Decisiones de la Junta Directiva", "Banco de la República", "banrep.gov.co", "2001-2026", "Logit de la Junta (7.2)"],
    ["F05", "Tasa de captación CDT a 360 días, semanal", "Banco de la República / Superfinanciera", "banrep.gov.co", "2008-2026", "Margen CDT (7.4)"],
    ["F06", "Valor neto del iShares COLCAP (iCOLCAP)", "BlackRock", "blackrock.com", "2011-2026", "ETF (7.5)"],
    ["F07", "Índice de precios de vivienda nueva, Medellín (IPVNBR)", "Banco de la República", "banrep.gov.co", "2004-2026", "Valorización alternativa (7.6)"],
    ["F08", "Índice de precios de la vivienda usada (IPVU)", "Banco de la República", "banrep.gov.co", "1988-2025", "Valorización (7.6)"],
    ["F09", "Índice de arrendamientos (IPC de arriendos)", "DANE", "dane.gov.co", "2010-2025", "Reajuste del canon (5.4.3)"],
    ["F10", "Vacancia de arriendos en Medellín", "Ciencuadras", "ciencuadras.com", "2026", "Validación de la vacancia"],
    ["F11", "Ley 820 de 2003 (tope del reajuste del canon)", "Congreso de Colombia", "secretariasenado.gov.co", "", "Canon"],
    ["F12", "Estatuto Tributario: GMF y retención en la fuente", "DIAN", "dian.gov.co", "2026", "Impuestos"],
    ["F13", "Valor de la UVT 2026: Resolución 000238 del 15 de diciembre de 2025", "DIAN", "dian.gov.co", "2026", "Exención del 4x1000"],
    ["F14", "Estrategias de inversión.docx", "Equipo", "", "nov. 2026", "Punto 11"],
    ["F15", "Escenarios.docx", "Equipo", "", "nov. 2026", "Punto 13"],
    ["F16", "Dependencia entre variables y estados económicos.docx", "Equipo", "", "nov. 2026", "Punto 9"],
    ["F17", "Diseño experimental.docx", "Equipo", "", "nov. 2026", "Punto 14"],
    ["F18", "diagrama_eventos.pdf y diagrama_de_flujo_completo.pdf", "Equipo", "", "nov. 2026", "Puntos 9 y 10"],
    ["F19", "Analisis_completo_estados_economicos_Markov.xlsx", "Equipo", "", "nov. 2026", "Punto 8"],
    ["F20", "Trabajo final, puntos 1 a 7 (Word y Excel consolidados)", "Equipo", "", "oct. 2026", "Puntos 1 a 7"],
    ["F21", "Mercado inmobiliario residencial, edición 275 (precio, canon y administración por m²; días para arrendar y vender)",
     "Lonja de Propiedad Raíz de Medellín y Antioquia", "lonja.org.co", "ago. 2025", "Canon (5.4.1), venta, sensibilidad de vacancia"],
    ["F22", "Observatorio del Arriendo: tiempo en oferta de apartamentos en Medellín", "Palomma", "palomma.com", "2025-2026", "Vacancia (7.7)"],
    ["F23", "Encuesta de Carga Financiera y Educación Financiera de los Hogares (IEFIC), microdatos 2017 y 2018",
     "DANE (Fuente: Departamento Administrativo Nacional de Estadística: www.dane.gov.co)", "microdatos.dane.gov.co, catálogo 626",
     "2017-2018", "Referencia de frecuencia de las necesidades de liquidez (12); microdatos no publicados"],
    ["F24", "Tasas de cuentas digitales (Nu, Pibank, Lulo, RappiPay)", "El Tiempo, El Colombiano, Rankia", "", "dic. 2023 a oct. 2026", "Cuenta (7.3)"],
    ["F25", "Montos mínimos de CDT", "Noticias Caracol; El País", "noticiascaracol.com; elpais.com.co", "2026", "Inversión mínima (5.2)"],
]


def fuentes(ws) -> None:
    f = titulo(ws, "Fuentes", "Fuentes de datos y documentos del equipo. La columna Fuente de PARAMETROS cita estas mismas fuentes; "
                              "abajo está cuántos parámetros sale de cada texto de fuente.")
    f = tabla(ws, f, ["ID", "Fuente", "Entidad", "Sitio", "Periodo", "Se usa en"], FUENTES, [6, 55, 34, 24, 14, 40])
    c = Counter(fl[9] for fl in FILAS)
    tabla(ws, f, ["Fuente citada en PARAMETROS", "Parámetros"], [[k, v] for k, v in c.most_common()], None,
          "Parámetros por fuente")


DESCR = {
    "escenario": ("Escenario: base, adverso o estrés", "texto"), "replica": ("Número de réplica r", "entero"),
    "semilla": ("Semilla de la réplica (semilla base + r)", "entero"), "estrategia": ("E1 a E4", "texto"),
    "W_nominal": ("Patrimonio final en pesos corrientes, inmueble neto de costos de venta", "COP"),
    "W_real": ("Patrimonio final en pesos de hoy (dividido por el índice de precios simulado)", "COP de hoy"),
    "rent_nominal": ("Rentabilidad anual compuesta nominal", "fracción por año"),
    "rent_real": ("Rentabilidad anual compuesta real", "fracción por año"),
    "caida_max": ("Mayor caída del patrimonio real desde un máximo previo", "fracción"),
    "perdida_real": ("1 si el patrimonio real final es menor que el capital inicial", "0 o 1"),
    "meta": ("1 si el patrimonio real final alcanza la meta", "0 o 1"),
    "t_meta": ("Primer mes en que el patrimonio real alcanza la meta (vacío si nunca)", "meses"),
    "episodios_iliquidez": ("Veces que la cuenta no pudo cubrir una obligación sin vender de urgencia o quedar en deuda", "conteo"),
    "meses_iliquidez": ("Cierres de mes en iliquidez", "meses"),
    "duracion_max_iliquidez": ("Episodio de iliquidez más largo", "meses"),
    "duracion_media_iliquidez": ("Duración media de los episodios", "meses"),
    "algun_episodio": ("1 si hubo al menos un episodio de iliquidez", "0 o 1"),
    "liquido_final": ("Cuenta + CDT + ETF al final", "COP"), "liquido_final_real": ("Lo mismo en pesos de hoy", "COP de hoy"),
    "deuda_final": ("Obligaciones sin cubrir al final", "COP"), "insolvente": ("1 si termina con deuda", "0 o 1"),
    "inflacion_acumulada": ("Inflación acumulada en el horizonte", "fracción"),
    "arriendo_neto": ("Arriendo cobrado menos gastos del inmueble", "COP"),
    "gastos_inmueble": ("Administración, predial, mantenimiento, pintura y comisión", "COP"),
    "participacion_arriendo": ("Arriendo neto sobre la ganancia total", "fracción"),
    "rent_inmueble": ("Rentabilidad anual del inmueble (valorización + arriendo neto)", "fracción por año"),
    "tiempo_vacante": ("Fracción del tiempo con el inmueble desocupado", "fracción"),
    "perdida_venta_urgente": ("Descuento y costos perdidos por vender el inmueble de urgencia", "COP"),
    "costos_totales": ("Suma de comisiones, impuestos y castigos", "COP"),
    "estado_final": ("Estado económico al final", "texto"), "tpm_final": ("TPM al final", "%"),
    "inflacion_final": ("Inflación anual al final", "%"),
}
ACUM = {
    "arriendo_bruto": "arriendo cobrado", "comision_inmobiliaria": "comisión de la inmobiliaria", "ph": "administración",
    "predial": "predial", "mantenimiento": "mantenimiento", "pintura": "pintura", "reparaciones": "reparaciones",
    "costo_compra": "costos de compra del inmueble", "comision_etf": "comisión del ETF", "gmf": "4x1000",
    "retencion": "retención en la fuente", "intereses_cuenta": "intereses de la cuenta", "intereses_cdt": "intereses de los CDT",
    "intereses_perdidos": "intereses perdidos por venta anticipada de CDT", "perdida_realizada_etf": "pérdida realizada al vender ETF",
    "descuento_urgente": "descuento por venta urgente", "costo_venta": "costos de venta del inmueble",
    "necesidades": "necesidades extraordinarias de liquidez", "cubierto_cuenta": "necesidades cubiertas con la cuenta",
    "cubierto_etf": "cubierto vendiendo ETF", "cubierto_cdt": "cubierto vendiendo CDT", "cubierto_inmueble": "cubierto vendiendo el inmueble",
    "meses_vacante": "meses con el inmueble desocupado", "meses_con_inmueble": "meses con inmueble",
}
CONTEOS = {
    "necesidades": "necesidades extraordinarias", "coberturas": "veces que se activó la cobertura de liquidez",
    "ventas_etf": "ventas de ETF", "ventas_cdt": "ventas anticipadas de CDT", "ventas_inmueble": "ventas urgentes del inmueble",
    "vacancias": "periodos de vacancia", "renovaciones": "renovaciones del contrato", "rebalanceos": "rebalanceos de E4",
    "reparaciones": "reparaciones", "eventos": "eventos procesados",
}


def diccionario(ws) -> None:
    cols = pd.read_csv(RES / "replicas.csv.gz", nrows=1).columns
    filas = []
    for c in cols:
        if c in DESCR:
            d, u = DESCR[c]
        elif c.startswith("acum_"):
            k = c[5:]
            d, u = f"Total en el horizonte: {ACUM.get(k, k)}", "meses" if "meses" in k else "COP corrientes"
        elif c.startswith("n_"):
            d, u = f"Número de {CONTEOS.get(c[2:], c[2:])}", "conteo"
        else:
            d, u = "", ""
        filas.append([c, d, u])
    f = titulo(ws, "Diccionario de resultados", "Columnas de outputs/resultados/replicas.csv.gz (una fila por réplica y estrategia). "
                                               "En la hoja RESULTADOS cada indicador aparece resumido con su intervalo de confianza de 95 %.")
    f = tabla(ws, f, ["Columna", "Significado", "Unidad"], filas, [30, 90, 18])
    tray = [["W", "Patrimonio nominal al cierre de cada mes"], ["W_real", "Patrimonio real al cierre de cada mes"],
            ["cuenta, cdt, etf, inmueble", "Valor de cada activo al cierre"], ["deuda", "Obligaciones sin cubrir"],
            ["tpm, inflacion, indice", "TPM, inflación anual e índice de precios"], ["estado", "0 expansión, 1 normal, 2 crisis"],
            ["iliquido, ocupado", "Indicadores del mes"]]
    tabla(ws, f, ["Serie mensual", "Significado"], tray, None, "Trayectorias mensuales (trazabilidad y bandas)")


def distribuciones(ws) -> None:
    filas = [
        ["Estado económico", "Cadena de Markov de 3 estados, paso trimestral", "P(i → j) por conteo de transiciones", ids("markov.", "eco."), "Hoja MARKOV", "Proporción y duración de estados"],
        ["Inflación de diciembre", "ln π_t = c + φ ln π_{t−1} + σZ, Z normal", "c, φ, σ por mínimos cuadrados", ids("inf."), "Hoja INF", "Mediana, percentiles, autocorrelación"],
        ["Decisión de la Junta", "Logit ordenado (baja, mantiene, sube) con tendencia de inflación y brecha de Taylor; tamaño del paso por frecuencias", "Máxima verosimilitud, 202 decisiones fechadas por su vigencia", ids("junta.", "tpm."), "Hoja TPM_Logit", "TPM al año 10, tasa real"],
        ["Tasa de la cuenta", "α × TPM", "α de las cuentas remuneradas", ids("cta.", "imp."), "Hoja TPM_Cuenta", ""],
        ["Margen del CDT", "TPM + a + b·ΔTPM(3 meses) + ε, ε por bootstrap de los residuos", "Mínimos cuadrados, datos semanales", ids("cdt."), "Hoja CDT", "Margen simulado contra observado"],
        ["ETF iCOLCAP", "g = μ + vol(estado)·√((ν−2)/ν)·T, T ~ t de Student con ν grados", "μ y vol por estado; ν por máxima verosimilitud", ids("etf."), "Hoja ETF y MARKOV_ETF", "Media, desviación, curtosis, peor mes"],
        ["Valorización anual", "AR(1) del logaritmo (IPVU): r_t = μ + φ(r_{t−1} − μ) + σZ, igual en los tres estados", "Máxima verosimilitud, 2001-2025", ids("val."), "Hoja INM", "Media y desviación anual"],
        ["Reajuste del canon", "a + b·IPC previo + σZ, con tope en el IPC previo", "Mínimos cuadrados con el índice de arriendos", ids("inm."), "Hoja INM", ""],
        ["No renovación del contrato", "Bernoulli(p(estado)) al vencer cada contrato", "Supuesto del equipo", ids("vac.p"), "", "Fracción del tiempo desocupado"],
        ["Duración de la vacancia", "Gamma con forma k y media m(estado)", "Forma y media normal ajustadas a la curva de Palomma; medias de expansión y crisis supuestas", ids("vac.media", "vac.forma"), "", ""],
        ["Necesidades de liquidez", "Proceso de Poisson con tasa λ(estado); monto lognormal indexado", "Supuesto del equipo (IEFIC 2017-2018 como referencia de frecuencia)", ids("liq."), "", ""],
        ["Reparaciones", "Poisson y lognormal (apagado)", "Pendiente: sin datos", ids("rep."), "", ""],
        ["Crisis forzada (estrés)", "Inicio U(1, H − 12) y duración U(12, 24) meses", "Escenarios.docx", ids("estres."), "", ""],
    ]
    f = titulo(ws, "Distribuciones y modelos", "Cada proceso aleatorio tiene su propio generador de números, derivado de la semilla de la réplica.")
    tabla(ws, f, ["Proceso", "Modelo", "Estimación", "IDs en PARAMETROS", "Datos", "Validación (hoja VALIDACION)"], filas,
          [24, 60, 34, 40, 24, 34])


def decisiones(ws) -> None:
    filas = [
        ["D01", "Junta", "Referencia: logit ordenado. Alternativa: tabla por tendencia de 7.2", "Con la tabla la TPM llega al tope de 15 % en cerca de la mitad de las réplicas", "Cerrada"],
        ["D02", "CDT", "Referencia: M5 (Word). Alternativa: margen por fase", "M5 usa 18 años de datos semanales; los valores por fase del diagrama eran de una versión anterior", "Cerrada"],
        ["D03", "Canon", "Referencia: modelo DANE con tope. Alternativa: todo el IPC (diagrama)", "Los datos muestran reajustes por debajo del IPC", "Cerrada"],
        ["D04", "ETF", "Referencia: t por estado (equipo). Alternativa: Jones-Faddy", "Conecta el ETF con el estado económico", "Cerrada"],
        ["D05", "Valorización", "Referencia: AR(1) del IPVU. Alternativa: normal por estado con el IPVNBR de Medellín", "El inmueble es usado y el IPVNBR mide vivienda nueva; el estado no cambia la valorización (Welch, p = 0,275)", "Cerrada; con la alternativa la P(meta) sube 9 a 15 puntos"],
        ["D06", "Vacancia", "5/10/30 % de no renovar y 1/1,5/6 meses de duración media", "Simula 3 % del tiempo vacante contra 0,6 % de Ciencuadras", "Cerrada"],
        ["D07", "Estrés", "Recortes de TPM de 0,5 a 2 puntos durante la crisis forzada", "No hay dato; es un supuesto", "Cerrada"],
        ["D08", "Impuestos", "UVT 2026 = 52.374 para la exención del 4x1000", "Resolución DIAN 000238 del 15 de diciembre de 2025", "Cerrada"],
        ["D09", "Reparaciones", "Apagadas (frecuencia 0)", "Sin datos de costo ni frecuencia", "Pendiente"],
        ["D10", "Inflación", "Diciembre se sortea al inicio del año y la ruta mensual es lineal", "Evita que la Junta mire el futuro", "Cerrada"],
        ["D11", "Liquidez", "Lo que no se cubre queda como deuda y cuenta como iliquidez", "El enunciado no dice qué pasa si nada alcanza", "Cerrada"],
        ["D12", "Arriendo", "En el mes en que cambia el contrato se cobra el canon viejo", "Regla de desempate", "Cerrada"],
        ["D13", "Estados económicos", "Se usa la cadena de Markov de la hoja MARKOV (umbral de crisis 1 %, expansión 4 %) sin cambios",
         "Limitaciones: 81 transiciones, así que los intervalos son amplios (IC 90 % bootstrap: C→C 0,25 a 0,89; N→C 0,024 a 0,167; "
         "estacionaria de crisis 0,03 a 0,37). Parte de la persistencia es mecánica: el crecimiento interanual de un trimestre comparte "
         "tres trimestres con el anterior (autocorrelación 0,69). Las tres crisis observadas duraron 4 trimestres; la cadena genera "
         "crisis más largas (racha máxima mediana 6, p95 14). La matriz no cambia entre 2006-2015 y 2016-2026 (G = 4,29, p = 0,37). "
         "Impacto bajo: la probabilidad de crisis × 0,5 / × 2 mueve la P(meta) de E2 en 0,005", "Limitación declarada"],
    ]
    f = titulo(ws, "Decisiones", "Detalle completo en docs/decision_log.md del repositorio.")
    tabla(ws, f, ["ID", "Tema", "Decisión", "Por qué", "Estado"], filas, [6, 14, 60, 60, 30])


def tpm_logit(ws) -> None:
    j = json.loads((RAIZ / "data" / "metadata" / "junta_logit.json").read_text(encoding="utf-8"))
    f = titulo(ws, "Logit ordenado de la Junta (7.2)",
               "P(baja) = F(k1 − x), P(sube) = 1 − F(k2 − x), con x = b_tend·tendencia + b_taylor·brecha y F logística. "
               "Tendencia: cambio de la inflación en 3 meses. Brecha: TPM − (r* + π + 0,5(π − meta)).")
    coef = [["b_tend", j["b_tend"], j["p_tend"]], ["b_taylor", j["b_taylor"], j["p_taylor"]], ["k1", j["k1"], None], ["k2", j["k2"], None],
            ["AIC solo tendencia", j["aic_solo_tendencia"], None], ["AIC con Taylor", j["aic_con_taylor"], None],
            ["Razón de verosimilitud", j["lr"], j["lr_p"]], ["Decisiones", j["n"], None],
            ["Tasa real histórica media (puntos)", j["tasa_real_media"], None]]
    f = tabla(ws, f, ["Coeficiente", "Valor", "p"], coef, [34, 14, 14, 12, 12, 12, 12, 14], f"{j['desde']} a {j['hasta']}")
    pasos = [["Subida", float(k), v] for k, v in j["subidas"].items()] + [["Bajada", float(k), v] for k, v in j["bajadas"].items()]
    f = tabla(ws, f, ["Tipo", "Tamaño (puntos)", "Probabilidad"], pasos, None, "Tamaño del paso cuando la Junta mueve la tasa")
    d = pd.read_csv(RAIZ / "data" / "raw" / "junta_decisiones_2001_2026.csv")
    tabla(ws, f, list(d.columns), d.values.tolist(), None, "Decisiones usadas en la estimación")


def validacion(ws) -> None:
    v = pd.read_csv(RES / "validacion.csv")
    f = titulo(ws, "Validación (punto 16)", "Corrida de 500 réplicas con los parámetros de referencia frente a los datos de cada proceso.")
    tabla(ws, f, ["Proceso", "Medida", "Simulado", "Observado", "Referencia", "Lectura"],
          [[r.proceso, r.medida, _num(r.simulado), _num(r.observado), r.referencia, r.lectura] for r in v.itertuples()],
          [14, 48, 14, 14, 40, 70])


def _num(x):
    try:
        return round(float(x), 4)
    except (TypeError, ValueError):
        return x


def resultados(ws) -> None:
    r = pd.read_csv(RES / "resumen.csv")
    c = json.loads((RES / "corrida.json").read_text(encoding="utf-8"))
    f = titulo(ws, "Resultados", f"{c['replicas']} réplicas por estrategia y escenario, semilla {c['semilla_base']}, capital 700 M, "
                                 "meta real 1.000 M, 10 años. Montos en millones de pesos de hoy salvo que se diga otra cosa.")
    filas = []
    for x in r.itertuples():
        filas.append([x.escenario, x.estrategia, x.p_meta, x.p_meta_lo, x.p_meta_hi, x.W_real_media / 1e6, x.W_real_p5 / 1e6,
                      x.W_real_p25 / 1e6, x.W_real_p50 / 1e6, x.W_real_p75 / 1e6, x.W_real_p95 / 1e6, x.cvar5_W_real / 1e6, x.rent_real_media, x.p_perdida_real,
                      x.caida_max_p50, x.p_iliquidez, x.p_iliquidez_lo, x.p_iliquidez_hi, x.p_venta_urgente, x.t_meta_p50_anios])
    f = tabla(ws, f, ["Escenario", "Estrategia", "P(meta)", "IC bajo", "IC alto", "W real media", "W real P5", "W real P25", "W real P50",
                      "W real P75", "W real P95", "CVaR 5 %", "Rent. real anual", "P(pérdida real)", "Caída máx. P50", "P(iliquidez)",
                      "IC bajo", "IC alto", "P(venta urgente)", "Años a la meta P50"], filas,
              [11, 10] + [12] * 18, "Indicadores principales")
    n = pd.read_csv(RES / "capital_necesario.csv")
    f = tabla(ws, f, ["Escenario", "Estrategia", "Meta (M)", "Capital para 80 % (M)", "Capital para 50 % (M)"],
              [[x.escenario, x.estrategia, x.meta / 1e6, x.capital_80 / 1e6, x.capital_50 / 1e6] for x in n.itertuples()], None,
              "Capital necesario (punto 17; vacío = fuera del rango de 400 a 1.000 M)")
    pr = pd.read_csv(RES / "precision_replicas.csv")
    f = tabla(ws, f, list(pr.columns), pr.values.tolist(), None, "Precisión de P(meta) de E2 según el número de réplicas")
    pa = pd.read_csv(RES / "comparaciones_pareadas.csv")
    tabla(ws, f, list(pa.columns), pa.values.tolist(), None, "Diferencias pareadas (a − b) con números aleatorios comunes")


def main(salida: Path) -> None:
    wb = Workbook()
    hojas = [("LEEME", leeme), ("PARAMETROS", escribir_parametros), ("FUENTES", fuentes), ("DICCIONARIO", diccionario),
             ("DISTRIBUCIONES", distribuciones), ("DECISIONES", decisiones), ("TPM_Logit", tpm_logit),
             ("VALIDACION", validacion), ("RESULTADOS", resultados)]
    wb.remove(wb.active)
    for nombre, fn in hojas:
        fn(wb.create_sheet(nombre))
    wb.save(salida)
    print("ok", salida)


if __name__ == "__main__":
    main(Path(sys.argv[1]))

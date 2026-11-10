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
from hoja_parametros import escribir as escribir_parametros, escribir_modelo  # noqa: E402
from parametros_def import FILAS  # noqa: E402

RES = RAIZ / "outputs" / "resultados"
AZUL = "404040"  # gris oscuro: encabezados sobrios
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


HOJAS_ENLAZABLES = {"LEEME", "MODELOS", "PARAMETROS", "Inflación", "Inflación, modelo", "ETF, datos", "ETF, modelo", "TPM_Cuenta", "CDT, datos", "CDT, modelo", "INM", "PIB, estados", "MARKOV", "VALIDACION",
                    "RESULTADOS", "FUENTES", "DICCIONARIO", "AUDITORIA"}


def leeme(ws) -> None:
    corrida = json.loads((RES / "corrida.json").read_text(encoding="utf-8"))
    f = titulo(ws, "Simulación de estrategias de inversión bajo incertidumbre: datos, modelos y resultados",
               f"Libro generado el {date.today():%d/%m/%Y}. El simulador lee los valores de las hojas PARAMETROS (lo que "
               "cambia el usuario) y MODELO (oculta, con los valores estimados con datos, que quedan fijos). Las demás hojas "
               "documentan el estudio de cada proceso aleatorio y los resultados.")
    hojas = [
        ["LEEME", "Esta hoja: qué hay en el libro"],
        ["MODELOS", "El modelo elegido para cada proceso, con sus datos, sus pruebas y la razón"],
        ["PARAMETROS", "Capital, horizonte, meta, liquidez mínima, réplicas, estrategias y necesidades de liquidez: lo que se puede cambiar"],
        ["Inflación", "Evolución de la inflación frente a la meta y serie mensual"],
        ["Inflación, modelo", "Pruebas, estimación y validación del modelo de la inflación"],
        ["TPM_Cuenta", "TPM y tasa de la cuenta remunerada"],
        ["CDT, datos", "Evolución del CDT a 360 días y la TPM, y serie semanal"],
        ["CDT, modelo", "Pruebas, modelos candidatos, estimación, validación y efecto sobre el inversionista"],
        ["ETF, datos", "Evolución del NAV del ICOLCAP y serie mensual de rendimientos"],
        ["ETF, modelo", "Pruebas, modelos candidatos, estimación, validación y dependencia del ETF"],
        ["INM", "Canon (Lonja), valorización (IPVU), reajuste del canon y vacancia"],
        ["PIB, estados", "Crecimiento del PIB 1995-2026 y clasificación en estados"],
        ["MARKOV", "Matriz de transición, pruebas y qué cambia con el estado"],
        ["MARKOV_PIB, MARKOV_ETF, MARKOV_Vivienda", "Datos que usa la hoja MARKOV"],
        ["VALIDACION", "Simulado contra observado"],
        ["RESULTADOS", "Diseño del experimento, indicadores por estrategia y escenario, capital necesario, comparaciones y sensibilidad"],
        ["FUENTES", "Fuentes de datos"],
        ["DICCIONARIO", "Qué significa cada columna de los resultados por réplica"],
        ["INF, ETF, CDT y PIB (ocultas)", "Cálculo completo de las pruebas; las hojas de Inflación, ETF, CDT y PIB leen de ellas"],
        ["AUDITORIA", "Hallazgos y cambios de organización del libro, con su estado"],
    ]
    f = tabla(ws, f, ["Hoja", "Contenido"], hojas, [34, 110], "Hojas")
    for fila in range(1, ws.max_row + 1):
        v = ws.cell(fila, 1).value
        if v in HOJAS_ENLAZABLES:
            ws.cell(fila, 1).hyperlink = f"#'{v}'!A1"
            ws.cell(fila, 1).font = Font(color="1F4E78", underline="single")
    archivos = []
    for p in sorted((RAIZ / "data" / "raw").glob("*.csv")):
        archivos.append([f"data/raw/{p.name}", sum(1 for _ in p.open(encoding="utf-8")) - 1, "Serie de entrada"])
    with gzip.open(RES / "replicas.csv.gz", "rt", encoding="utf-8") as g:
        archivos.append(["outputs/resultados/replicas.csv.gz", sum(1 for _ in g) - 1, "Una fila por réplica, estrategia y escenario"])
    for nombre, desc in (("bandas.csv", "Percentiles mes a mes del patrimonio real y la cuenta"),
                         ("capital_meta.csv", "Probabilidad de meta por capital, meta y estrategia"),
                         ("tornado.csv", "Sensibilidad de un factor a la vez"),
                         ("traza_E4_replica0.csv", "Todos los eventos de una réplica de ejemplo")):
        archivos.append([f"outputs/resultados/{nombre}", len(pd.read_csv(RES / nombre)), desc])
    f = tabla(ws, f, ["Archivo del repositorio", "Filas", "Contenido"], archivos, None,
              "Series que no caben bien en Excel (están completas en el repositorio, sin recortar)")
    ws.cell(f, 1, f"Corrida de los resultados: {corrida['replicas']} réplicas, semilla {corrida['semilla_base']}, "
                  f"modelo versión {corrida['version_modelo']}, {corrida['fecha']}.")


FUENTES = [
    ["F01", "IPC y variación anual", "DANE", "dane.gov.co", "2001-2025", "Inflación (7.1), canon"],
    ["F02", "PIB trimestral a precios constantes", "DANE", "dane.gov.co", "2006-2026", "Estados económicos"],
    ["F03", "Tasa de política monetaria diaria", "Banco de la República", "banrep.gov.co", "2001-2026", "TPM (7.2)"],
    ["F04", "Decisiones de la Junta Directiva", "Banco de la República", "banrep.gov.co", "2001-2026", "Logit de la Junta (7.2)"],
    ["F05", "Tasa de captación CDT a 360 días, semanal", "Banco de la República / Superfinanciera", "banrep.gov.co", "2008-2026", "Margen CDT (7.4)"],
    ["F06", "Valor neto del iShares COLCAP (iCOLCAP)", "BlackRock", "blackrock.com", "2011-2026", "ETF (7.5)"],
    ["F07", "Índice de precios de vivienda nueva, Medellín (IPVNBR)", "Banco de la República", "banrep.gov.co", "2004-2026", "Valorización alternativa (7.6)"],
    ["F08", "Índice de precios de la vivienda usada (IPVU)", "Banco de la República", "banrep.gov.co", "1988-2025", "Valorización (7.6)"],
    ["F09", "Índice de arrendamientos (IPC de arriendos)", "DANE", "dane.gov.co", "2010-2025", "Reajuste del canon (5.4.3)"],
    ["F10", "Vacancia de arriendos en Medellín", "Ciencuadras", "ciencuadras.com", "2026", "Validación de la vacancia"],
    ["F11", "Ley 820 de 2003 (tope del reajuste del canon)", "Congreso de Colombia", "secretariasenado.gov.co", "", "Canon"],
    ["F12", "Estatuto Tributario (GMF) y Decretos 2418 de 2013, 1625 de 2016 y 572 de 2025 (retención en la fuente)", "DIAN / MinHacienda", "normograma.dian.gov.co", "2026", "Impuestos"],
    ["F13", "Valor de la UVT 2026: Resolución 000238 del 15 de diciembre de 2025", "DIAN", "dian.gov.co", "2026", "Exención del 4x1000"],
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
    f = titulo(ws, "Fuentes de datos")
    tabla(ws, f, ["ID", "Fuente", "Entidad", "Sitio", "Periodo", "Se usa en"], FUENTES, [6, 55, 34, 24, 14, 40])


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


def _sm(prefijo: str) -> str:
    """Cambio en P(meta) (puntos) al usar el modelo alternativo, E1 / E2 / E3 / E4."""
    sm = pd.read_csv(RES / "sensibilidad_modelos.csv")
    g = sm[sm.modelo.str.startswith(prefijo)].set_index("estrategia")
    return " / ".join(f"{(g.loc[e, 'alt_p_meta'] - g.loc[e, 'ref_p_meta']) * 100:+.1f}".replace(".", ",") for e in ("E1", "E2", "E3", "E4"))


def _tor(prefijo: str) -> str:
    """P(meta) de E2 con el valor bajo y el alto de un factor del tornado."""
    t = pd.read_csv(RES / "tornado.csv")
    x = t[t.factor.str.startswith(prefijo) & (t.estrategia == "E2")].iloc[0]
    return f"{x.bajo_p_meta * 100:.0f} % y {x.alto_p_meta * 100:.0f} %"


def modelos(ws) -> None:
    """Qué modelo elegimos para cada proceso, con qué datos, qué pruebas lo respaldan y por qué."""
    f = titulo(ws, "Modelos elegidos",
               "Por cada proceso aleatorio: los datos, las pruebas que respaldan la elección, el modelo que usa el simulador y "
               "la razón. Los valores estimados quedan fijos; no se modifican desde el libro ni desde la aplicación.")
    filas = [
        ["Inflación anual", "IPC del DANE, 2001 a 2025 (25 años)",
         "Normal y lognormal; años independientes frente a AR(1)",
         "Normalidad: depende de la prueba (Jarque-Bera p = 0,002; Shapiro-Wilk p = 0,057). En logaritmos no se rechaza "
         "(Shapiro-Wilk p = 0,81). Autocorrelación de ln π: 0,42 (p = 0,043). Media frente a la meta de 3 %: p = 0,020",
         "AR(1) en logaritmos: ln π(t) = c + φ·ln π(t−1) + σ·Z, con φ = 0,41 y nivel de largo plazo 4,5 %",
         "La inflación nunca es negativa y sus saltos son hacia arriba (lognormal), y un año depende del anterior. "
         "Pronosticada un año adelante cae en la banda de 90 % en 8 de 10 años de 2016 a 2025",
         f"Nivel 3,5 % o 6,5 %: E2 entre {_tor('Nivel de la inflación')}. Sin memoria o con φ = 0,6: E2 entre {_tor('Memoria')}",
         ids("inf."), "INF"],
        ["Decisión de la Junta (TPM)", "202 decisiones de 2001 a 2026 (52 subidas, 48 bajadas, 102 sin cambio)",
         "Tabla por tendencia de la inflación (7.2); logit ordenado solo con la tendencia; logit con tendencia y brecha de Taylor",
         "Agregar la brecha de Taylor mejora el ajuste (razón de verosimilitud 6,88; p = 0,009; AIC de 336,4 a 331,5). "
         "Brant: no se rechaza la proporcionalidad (p = 0,48). Las probabilidades predichas coinciden con las observadas por tercios",
         "Logit ordenado (baja, mantiene o sube) con la tendencia de la inflación y la brecha de Taylor; el tamaño del cambio "
         "se sortea con las frecuencias observadas. TPM entre 1,75 % y 15 %",
         "Con la tabla la TPM llega al tope de 15 % en cerca de la mitad de las réplicas, algo que nunca pasó (máximo 13,25 %). "
         "El logit reproduce las decisiones observadas",
         f"Tabla por tendencia: {_sm('Junta: tabla')}. Logit solo con 2017-2026: {_sm('Junta: logit')}",
         ids("junta.", "tpm."), "TPM_Logit, TPM_Cuenta"],
        ["Tasa de la cuenta", "19 tasas de 4 cuentas digitales en 5 fechas (dic-2023 a oct-2026)",
         "Proporción fija de la TPM (α)",
         "α no depende del nivel de la TPM (correlación −0,26; p = 0,29). IC de α por entidad: 0,755 a 1,046",
         "Tasa de la cuenta = 0,898 × TPM, retención de 7 %",
         "Las cuentas digitales siguen a la TPM con una proporción estable",
         f"α = 0,80 o 0,95: E2 entre {_tor('Tasa de la cuenta')}",
         ids("cta."), "TPM_Cuenta"],
        ["Margen del CDT sobre la TPM", "CDT a 360 días y TPM, semanal, 2008 a 2026",
         "Margen fijo, margen por fase de la TPM y modelos con el cambio reciente de la TPM (M2 a M6)",
         "El coeficiente del cambio de la TPM en 13 semanas es significativo (errores HAC). Los residuos no son normales "
         "(Jarque-Bera p < 0,001). Fuera de muestra (2020-2026) M5 tiene un error de 1,78 puntos y el margen fijo 1,69",
         "M5: margen = 1,06 + 0,67·ΔTPM(3 meses) + residuo sorteado de los observados; retención de 4 %",
         "Es el único que recoge que el margen se abre cuando la TPM sube y se cierra cuando baja, aunque fuera de muestra no "
         "predice mejor que el margen fijo. Los residuos se sortean de los datos porque no son normales",
         f"Margen por fase: {_sm('CDT')}. Constante ± 0,5 puntos: E2 entre {_tor('Margen del CDT')}",
         ids("cdt."), "CDT"],
        ["Rendimiento del ETF", "NAV del iShares COLCAP, 182 meses (2011 a 2026)",
         "Normal, logística, Laplace, t, t con la media muestral, t asimétrica de Jones-Faddy y otros (8 modelos)",
         "Normalidad rechazada (Jarque-Bera 361,7). Sin memoria (Ljung-Box p ≥ 0,16) ni efecto ARCH (p = 0,997). Por estado: "
         "la media no cambia (Kruskal-Wallis p = 0,39) y la volatilidad sí (Levene p = 0,004)",
         "t de Student con 5 grados de libertad, media 0,36 % mensual y volatilidad 4,7 % (expansión y normal) u 8,8 % (crisis)",
         "Tiene colas pesadas, no tiene memoria y permite que el estado de la economía mueva el riesgo. "
         "La Jones-Faddy ajusta mejor la asimetría pero no tiene estados",
         f"Jones-Faddy: {_sm('ETF')}. Media 0 o 0,86 %: E2 entre {_tor('Rendimiento medio del ETF')}",
         ids("etf."), "ETF, MARKOV_ETF"],
        ["Valorización del inmueble", "Índice de precios de la vivienda usada (IPVU), 2001 a 2025",
         "Normal independiente, AR(1) y normal por estado con el índice de vivienda nueva de Medellín (IPVNBR)",
         "Variación normal (Jarque-Bera p = 0,92). φ significativo (t = 4,5; p < 0,001) y residuos sin autocorrelación "
         "(p = 0,17). El estado no la cambia (Welch p = 0,275) ni la inflación del mismo año (p = 0,84)",
         "AR(1) del logaritmo: media 7,55 % anual, φ = 0,676, σ = 2,36 %; arranca en la variación de 2025 (8,09 %)",
         "El inmueble es usado y el IPVU mide vivienda usada; los años buenos y malos vienen seguidos. Fuera de muestra "
         "cae en la banda de 90 % los 10 años de 2016 a 2025",
         f"IPVNBR por estado: {_sm('Valorización')}. Media ± 2 puntos: E2 entre {_tor('Valorización')}",
         ids("val."), "INM, MARKOV_Vivienda"],
        ["Reajuste del canon", "IPC de arriendos del DANE, 2010 a 2025 (16 años)",
         "Sube todo el IPC (diagrama) frente a una regresión sobre el IPC del año anterior",
         "El coeficiente sobre el IPC previo es significativo (p < 0,001, errores robustos; R² = 0,71). Los arriendos suben en promedio menos que el IPC "
         "(3,93 % frente a 4,90 %)",
         "Reajuste = 1,68 % + 0,48·IPC previo + ε, sin pasar del IPC previo (Ley 820); canon inicial 0,60 % del valor (Lonja)",
         "Los datos muestran reajustes menores que el IPC; el tope legal se respeta",
         f"Todo el IPC: {_sm('Canon')}. Canon inicial 0,50 % o 0,70 %: E2 entre {_tor('Canon inicial')}",
         ids("inm.reajuste", "inm.canon"), "INM"],
        ["Vacancia", "Curva de tiempo en oferta de Palomma (25.084 apartamentos de Medellín, 2025-2026)",
         "Exponencial, gamma, Weibull y lognormal para la duración",
         "La gamma con forma 1,1 reproduce la curva; la probabilidad de no renovar no tiene datos (supuesto)",
         "No renovar con probabilidad 5 % / 10 % / 30 % según el estado; duración gamma con media 1 / 1,5 / 6 meses",
         "Es conservador a propósito: el apartamento pasa 2,5 % del tiempo vacío, frente a 0,6 % que reporta Ciencuadras",
         f"Duración de 4 meses (Lonja): {_sm('Vacancia')}. Probabilidad × 0,5 o × 2: E2 entre {_tor('Probabilidad de no renovar')}",
         ids("vac."), "INM, MARKOV"],
        ["Estado de la economía", "PIB trimestral del DANE, 2006 a 2026 (82 trimestres)",
         "Estados independientes frente a una cadena de Markov de 3 estados (crisis si el PIB crece menos de 1 %, "
         "expansión desde 4 %)",
         "El estado depende del trimestre anterior (χ² = 81,8; p < 0,001). La matriz no cambia entre 2006-2015 "
         "y 2016-2026 (G = 4,29; p = 0,37)",
         "Cadena de Markov trimestral; mueve la volatilidad del ETF, la vacancia, la frecuencia de las necesidades y el reparto "
         "de E4. No mueve la inflación, la TPM ni la valorización",
         "Las crisis duran varios trimestres seguidos. En el largo plazo la cadena pasa 32 / 52 / 16 % del tiempo en cada "
         "estado, frente a 37 / 49 / 15 % observado",
         f"Probabilidad de entrar en crisis × 0,5 o × 2: E2 entre {_tor('Probabilidad de crisis')}",
         ids("markov.", "eco."), "MARKOV, MARKOV_PIB, PIB"],
        ["Necesidades de liquidez", "Sin datos de monto. La IEFIC del DANE (2017-2018) solo sirve de referencia para la frecuencia",
         "Proceso de Poisson con monto lognormal",
         "Sin pruebas posibles por falta de datos de monto. La IEFIC sugiere un gasto mayor que el ingreso cada 16 a 23 meses",
         "Una cada 24 meses (cada 12 en crisis); monto lognormal con mediana de 3 % del capital inicial en pesos de hoy",
         "Representa imprevistos grandes (salud, desempleo, reparación); por eso los resultados se muestran con y sin ellos",
         f"Sin necesidades: {_sm('Sin necesidades')}. Frecuencia × 0,5 o × 1,5: E2 entre {_tor('Frecuencia de necesidades')}",
         ids("liq."), "PARAMETROS"],
    ]
    filas = [[r[0], r[1], r[3], r[4], r[5]] for r in filas]
    tabla(ws, f, ["Proceso", "Datos", "Pruebas y resultado", "Modelo que usa el simulador", "Por qué lo elegimos"], filas,
          [22, 32, 56, 48, 48])
    ws.freeze_panes = "B6"


def diseno_experimento(ws, f: int) -> int:
    c = json.loads((RES / "corrida.json").read_text(encoding="utf-8"))
    pr = pd.read_csv(RES / "precision_replicas.csv").set_index("replicas")
    pa = pd.read_csv(RES / "comparaciones_pareadas.csv")
    rn = c["replicas_necesarias_W_real"]
    filas = [
        ["Configuraciones", "4 estrategias × 3 escenarios (base, adverso, estrés) = 12"],
        ["Réplicas por configuración", f"{c['replicas']:,}".replace(",", ".") + ": n = 1,96² × 0,25 / 0,03², para que P(meta) tenga un error de ±3 puntos "
                                       "con 95 % de confianza en el peor caso"],
        ["Comprobación del número", f"Un piloto de 200 réplicas pide entre {min(rn.values())} y {max(rn.values())} para un error de "
                                    f"2 % en el patrimonio real medio. Con 2.400 réplicas P(meta) de E2 queda en "
                                    f"{pr.loc[2400, 'p_meta'] * 100:.1f} %, dentro del intervalo de la corrida de {c['replicas']}".replace(".", ",")],
        ["Semillas", f"Semilla base {c['semilla_base']}; la réplica r usa base + r, así que cualquier réplica se puede repetir sola"],
        ["Números aleatorios comunes", "Cada proceso tiene su propio generador y las cuatro estrategias enfrentan la misma economía en "
                                       "cada réplica; la varianza de las diferencias entre estrategias baja hasta 97 %"],
        ["Intervalos de confianza", "Wilson para proporciones, t para medias, bootstrap para percentiles y CVaR, t pareada para "
                                    f"comparar estrategias; {len(pa)} comparaciones con corrección de Holm"],
        ["Sensibilidad", f"{pd.read_csv(RES / 'tornado.csv').factor.nunique()} factores uno a la vez (400 réplicas), "
                         f"{pd.read_csv(RES / 'sensibilidad_modelos.csv').modelo.nunique()} modelos alternativos y una "
                         "interacción 2 × 2 de inflación por necesidades de liquidez"],
        ["Capital-meta", "7 capitales (400 a 1.000 M) × 5 metas (700 a 1.500 M) × 4 estrategias × 2 escenarios, 500 réplicas por celda"],
    ]
    return tabla(ws, f, ["Elemento", "Cómo se hizo"], filas, None, "Diseño del experimento (sección 14 del informe)")


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
    f = titulo(ws, "Validación", "Corrida de 500 réplicas con los parámetros de referencia frente a los datos de cada "
               "proceso. Tipo: 'Reproduce insumos' compara con los mismos datos con que se estimó el proceso (verifica el "
               "generador); 'Contraste' compara con algo que el modelo no usó al estimarse (validación).")
    tabla(ws, f, ["Proceso", "Medida", "Tipo", "Simulado", "Observado", "Referencia", "Lectura"],
          [[r.proceso, r.medida, r.tipo, _num(r.simulado), _num(r.observado), r.referencia, r.lectura] for r in v.itertuples()],
          [14, 48, 18, 14, 14, 40, 70])


def _num(x):
    try:
        return round(float(x), 4)
    except (TypeError, ValueError):
        return x


def resultados(ws) -> None:
    r = pd.read_csv(RES / "resumen.csv")
    c = json.loads((RES / "corrida.json").read_text(encoding="utf-8"))
    f = titulo(ws, "Resultados", f"{c['replicas']:,}".replace(",", ".") + f" réplicas por estrategia y escenario, semilla {c['semilla_base']}, capital 700 M, "
                                 "meta real 1.000 M, 10 años. Montos en millones de pesos de hoy salvo que se diga otra cosa.")
    f = diseno_experimento(ws, f)
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
              "Capital necesario (vacío = fuera del rango de 400 a 1.000 M)")
    pr = pd.read_csv(RES / "precision_replicas.csv")
    f = tabla(ws, f, list(pr.columns), pr.values.tolist(), None, "Precisión de P(meta) de E2 según el número de réplicas")
    pa = pd.read_csv(RES / "comparaciones_pareadas.csv")
    f = tabla(ws, f, list(pa.columns), pa.values.tolist(), None, "Diferencias pareadas (a − b) con números aleatorios comunes")
    t = pd.read_csv(RES / "tornado.csv")
    t = t[t.estrategia == "E2"].sort_values("rango_p_meta", ascending=False)
    f = tabla(ws, f, ["Factor", "Fuente de incertidumbre", "P(meta) referencia", "P(meta) valor bajo", "P(meta) valor alto", "Rango (puntos)"],
              [[x.factor, x.fuente, x.ref_p_meta, x.bajo_p_meta, x.alto_p_meta, x.rango_p_meta * 100] for x in t.itertuples()], None,
              "Sensibilidad de un factor a la vez, E2 (tornado, 400 réplicas por punto)")
    sm = pd.read_csv(RES / "sensibilidad_modelos.csv")
    f = tabla(ws, f, ["Modelo alternativo", "Estrategia", "P(meta) referencia", "P(meta) alternativa"],
              [[x.modelo, x.estrategia, x.ref_p_meta, x.alt_p_meta] for x in sm.itertuples()], None,
              "Sensibilidad a la elección de modelos (400 réplicas)")
    it = pd.read_csv(RES / "interaccion.csv")
    tabla(ws, f, ["Estrategia", "Inflación 3,5 %, necesidades × 0,5", "3,5 %, × 1,5", "6,5 %, × 0,5", "6,5 %, × 1,5", "Interacción (puntos)"],
          [[x.estrategia, x.inf_baja_liq_baja, x.inf_baja_liq_alta, x.inf_alta_liq_baja, x.inf_alta_liq_alta, x.interaccion * 100]
           for x in it.itertuples()], None, "Interacción de inflación y necesidades de liquidez: P(meta)")


AUDITORIA = [
    ["A01", "INF", "Serie mensual y anual", "El 6,29 % de 2026 es la inflación de septiembre, no de diciembre; el modelo lo usa solo como inflación inicial (P033).", "Aclaración temporal", "Media", "Se conserva el dato; la hoja Inflación lo rotula como septiembre de 2026", "Resuelto"],
    ["A02", "PARAMETROS", "Hoja completa", "Mezclaba valores de inversionista con valores estimados de los procesos aleatorios.", "Estructural", "Alta", "PARAMETROS solo trae lo editable; los valores estimados pasaron a la hoja oculta MODELO con las mismas claves", "Resuelto"],
    ["A03", "PARAMETROS", "Columnas Fuente y Soporte", "Citaban secciones del informe y no aportaban a quien edita.", "Visual", "Baja", "Se quitaron de la hoja visible; siguen en MODELO", "Resuelto"],
    ["A04", "MODELOS", "Columnas de alternativa y parámetros", "Mostraban más de un modelo por proceso.", "Estructural", "Media", "Queda solo el modelo elegido y la prueba que lo respalda", "Resuelto"],
    ["A05", "PARAMETROS", "P153 (necesidades de liquidez)", "El monto mediano era 20 millones fijos, sin relación con el capital.", "Metodológico", "Alta", "Pasa a ser 3 % del capital inicial (liq.mediana_pct); se corrieron de nuevo los experimentos", "Resuelto, con autorización del usuario"],
    ["A06", "PARAMETROS", "Reparaciones (P160 a P162)", "Sin datos de costo ni frecuencia.", "Dato faltante", "Media", "Salen de la vista y del informe; quedan apagadas en MODELO", "Resuelto"],
    ["A07", "RESULTADOS", "Capital-meta", "La malla de capitales y metas estaba fija en el código.", "Compatibilidad", "Media", "Se calcula como múltiplos del capital y la meta configurados; con la referencia da la misma malla", "Resuelto"],
    ["A08", "Todas", "Hojas del equipo", "Algunas notas mencionaban el enunciado o al equipo.", "Visual", "Baja", "Se reescribieron en la copia del libro; los originales no se tocan", "Resuelto"],
    ["A09", "Todas", "Datos originales y procesados", "Conviven en la misma hoja con pruebas y texto, en bloques horizontales.", "Estructural", "Media", "Inflación: las hojas Inflación e Inflación, modelo leen de INF con fórmulas, sin mover nada. Faltan TPM, inmueble y Markov", "Parcial"],
    ["A10", "PIB", "Nota de crisis y conteo desde 2011", "Decía que crisis es crecimiento negativo y contaba 4 trimestres de crisis desde 2011, pero los umbrales vigentes son 1 % y 4 % (12 trimestres de crisis desde 2006).", "Inconsistencia", "Media", "La hoja PIB, estados recalcula el conteo con fórmula y describe la regla con los umbrales vigentes; la hoja original queda oculta sin cambios", "Resuelto"],
    ["A10", "VALIDACION", "Autocorrelación de la inflación", "La simulada (0,19) queda por debajo de la observada (0,42) porque se mide con 10 años por réplica.", "Posible limitación estadística", "Baja", "Se documenta en el informe; no se cambia el modelo", "Documentado"],
]


def auditoria(ws) -> None:
    f = titulo(ws, "Auditoría y cambios del libro",
               "Registro de los hallazgos de organización y de los cambios hechos. Las correcciones metodológicas se hicieron solo con autorización.")
    tabla(ws, f, ["ID", "Hoja", "Ubicación", "Hallazgo", "Tipo", "Gravedad", "Acción", "Estado"], AUDITORIA,
          [6, 14, 28, 60, 22, 10, 60, 24])
    ws.freeze_panes = "A6"


def main(salida: Path) -> None:
    wb = Workbook()
    hojas = [("LEEME", leeme), ("MODELOS", modelos), ("PARAMETROS", escribir_parametros), ("MODELO", escribir_modelo), ("FUENTES", fuentes),
             ("DICCIONARIO", diccionario), ("AUDITORIA", auditoria), ("TPM_Logit", tpm_logit),
             ("VALIDACION", validacion), ("RESULTADOS", resultados)]
    wb.remove(wb.active)
    for nombre, fn in hojas:
        fn(wb.create_sheet(nombre))
    wb.save(salida)
    print("ok", salida)


if __name__ == "__main__":
    main(Path(sys.argv[1]))

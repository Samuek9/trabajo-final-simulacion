from pathlib import Path

import pandas as pd
import streamlit as st

from components import estado, ui
from investment_sim import __version__
from investment_sim.engine.events import Tipo
from investment_sim.models import inflation, markov

RAIZ = Path(__file__).resolve().parents[2]
p = estado.referencia()
T = p.tabla

st.title("Metodología")
st.write(f"Modelo de simulación de eventos discretos, versión {__version__}. El reloj salta de evento en evento; los eventos "
         "esperan en una cola ordenada por tiempo y, si coinciden, por la prioridad del diagrama de flujo. Cada proceso "
         "aleatorio tiene su propio generador, derivado de la semilla de la réplica, para que cambiar un proceso no altere "
         "los números de los demás.")

st.markdown("### Eventos")
filas = [{"N.º": t.numero, "Evento": t.nombre, "Prioridad en empates": t.prioridad} for t in Tipo if t.numero > 0]
st.dataframe(pd.DataFrame(filas).sort_values("Prioridad en empates"), hide_index=True, width="stretch")
st.caption("La cobertura de liquidez (evento 12) no entra a la cola: la llama el evento que genera el faltante, en el mismo "
           "instante. Orden de cobertura: cuenta, ETF, venta anticipada del CDT y, como último recurso, venta urgente del inmueble.")

st.markdown("### Procesos aleatorios")
dist = pd.DataFrame([
    ["Estado económico", "Cadena de Markov trimestral de 3 estados", "PIB real DANE 2006-2026", "Cada trimestre"],
    ["Inflación", "ln π = a + φ ln π anterior + σZ, interpolada mes a mes", "IPC DANE 2001-2025", "Cada diciembre"],
    ["Decisión de la Junta", "Logit ordenado con tendencia de la inflación y brecha de Taylor", "266 decisiones BanRep 2001-2026", "8 reuniones al año"],
    ["Margen del CDT", "TPM + a + b·ΔTPM + residuo observado (bootstrap)", "Tasas semanales de CDT 2008-2026", "Cada vencimiento"],
    ["ETF iCOLCAP", "t de Student con ν = 5 y volatilidad del estado", "Precios mensuales iCOLCAP", "Cada mes"],
    ["Valorización", "Normal del logaritmo según el estado", "IPVNBR Medellín", "Cada año"],
    ["Reajuste del canon", "a + b·IPC previo + ε, tope IPC previo", "DANE (índice de arrendamientos)", "Cada renovación"],
    ["Vacancia", "Bernoulli al vencer el contrato; duración gamma", "Supuesto del equipo", "Cada vencimiento"],
    ["Necesidades de liquidez", "Poisson según el estado; monto lognormal", "Supuesto del equipo", "Tiempo continuo"],
], columns=["Proceso", "Modelo", "Datos", "Frecuencia"])
st.dataframe(dist, hide_index=True, width="stretch")

c1, c2 = st.columns(2)
with c1:
    st.markdown("**Matriz de transición (referencia)**")
    P = markov.matriz(p)
    st.dataframe(pd.DataFrame(P, index=markov.ESTADOS, columns=markov.ESTADOS).round(4), width="stretch")
    pi = markov.estacionaria(P)
    st.caption("Largo plazo: " + " · ".join(f"{s} {x * 100:.0f} %" for s, x in zip(markov.ESTADOS, pi)))
with c2:
    st.markdown("**Inflación de largo plazo implícita**")
    st.write(f"{ui.num(inflation.nivel_largo_plazo(p.f('inf.c'), p.f('inf.phi')), 2)} % con φ = {ui.num(p.f('inf.phi'), 3)} "
             f"y σ = {ui.num(p.f('inf.sigma'), 3)}.")
    st.markdown("**Números aleatorios comunes**")
    st.write("La réplica r usa la misma semilla para las cuatro estrategias. En la corrida del informe esto bajó la varianza "
             "de la diferencia de patrimonio entre E4 y E2 en cerca de 96 %.")

st.markdown("### Supuestos y pendientes")
sup = [t for t in T.values() if t.estado in ("Supuesto", "Pendiente")]
st.dataframe(pd.DataFrame([{"ID": t.id, "Parámetro": t.nombre, "Valor": t.base, "Unidad": t.unidad, "Estado": t.estado,
                            "Fuente": t.fuente} for t in sup]), hide_index=True, width="stretch")

st.markdown("### Todos los parámetros")
grupo = st.selectbox("Grupo", ["Todos"] + sorted({t.grupo for t in T.values()}))
todos = [t for t in T.values() if grupo == "Todos" or t.grupo == grupo]
st.dataframe(pd.DataFrame([{"ID": t.id, "Clave": t.clave, "Parámetro": t.nombre, "Base": t.base, "Adverso": t.adverso,
                            "Estrés": t.estres, "Unidad": t.unidad, "Estado": t.estado, "Fuente": t.fuente}
                           for t in todos]).astype(str).replace("None", ""), hide_index=True, width="stretch")
st.caption("Vacío en Adverso o Estrés = mismo valor que en Base. La hoja PARAMETROS de data/master_datos.xlsx es la fuente; "
           "el motor no tiene números escritos en el código.")

st.markdown("### Diagramas del equipo")
for nombre, archivo in (("Diagrama de eventos", "diagrama_eventos.pdf"), ("Diagrama de flujo completo", "diagrama_de_flujo_completo.pdf")):
    f = RAIZ / "docs" / "diagramas" / archivo
    if f.exists():
        st.download_button(f"{nombre} (PDF)", f.read_bytes(), archivo, "application/pdf")
img = sorted((RAIZ / "docs" / "diagramas").glob("*.png")) if (RAIZ / "docs" / "diagramas").exists() else []
for f in img:
    st.image(str(f), caption=f.stem.replace("_", " "), width="stretch")

import streamlit as st

from components import estado, ui
from investment_sim.config.parameters import advertencias

p = estado.params()

st.title("Investment Lab")
ui.subtitulo("Simulación de eventos discretos de estrategias de inversión bajo incertidumbre: cuenta remunerada, CDT, ETF e inmueble en arriendo.")

st.markdown("##### Configuración actual")
c = st.columns(4)
with c[0]:
    ui.tarjeta("Capital inicial", ui.cop(p.f("inv.capital")))
with c[1]:
    ui.tarjeta("Horizonte", f"{p.i('inv.horizonte') // 12} años")
with c[2]:
    ui.tarjeta("Meta real", ui.cop(p.f("inv.meta_real")), "pesos de hoy")
with c[3]:
    ui.tarjeta("Liquidez mínima", ui.pct(p.f("inv.liquidez_min"), 0), "del patrimonio, en la cuenta")

st.markdown("##### Estrategias")
desc = {"E1": "Conserva el capital y prioriza la liquidez.", "E2": "Reparte entre los cuatro activos.",
        "E3": "Busca crecimiento con ETF e inmueble.", "E4": "Parte de E2 y rebalancea según el estado de la economía."}
cols = st.columns(4)
for col, e in zip(cols, ("E1", "E2", "E3", "E4")):
    with col:
        st.markdown(f"<div class='franja' style='border-color:{ui.COLORES[e]}'><b>{ui.NOMBRES[e]}</b></div>", unsafe_allow_html=True)
        ui.barra_composicion(p.lista(f"{e}.pesos"), leyenda=False)
        st.caption(desc[e])
st.markdown("<div class='leyenda'>" + "".join(f"<span><i style='background:{c}'></i>{a}</span>"
            for a, c in zip(ui.ACTIVOS, ui.COL_ACTIVOS)) + "</div>", unsafe_allow_html=True)

st.write("")
b = st.columns(2)
with b[0]:
    st.page_link("pages/02_inversionista.py", label="Configurar al inversionista")
with b[1]:
    st.page_link("pages/05_experimento.py", label="Ir al experimento")
for a in advertencias(p):
    st.info(a)
st.caption("Herramienta académica: los resultados describen el modelo y sus supuestos, no son una recomendación de inversión.")

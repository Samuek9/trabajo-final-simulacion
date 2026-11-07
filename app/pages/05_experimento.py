import os

import streamlit as st

from components import estado, ui
from investment_sim.experiments.runner import correr
from investment_sim.metrics.summary import resumen

p = estado.params()

st.title("Experimento")
st.write("Cada réplica r usa la semilla base + r para las cuatro estrategias, así todas enfrentan la misma economía "
         "(números aleatorios comunes) y las diferencias entre ellas se deben a la estrategia, no al azar.")

c1, c2 = st.columns(2)
with c1:
    escenarios = st.multiselect("Escenarios", list(ui.ESCENARIOS), default=["base"], format_func=ui.ESCENARIOS.get)
    estrategias = st.multiselect("Estrategias", ["E1", "E2", "E3", "E4"], default=["E1", "E2", "E3", "E4"],
                                 format_func=ui.NOMBRES.get)
with c2:
    n = st.number_input("Réplicas por estrategia", 20, 20_000, p.i("exp.replicas"), 50)
    semilla = st.number_input("Semilla base", 0, 2**31 - 1, p.i("exp.semilla"))
    nucleos = os.cpu_count() or 1
    procesos = st.slider("Procesos en paralelo", 1, max(nucleos, 1), min(4, nucleos),
                         help="El resultado no cambia con el número de procesos; solo el tiempo.")

st.caption("Escenario adverso: crisis más frecuentes y largas, inflación alrededor de 6,5 %, cuenta al 80 % de la TPM, ETF sin "
           "rendimiento medio, valorización 2 puntos menor, más vacancia y más necesidades de liquidez. "
           "Estrés: una crisis forzada de 12 a 24 meses en un momento al azar, con caída del ETF, del inmueble y recortes de la TPM.")

listo = bool(escenarios and estrategias)
if st.button("Ejecutar simulación", type="primary", disabled=not listo):
    st.session_state["resultados"] = {}
    barra = st.progress(0.0, "Preparando...")
    for i, esc in enumerate(escenarios):
        pe = estado.params(esc)
        def avance(x, i=i, esc=esc):
            barra.progress((i + x) / len(escenarios), f"{ui.ESCENARIOS[esc]}: {x * 100:.0f} %")
        r = correr(pe, estrategias, n=int(n), semilla=int(semilla), procesos=int(procesos), progreso=avance)
        st.session_state["resultados"][esc] = {"exp": r, "resumen": resumen(r.tabla), "params": pe}
    barra.progress(1.0, "Listo")
    st.success("Simulación terminada. Los resultados están en la página Resultados.")

res = st.session_state.get("resultados", {})
if res:
    st.markdown("##### Última ejecución")
    for esc, d in res.items():
        r = d["exp"]
        st.write(f"**{ui.ESCENARIOS[esc]}**: {r.n:,} réplicas × {len(r.estrategias)} estrategias, semilla {r.semilla}, "
                 f"{r.segundos:.1f} s, {len(r.cambios)} parámetro(s) cambiado(s).".replace(",", "."))
    st.caption("Abre Resultados en el menú de la izquierda.")
else:
    st.info("Todavía no hay resultados en esta sesión. La página Resultados también puede mostrar la corrida del informe.")

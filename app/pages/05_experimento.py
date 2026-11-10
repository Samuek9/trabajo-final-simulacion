import os

import streamlit as st

from components import estado, ui
from investment_sim.experiments.runner import correr
from investment_sim.metrics.summary import resumen

p = estado.params()
T = p.tabla

st.title("Experimento")
ui.subtitulo("Las estrategias se evalúan con números aleatorios comunes: la réplica r usa la semilla base + r en todas.")

DESC = {"base": "Condiciones estimadas con los datos.",
        "adverso": "Crisis más frecuentes, inflación y vacancia altas, ETF sin rendimiento medio.",
        "estres": "Crisis forzada de 12 a 24 meses con caída del ETF y del inmueble."}

st.markdown("##### Escenarios y estrategias")
c1, c2 = st.columns(2)
with c1:
    escenarios = st.multiselect("Escenarios", list(ui.ESCENARIOS), default=["base"], format_func=ui.ESCENARIOS.get)
    for e in escenarios:
        st.caption(f"**{ui.ESCENARIOS[e]}.** {DESC[e]}")
with c2:
    estrategias = st.multiselect("Estrategias", ["E1", "E2", "E3", "E4"], default=["E1", "E2", "E3", "E4"], format_func=ui.NOMBRES.get)

st.markdown("##### Réplicas y semilla")
c3, c4, c5 = st.columns(3)
with c3:
    n = st.number_input("Réplicas por estrategia", 20, 20_000, p.i("exp.replicas"), 50, help=ui.pie_parametro(T["exp.replicas"]))
with c4:
    semilla = st.number_input("Semilla base", 0, 2**31 - 1, p.i("exp.semilla"))
with c5:
    nucleos = os.cpu_count() or 1
    procesos = st.slider("Procesos en paralelo", 1, max(nucleos, 1), min(4, nucleos), help="No cambia el resultado, solo el tiempo.")
for k, v in (("exp.replicas", int(n)), ("exp.semilla", int(semilla))):
    if v != p.valores[k]:
        estado.fijar(k, v)

with st.expander("Réplicas sugeridas"):
    st.caption("Un piloto de 100 réplicas por estrategia estima qué tan cerca de 50 % queda la probabilidad de meta con la configuración actual.")
    margen = st.number_input("Margen de error (± puntos porcentuales)", 0.5, 10.0, 3.0, 0.5)
    if st.button("Calcular"):
        with st.spinner("Corriendo el piloto..."):
            rp = correr(estado.params(), ["E1", "E2", "E3", "E4"], n=100, semilla=int(semilla), procesos=1)
        pm = resumen(rp.tabla).set_index("estrategia").p_meta
        peor = float(min(pm, key=lambda v: abs(v - 0.5)))
        st.session_state["n_sugerido"] = (int(1.96 ** 2 * peor * (1 - peor) / (margen / 100) ** 2) + 1, peor, margen)
    if "n_sugerido" in st.session_state:
        sug, peor, mg = st.session_state["n_sugerido"]
        st.info(f"{sug:,} réplicas para ±{mg:g} puntos (probabilidad de meta más cercana a 50 % en el piloto: {peor * 100:.0f} %).".replace(",", "."))
    err = 1.96 * (0.25 / max(int(n), 1)) ** 0.5
    st.caption(f"Con {int(n):,} réplicas el error de la probabilidad de meta es de ±{err * 100:.1f} puntos como máximo.".replace(",", "."))

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
    st.success("Simulación terminada.")
    st.page_link("pages/06_resultados.py", label="Ver resultados")

res = st.session_state.get("resultados", {})
if res:
    st.markdown("##### Última ejecución")
    for esc, d in res.items():
        r = d["exp"]
        st.write(f"**{ui.ESCENARIOS[esc]}** · {r.n:,} réplicas × {len(r.estrategias)} estrategias · semilla {r.semilla} · {r.segundos:.1f} s".replace(",", "."))

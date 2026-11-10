import os

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from components import charts, estado, ui
from investment_sim.experiments import capital_goal, sensitivity

st.title("Sensibilidad y capital-meta")
NUC = os.cpu_count() or 1
res_sesion = st.session_state.get("resultados", {})
if not res_sesion:
    st.info("Estos análisis usan la configuración y las estrategias de la última ejecución. Ejecute una simulación en Experimento.")
    st.page_link("pages/05_experimento.py", label="Ir al experimento")
    st.stop()

escs = list(res_sesion)
EST = list(next(iter(res_sesion.values()))["exp"].estrategias)
n_base = next(iter(res_sesion.values()))["exp"].n
p = estado.params()

t1, t2, t3 = st.tabs(["Capital-meta", "Sensibilidad de parámetros", "Escenarios"])

with t1:
    ui.subtitulo("Probabilidad de alcanzar la meta según el capital inicial, la meta y la estrategia.")
    CAP0, META0 = capital_goal.malla(p)
    sufijo = f"{p.f('inv.capital'):.0f}_{p.f('inv.meta_real'):.0f}"
    st.markdown("##### Rangos del análisis")
    g = st.columns(3)
    with g[0]:
        cap_lo = st.number_input("Capital desde (millones)", 1.0, 100_000.0, float(min(CAP0) / 1e6), 10.0, key=f"cap_lo_{sufijo}")
    with g[1]:
        cap_hi = st.number_input("Capital hasta (millones)", 1.0, 100_000.0, float(max(CAP0) / 1e6), 10.0, key=f"cap_hi_{sufijo}")
    with g[2]:
        n_cap = st.number_input("Número de capitales", 2, 12, len(CAP0), 1, key=f"n_cap_{sufijo}")
    g = st.columns(3)
    with g[0]:
        meta_lo = st.number_input("Meta desde (millones)", 1.0, 100_000.0, float(min(META0) / 1e6), 10.0, key=f"meta_lo_{sufijo}")
    with g[1]:
        meta_hi = st.number_input("Meta hasta (millones)", 1.0, 100_000.0, float(max(META0) / 1e6), 10.0, key=f"meta_hi_{sufijo}")
    with g[2]:
        n_meta = st.number_input("Número de metas", 2, 12, len(META0), 1, key=f"n_meta_{sufijo}")
    if cap_hi <= cap_lo or meta_hi <= meta_lo:
        st.error("El valor «hasta» debe ser mayor que «desde».")
        st.stop()
    CAPS = [round(x) * 1e6 for x in np.linspace(cap_lo, cap_hi, int(n_cap))]
    METAS = [round(x) * 1e6 for x in np.linspace(meta_lo, meta_hi, int(n_meta))]
    n_cm = st.number_input("Réplicas por celda", 50, 20_000, max(50, min(int(n_base), 500)), 50, key="n_cm",
                           help="Cada capital se simula con estas réplicas; las metas se evalúan sobre el mismo resultado.")
    st.caption(f"{len(CAPS)} capitales × {len(escs)} escenario(s) × {len(EST)} estrategia(s) × {int(n_cm):,} réplicas.".replace(",", "."))
    if st.button("Calcular capital-meta", type="primary"):
        barra = st.progress(0.0)
        partes = []
        for i, esc in enumerate(escs):
            partes.append(capital_goal.matriz(res_sesion[esc]["params"], CAPS, METAS, EST, n=int(n_cm), semilla=p.i("exp.semilla"),
                                              procesos=min(4, NUC), progreso=lambda x, i=i: barra.progress((i + x) / len(escs))))
        st.session_state["cm"] = pd.concat(partes, ignore_index=True)
    cm = st.session_state.get("cm")
    if cm is not None and len(cm):
        esc_cm = st.radio("Escenario", list(dict.fromkeys(cm.escenario)), format_func=ui.ESCENARIOS.get, horizontal=True, key="esc_cm") \
            if cm.escenario.nunique() > 1 else cm.escenario.iloc[0]
        cm = cm[cm.escenario == esc_cm]
        e = st.radio("Estrategia", EST, format_func=ui.NOMBRES.get, horizontal=True, key="e_cm")
        st.plotly_chart(charts.mapa_capital_meta(cm, e), width="stretch", config=charts.CONFIG)
        objetivo = st.slider("Probabilidad objetivo", 0.1, 0.95, 0.8, 0.05)
        filas = []
        for m in sorted(cm.meta.unique()):
            fila = {"Meta real": ui.cop(m)}
            for x in EST:
                v = capital_goal.capital_necesario(cm, m, x, objetivo)
                fila[ui.NOMBRES[x]] = ui.cop(v) if v == v else "fuera del rango"
            filas.append(fila)
        st.markdown(f"**Capital necesario para {objetivo * 100:.0f} % de probabilidad**")
        st.dataframe(pd.DataFrame(filas), hide_index=True, width="stretch")
        st.caption("«Fuera del rango»: ningún capital del rango alcanza la probabilidad objetivo, o el más bajo ya la supera.")

with t2:
    ui.subtitulo("Efecto de llevar cada factor a un valor bajo y a uno alto, con los demás en la configuración actual.")
    c = st.columns(3)
    with c[0]:
        n_t = st.number_input("Réplicas por punto", 50, 20_000, max(50, min(int(n_base), 400)), 50, key="n_t")
    with c[1]:
        medida = st.radio("Medir", ["p_meta", "W"], format_func={"p_meta": "Probabilidad de meta", "W": "Patrimonio real mediano"}.get,
                          horizontal=True)
    with c[2]:
        e_t = st.selectbox("Estrategia", EST, format_func=ui.NOMBRES.get, key="e_t")
    st.caption(f"{len(sensitivity.factores(p))} factores × 2 valores × {int(n_t):,} réplicas × {len(EST)} estrategias.".replace(",", "."))
    if st.button("Calcular tornado", type="primary"):
        barra = st.progress(0.0)
        st.session_state["tornado"] = sensitivity.tornado(p, EST, n=int(n_t), semilla=p.i("exp.semilla"),
                                                          progreso=barra.progress, procesos=min(4, NUC))
    tor = st.session_state.get("tornado")
    if tor is not None:
        st.plotly_chart(charts.tornado(tor, e_t, medida), width="stretch", config=charts.CONFIG)

with t3:
    if len(res_sesion) < 2:
        st.info("Se necesitan al menos dos escenarios ejecutados en Experimento.")
    else:
        res = pd.concat([d["resumen"] for d in res_sesion.values()], ignore_index=True)
        medidas = {"p_meta": "Probabilidad de meta (%)", "p_iliquidez": "Probabilidad de iliquidez (%)",
                   "W_real_p50": "Patrimonio real mediano (M)", "caida_max_p50": "Caída máxima mediana (%)"}
        m = st.selectbox("Indicador", list(medidas), format_func=medidas.get)
        esc_v = 1e-6 if m == "W_real_p50" else 100
        fig = go.Figure()
        for esc, g in res.groupby("escenario", sort=False):
            fig.add_trace(go.Bar(x=[ui.NOMBRES[e] for e in g.estrategia], y=g[m] * esc_v, name=ui.ESCENARIOS[esc]))
        fig.update_layout(barmode="group", title=medidas[m], **{k: v for k, v in charts.LAYOUT.items() if k != "hovermode"})
        st.plotly_chart(fig, width="stretch", config=charts.CONFIG)

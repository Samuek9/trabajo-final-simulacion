import os

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from components import charts, estado, fuente, ui
from investment_sim.experiments import capital_goal, sensitivity

st.title("Sensibilidad y capital-meta")
p = estado.params()
NUC = os.cpu_count() or 1
EST = ["E1", "E2", "E3", "E4"]


def lista_millones(txt: str) -> list[float]:
    return [float(x.strip()) * 1e6 for x in txt.replace(";", ",").split(",") if x.strip()]


t1, t2, t3, t4 = st.tabs(["Capital-meta", "Tornado", "Elección de modelos", "Escenarios"])

with t1:
    st.write("¿Qué probabilidad hay de llegar a cada meta según el capital inicial? Cada celda corre todas las réplicas con ese "
             "capital (punto 17). El capital necesario es el que da la probabilidad objetivo, interpolando entre capitales.")
    c = st.columns(4)
    with c[0]:
        caps = st.text_input("Capitales (millones)", "400, 500, 600, 700, 800, 900, 1000")
    with c[1]:
        metas = st.text_input("Metas reales (millones)", "700, 800, 1000, 1200, 1500")
    with c[2]:
        n_cm = st.number_input("Réplicas por celda", 50, 5000, 300, 50, key="n_cm")
    with c[3]:
        esc_cm = st.selectbox("Escenario", list(ui.ESCENARIOS), format_func=ui.ESCENARIOS.get, key="esc_cm")
    if st.button("Calcular capital-meta", type="primary"):
        barra = st.progress(0.0)
        st.session_state["cm"] = capital_goal.matriz(estado.params(esc_cm), lista_millones(caps), lista_millones(metas), EST,
                                                    n=int(n_cm), semilla=p.i("exp.semilla"), procesos=min(4, NUC),
                                                    progreso=barra.progress)
    cm = st.session_state.get("cm")
    origen = "Cálculo de esta sesión"
    if cm is None and fuente.guardado("capital_meta.csv") is not None:
        cm = fuente.guardado("capital_meta.csv")
        cm = cm[cm.escenario == esc_cm] if esc_cm in set(cm.escenario) else cm[cm.escenario == "base"]
        origen = "Corrida del informe (500 réplicas por celda, parámetros de referencia)"
    if cm is not None and len(cm):
        st.caption(f"{origen} · escenario {ui.ESCENARIOS[cm.escenario.iloc[0]]}")
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
        st.caption("«Fuera del rango» quiere decir que ni el capital más alto probado alcanza la probabilidad, o que el más bajo ya la supera.")

with t2:
    st.write("Cada barra mueve un factor a un valor bajo y uno alto y deja todo lo demás en su valor de referencia. "
             "Las barras más largas son las fuentes de incertidumbre que más pesan en el resultado.")
    c = st.columns(3)
    with c[0]:
        n_t = st.number_input("Réplicas por punto", 50, 3000, 200, 50, key="n_t")
    with c[1]:
        medida = st.radio("Medir", ["p_meta", "W"], format_func={"p_meta": "Probabilidad de meta", "W": "Patrimonio real mediano"}.get,
                          horizontal=True)
    with c[2]:
        e_t = st.selectbox("Estrategia", EST, format_func=ui.NOMBRES.get, key="e_t")
    st.caption(f"{len(sensitivity.factores(p))} factores × 2 valores × {int(n_t)} réplicas × 4 estrategias.")
    if st.button("Calcular tornado", type="primary"):
        barra = st.progress(0.0)
        st.session_state["tornado"] = sensitivity.tornado(p, EST, n=int(n_t), semilla=p.i("exp.semilla"),
                                                          progreso=barra.progress, procesos=min(4, NUC))
    tor = st.session_state.get("tornado")
    if tor is None:
        tor = fuente.guardado("tornado.csv")
        if tor is not None:
            st.caption("Corrida del informe (400 réplicas por punto, parámetros de referencia)")
    if tor is not None:
        st.plotly_chart(charts.tornado(tor, e_t, medida), width="stretch", config=charts.CONFIG)

with t3:
    st.write("Para varios procesos hubo más de un modelo candidato. Aquí se cambia uno a la vez por su alternativa "
             "para ver si la conclusión depende de esa elección.")
    sm = fuente.guardado("sensibilidad_modelos.csv")
    if st.button("Calcular con los parámetros actuales"):
        with st.spinner("Corriendo..."):
            st.session_state["sm"] = sensitivity.sensibilidad_modelos(p, EST, n=300, semilla=p.i("exp.semilla"), procesos=min(4, NUC))
    sm = st.session_state.get("sm", sm)
    if sm is not None:
        q = sm.copy()
        q["Cambio en prob. de meta (puntos)"] = ((q.alt_p_meta - q.ref_p_meta) * 100).round(1)
        q["Cambio en patrimonio mediano (M)"] = ((q.alt_W - q.ref_W) / 1e6).round(1)
        piv = q.pivot(index="modelo", columns="estrategia", values="Cambio en prob. de meta (puntos)")
        st.markdown("**Cambio en la probabilidad de meta (puntos) frente al modelo de referencia**")
        st.dataframe(piv, width="stretch")
        piv2 = q.pivot(index="modelo", columns="estrategia", values="Cambio en patrimonio mediano (M)")
        st.markdown("**Cambio en el patrimonio real mediano (millones)**")
        st.dataframe(piv2, width="stretch")

with t4:
    res = None
    if st.session_state.get("resultados") and len(st.session_state["resultados"]) > 1:
        res = pd.concat([d["resumen"] for d in st.session_state["resultados"].values()], ignore_index=True)
        st.caption("Ejecución de esta sesión")
    elif fuente.guardado("resumen.csv") is not None:
        res = fuente.guardado("resumen.csv")
        st.caption("Corrida del informe")
    if res is None:
        st.info("Ejecuta al menos dos escenarios en Experimento.")
    else:
        medidas = {"p_meta": "Probabilidad de meta (%)", "p_iliquidez": "Probabilidad de iliquidez (%)",
                   "W_real_p50": "Patrimonio real mediano (M)", "caida_max_p50": "Caída máxima mediana (%)"}
        m = st.selectbox("Indicador", list(medidas), format_func=medidas.get)
        esc_v = 1e-6 if m == "W_real_p50" else 100
        fig = go.Figure()
        for esc, g in res.groupby("escenario", sort=False):
            fig.add_trace(go.Bar(x=[ui.NOMBRES[e] for e in g.estrategia], y=g[m] * esc_v, name=ui.ESCENARIOS[esc]))
        fig.update_layout(barmode="group", title=medidas[m], **{k: v for k, v in charts.LAYOUT.items() if k != "hovermode"})
        st.plotly_chart(fig, width="stretch", config=charts.CONFIG)

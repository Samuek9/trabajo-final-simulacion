import io
import json

import pandas as pd
import streamlit as st

from components import charts, estado, fuente, ui

st.title("Resultados")

if not st.session_state.get("resultados"):
    st.info("No hay resultados. Ejecute una simulación en Experimento.")
    st.page_link("pages/05_experimento.py", label="Ir al experimento")
    st.stop()

escs = fuente.escenarios()
esc = st.radio("Escenario", escs, format_func=ui.ESCENARIOS.get, horizontal=True) if len(escs) > 1 else escs[0]
d = fuente.datos(esc)
res = d["resumen"].set_index("estrategia")
p = d["params"]
meta = p.f("inv.meta_real")
st.caption(d["nota"] + f" · capital {ui.cop(p.f('inv.capital'))} · meta real {ui.cop(meta)} · {p.i('inv.horizonte') // 12} años")


def ic(r, col, f=ui.pct):
    return f"IC 95 %: {f(r[col + '_lo'])} a {f(r[col + '_hi'])}"


st.markdown("##### Probabilidad de alcanzar la meta")
cols = st.columns(len(res))
for col, (e, r) in zip(cols, res.iterrows()):
    with col:
        ui.tarjeta(ui.NOMBRES[e], ui.pct(r.p_meta), ic(r, "p_meta"), ui.COLORES[e])
st.markdown("##### Patrimonio real mediano al final")
cols = st.columns(len(res))
for col, (e, r) in zip(cols, res.iterrows()):
    with col:
        ui.tarjeta(ui.NOMBRES[e], ui.cop(r.W_real_p50), ic(r, "W_real_p50", ui.cop), ui.COLORES[e])
st.markdown("##### Riesgo de iliquidez")
cols = st.columns(len(res))
for col, (e, r) in zip(cols, res.iterrows()):
    with col:
        ui.tarjeta(ui.NOMBRES[e], ui.pct(r.p_iliquidez), f"réplicas con algún episodio · venta urgente {ui.pct(r.p_venta_urgente)}", ui.COLORES[e])

t1, t2, t3, t4, t5 = st.tabs(["Evolución del patrimonio", "Patrimonio final", "Riesgo y meta", "Indicadores completos", "Comparaciones pareadas"])
with t1:
    if d["bandas"]:
        st.plotly_chart(charts.abanico_bandas(d["bandas"], meta), width="stretch", config=charts.CONFIG)
        st.caption("Mediana y bandas de percentiles 25-75 y 5-95.")
    if d["iliq"]:
        st.plotly_chart(charts.serie_estrategias(d["iliq"], "Réplicas en iliquidez cada mes", "% de réplicas"),
                        width="stretch", config=charts.CONFIG)
with t2:
    st.plotly_chart(charts.distribucion_final(d["tabla"], "W_real", meta), width="stretch", config=charts.CONFIG)
with t3:
    st.plotly_chart(charts.riesgo_rentabilidad(d["resumen"]), width="stretch", config=charts.CONFIG)
with t4:
    st.caption("Grupos 1 a 3: indicadores mínimos del enunciado. Grupos 4 y 5: adicionales.")
    pct1, n1 = ui.pct, (lambda x: ui.num(x, 1))
    grupos = {
        "1. Patrimonio y rentabilidad": {
            "Patrimonio nominal medio": ("W_nominal_media", ui.cop), "Patrimonio real medio": ("W_real_media", ui.cop),
            "Patrimonio real, percentil 5": ("W_real_p5", ui.cop), "Patrimonio real, percentil 25": ("W_real_p25", ui.cop),
            "Patrimonio real, percentil 50": ("W_real_p50", ui.cop), "Patrimonio real, percentil 75": ("W_real_p75", ui.cop),
            "Patrimonio real, percentil 95": ("W_real_p95", ui.cop),
            "Patrimonio nominal, percentil 5": ("W_nominal_p5", ui.cop), "Patrimonio nominal, percentil 25": ("W_nominal_p25", ui.cop),
            "Patrimonio nominal, percentil 50": ("W_nominal_p50", ui.cop), "Patrimonio nominal, percentil 75": ("W_nominal_p75", ui.cop),
            "Patrimonio nominal, percentil 95": ("W_nominal_p95", ui.cop),
            "Rentabilidad real anual media": ("rent_real_media", ui.pct), "Rentabilidad nominal anual media": ("rent_nominal_media", ui.pct),
            "Variabilidad: desviación del patrimonio real": ("W_real_sd", ui.cop)},
        "2. Riesgo y meta": {
            "Pérdida máxima durante la trayectoria (mediana)": ("caida_max_p50", ui.pct),
            "Pérdida máxima durante la trayectoria (percentil 95)": ("caida_max_p95", ui.pct),
            "Probabilidad de terminar con patrimonio real inferior al inicial": ("p_perdida_real", ui.pct),
            "Probabilidad de alcanzar la meta": ("p_meta", ui.pct),
            "Momento de primer cumplimiento de la meta (años, mediana, si la alcanza)": ("t_meta_p50_anios", n1)},
        "3. Liquidez": {
            "Probabilidad de algún episodio de iliquidez": ("p_iliquidez", ui.pct),
            "Episodios de iliquidez por réplica (frecuencia media)": ("episodios_iliquidez_media", lambda x: ui.num(x, 3)),
            "Duración de un episodio de iliquidez (meses, media)": ("duracion_iliquidez_media", n1),
            "Meses en iliquidez por réplica (media)": ("meses_iliquidez_media", lambda x: ui.num(x, 2))},
        "4. Adicionales de riesgo": {
            "CVaR 5 % (media del peor 5 % del patrimonio real)": ("cvar5_W_real", ui.cop),
            "Probabilidad de venta urgente del inmueble": ("p_venta_urgente", ui.pct),
            "Probabilidad de insolvencia": ("p_insolvencia", ui.pct)},
        "5. Adicionales de desempeño": {
            "Necesidades de liquidez (media)": ("n_necesidades_media", lambda x: ui.num(x, 2)),
            "Costos totales (media)": ("costos_totales_media", ui.cop),
            "Participación del arriendo en los ingresos corrientes": ("participacion_arriendo_media", ui.pct),
            "Rebalanceos (media)": ("n_rebalanceos_media", lambda x: ui.num(x, 2))},
    }
    filas = {(g, k): {ui.NOMBRES[e]: f(r[c]) for e, r in res.iterrows() if c in r}
             for g, ks in grupos.items() for k, (c, f) in ks.items()}
    tabla = pd.DataFrame(filas).T
    tabla.index = pd.MultiIndex.from_tuples(tabla.index, names=["Grupo", "Indicador"])
    st.dataframe(tabla, width="stretch", height=640)
    tabla = tabla.reset_index()
    with st.expander("Liquidez frente a patrimonio"):
        st.write("Una estrategia puede tener patrimonio alto e iliquidez: el inmueble vale, pero no se vende rápido. Cada necesidad "
                 "extraordinaria se cubre en este orden: cuenta, ETF, CDT (con castigo por venta anticipada) y venta urgente del "
                 "inmueble con descuento. Lo no cubierto queda como deuda y cuenta como episodio de iliquidez.")
with t5:
    pares = d["pares"]
    if pares is not None and len(pares):
        nombres_var = {"W_real": "Patrimonio real (M)", "meta": "Prob. de meta (puntos)",
                       "algun_episodio": "Prob. de iliquidez (puntos)", "caida_max": "Caída máxima (puntos)"}
        v = st.selectbox("Variable", list(nombres_var), format_func=nombres_var.get)
        esc_v = 1e-6 if v == "W_real" else 100
        q = pares[pares.variable == v].copy()
        q["Comparación"] = q.a + " − " + q.b
        q["Diferencia"] = (q.diferencia * esc_v).round(2)
        q["IC 95 %"] = [f"{lo * esc_v:.2f} a {hi * esc_v:.2f}" for lo, hi in zip(q.lo, q.hi)]
        q["¿Distinta de cero?"] = q.significativa.map({True: "Sí", False: "No"})
        q["Reducción de varianza por CRN"] = q.reduccion_varianza.map(lambda x: ui.pct(x, 0))
        st.dataframe(q[["Comparación", "Diferencia", "IC 95 %", "¿Distinta de cero?", "Reducción de varianza por CRN"]],
                     hide_index=True, width="stretch")
        st.caption("Diferencias réplica a réplica con números aleatorios comunes.")

st.markdown("##### Descargas")
c = st.columns(4)
with c[0]:
    st.download_button("Réplicas (CSV)", d["tabla"].to_csv(index=False).encode("utf-8"), f"replicas_{esc}.csv", "text/csv")
with c[1]:
    st.download_button("Resumen (CSV)", d["resumen"].to_csv(index=False).encode("utf-8"), f"resumen_{esc}.csv", "text/csv")
with c[2]:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        tabla.to_excel(xw, sheet_name="Indicadores", index=False)
        d["resumen"].to_excel(xw, sheet_name="Resumen", index=False)
        if d["pares"] is not None and len(d["pares"]):
            d["pares"].to_excel(xw, sheet_name="Pareadas", index=False)
        pd.DataFrame([{"clave": k, "valor": str(v)} for k, v in p.valores.items()]).to_excel(xw, sheet_name="Parametros", index=False)
    st.download_button("Resumen (Excel)", buf.getvalue(), f"resultados_{esc}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
with c[3]:
    conf = {"escenario": esc, "cambios_frente_a_valores_iniciales": dict(p.cambios),
            "valores": {k: (v if isinstance(v, (int, float, str)) else str(v)) for k, v in p.valores.items()}}
    st.download_button("Parámetros usados (JSON)", json.dumps(conf, indent=2, ensure_ascii=False).encode("utf-8"),
                       f"parametros_{esc}.json", "application/json")

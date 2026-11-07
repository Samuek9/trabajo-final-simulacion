import io
import json

import pandas as pd
import streamlit as st

from components import charts, estado, fuente, ui

st.title("Resultados")

opciones = {}
if st.session_state.get("resultados"):
    opciones["sesion"] = "Mi ejecución"
if fuente.hay_guardados():
    opciones["informe"] = "Corrida del informe"
if not opciones:
    st.info("No hay resultados todavía. Ve a Experimento y ejecuta la simulación.")
    st.stop()

c1, c2 = st.columns(2)
with c1:
    origen = st.radio("Resultados de", list(opciones), format_func=opciones.get, horizontal=True)
with c2:
    escs = fuente.escenarios(origen)
    esc = st.radio("Escenario", escs, format_func=ui.ESCENARIOS.get, horizontal=True)

d = fuente.datos(origen, esc)
res = d["resumen"].set_index("estrategia")
p = d["params"] or estado.referencia(esc)
meta = p.f("inv.meta_real")
st.caption(d["nota"] + f" · capital {ui.cop(p.f('inv.capital'))} · meta real {ui.cop(meta)} · {p.i('inv.horizonte') // 12} años")


def ic(r, col, f=ui.pct):
    return f"IC 95 %: {f(r[col + '_lo'])} a {f(r[col + '_hi'])}"


st.markdown("### Probabilidad de alcanzar la meta")
cols = st.columns(len(res))
for col, (e, r) in zip(cols, res.iterrows()):
    with col:
        ui.tarjeta(ui.NOMBRES[e], ui.pct(r.p_meta), ic(r, "p_meta"), ui.COLORES[e])
st.markdown("### Patrimonio real mediano al final")
cols = st.columns(len(res))
for col, (e, r) in zip(cols, res.iterrows()):
    with col:
        ui.tarjeta(ui.NOMBRES[e], ui.cop(r.W_real_p50), ic(r, "W_real_p50", ui.cop), ui.COLORES[e])
st.markdown("### Riesgo de liquidez")
cols = st.columns(len(res))
for col, (e, r) in zip(cols, res.iterrows()):
    with col:
        ui.tarjeta(ui.NOMBRES[e], ui.pct(r.p_iliquidez), f"réplicas con algún episodio · venta urgente {ui.pct(r.p_venta_urgente)}",
                   ui.COLORES[e])

t1, t2, t3, t4, t5 = st.tabs(["Trayectorias", "Distribución final", "Meta vs. iliquidez", "Tabla completa", "Comparaciones pareadas"])
with t1:
    if d["bandas"]:
        st.plotly_chart(charts.abanico_bandas(d["bandas"], meta), width="stretch", config=charts.CONFIG)
        st.caption("Línea: mediana. Bandas: percentiles 25 a 75 y 5 a 95 de las réplicas, mes a mes.")
    if d["iliq"]:
        st.plotly_chart(charts.serie_estrategias(d["iliq"], "Réplicas en iliquidez cada mes", "% de réplicas"),
                        width="stretch", config=charts.CONFIG)
with t2:
    st.plotly_chart(charts.distribucion_final(d["tabla"], "W_real", meta), width="stretch", config=charts.CONFIG)
with t3:
    st.plotly_chart(charts.riesgo_rentabilidad(d["resumen"]), width="stretch", config=charts.CONFIG)
with t4:
    filas = {
        "Probabilidad de meta": ("p_meta", ui.pct), "Patrimonio real medio": ("W_real_media", ui.cop),
        "Patrimonio real P5": ("W_real_p5", ui.cop), "Patrimonio real mediano": ("W_real_p50", ui.cop),
        "Patrimonio real P95": ("W_real_p95", ui.cop), "CVaR 5 % (media del peor 5 %)": ("cvar5_W_real", ui.cop),
        "Rentabilidad real anual media": ("rent_real_media", ui.pct), "Rentabilidad nominal anual media": ("rent_nominal_media", ui.pct),
        "Probabilidad de pérdida real": ("p_perdida_real", ui.pct), "Caída máxima mediana": ("caida_max_p50", ui.pct),
        "Caída máxima P95": ("caida_max_p95", ui.pct), "Probabilidad de iliquidez": ("p_iliquidez", ui.pct),
        "Meses en iliquidez (media)": ("meses_iliquidez_media", lambda x: ui.num(x, 2)),
        "Probabilidad de venta urgente del inmueble": ("p_venta_urgente", ui.pct),
        "Probabilidad de insolvencia": ("p_insolvencia", ui.pct),
        "Años hasta la meta (mediana, si la alcanza)": ("t_meta_p50_anios", lambda x: ui.num(x, 1)),
        "Necesidades de liquidez (media)": ("n_necesidades_media", lambda x: ui.num(x, 2)),
        "Costos totales (media)": ("costos_totales_media", ui.cop),
        "Participación del arriendo en el retorno": ("participacion_arriendo_media", ui.pct),
        "Rebalanceos (media)": ("n_rebalanceos_media", lambda x: ui.num(x, 2)),
    }
    tabla = pd.DataFrame({ui.NOMBRES[e]: {k: f(r[c]) for k, (c, f) in filas.items() if c in r} for e, r in res.iterrows()})
    st.dataframe(tabla, width="stretch")
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
        st.caption("Diferencias réplica a réplica. Como las estrategias comparten la economía de cada réplica, el intervalo es "
                   "mucho más angosto que si se compararan corridas independientes.")

st.markdown("### Descargas")
c = st.columns(4)
with c[0]:
    st.download_button("Réplicas (CSV)", d["tabla"].to_csv(index=False).encode("utf-8"), f"replicas_{esc}.csv", "text/csv")
with c[1]:
    st.download_button("Resumen (CSV)", d["resumen"].to_csv(index=False).encode("utf-8"), f"resumen_{esc}.csv", "text/csv")
with c[2]:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        tabla.to_excel(xw, sheet_name="Indicadores")
        d["resumen"].to_excel(xw, sheet_name="Resumen", index=False)
        if d["pares"] is not None and len(d["pares"]):
            d["pares"].to_excel(xw, sheet_name="Pareadas", index=False)
        pd.DataFrame([{"clave": k, "valor": str(v)} for k, v in p.valores.items()]).to_excel(xw, sheet_name="Parametros", index=False)
    st.download_button("Resumen (Excel)", buf.getvalue(), f"resultados_{esc}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
with c[3]:
    conf = {"escenario": esc, "cambios_frente_a_referencia": dict(p.cambios),
            "valores": {k: (v if isinstance(v, (int, float, str)) else str(v)) for k, v in p.valores.items()}}
    st.download_button("Parámetros usados (JSON)", json.dumps(conf, indent=2, ensure_ascii=False).encode("utf-8"),
                       f"parametros_{esc}.json", "application/json")
st.caption("Cada gráfico se puede bajar en PNG con el ícono de cámara que aparece al pasar el mouse.")

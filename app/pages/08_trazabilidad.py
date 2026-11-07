import pandas as pd
import streamlit as st

from components import charts, estado, ui
from investment_sim.experiments.runner import una_replica

st.title("Trazabilidad de una réplica")
st.write("Corre una sola réplica y muestra cada evento en el orden en que lo procesó el motor: hora simulada, número del "
         "evento en el diagrama, lo que pasó y cómo quedaron la cuenta y el patrimonio. Sirve para revisar a mano que las "
         "reglas se cumplen.")

c = st.columns(3)
with c[0]:
    esc = st.selectbox("Escenario", list(ui.ESCENARIOS), format_func=ui.ESCENARIOS.get)
with c[1]:
    e = st.selectbox("Estrategia", ["E1", "E2", "E3", "E4"], index=3, format_func=ui.NOMBRES.get)
with c[2]:
    r = st.number_input("Réplica", 0, 1_000_000, 0, help="Usa la semilla base + réplica, la misma que en el experimento.")

p = estado.params(esc)
semilla = p.i("exp.semilla") + int(r)
rep = una_replica(p, e, semilla, traza=True)
ind = rep.indicadores

k = st.columns(4)
with k[0]:
    ui.tarjeta("Patrimonio real final", ui.cop(ind["W_real"]), "meta alcanzada" if ind["meta"] else "no alcanza la meta", ui.COLORES[e])
with k[1]:
    ui.tarjeta("Rentabilidad real anual", ui.pct(ind["rent_real"]), f"nominal {ui.pct(ind['rent_nominal'])}")
with k[2]:
    ui.tarjeta("Caída máxima", ui.pct(ind["caida_max"]), f"{int(ind['n_necesidades'])} necesidades de liquidez")
with k[3]:
    ui.tarjeta("Meses en iliquidez", str(int(ind["meses_iliquidez"])), f"{int(ind['n_ventas_inmueble'])} venta(s) del inmueble")

st.plotly_chart(charts.trayectoria_replica(rep.trayectoria), width="stretch", config=charts.CONFIG)

traza = pd.DataFrame(rep.traza)
nombres = sorted(traza.evento.unique())
filtro = st.multiselect("Eventos", nombres, default=[x for x in nombres if x not in ("Cierre de mes", "Pago de arriendo")],
                        help="Por defecto se ocultan los cierres de mes y los pagos de arriendo, que son muchos.")
vista = traza[traza.evento.isin(filtro)].copy()
vista["cuenta"] = (vista.cuenta / 1e6).round(2)
vista["W"] = (vista.W / 1e6).round(2)
vista = vista.rename(columns={"t": "Mes", "n": "N.º", "evento": "Evento", "detalle": "Qué pasó", "estado": "Estado",
                              "cuenta": "Cuenta (M)", "W": "Patrimonio (M)"})
st.dataframe(vista, hide_index=True, width="stretch", height=480)
st.caption(f"{len(traza):,} eventos en total; se muestran {len(vista):,}. Semilla {semilla}.".replace(",", "."))
st.download_button("Descargar la traza completa (CSV)", traza.to_csv(index=False).encode("utf-8"),
                   f"traza_{esc}_{e}_r{r}.csv", "text/csv")
st.download_button("Descargar la trayectoria mensual (CSV)",
                   pd.DataFrame(rep.trayectoria).rename_axis("mes").to_csv().encode("utf-8"),
                   f"trayectoria_{esc}_{e}_r{r}.csv", "text/csv")

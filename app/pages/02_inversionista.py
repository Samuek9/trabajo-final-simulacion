import streamlit as st

from components import estado, ui
from investment_sim.config.parameters import advertencias

p = estado.params()
T = p.tabla

st.title("Inversionista")
st.write("Ningún valor es fijo: el capital, la meta, el horizonte y la liquidez mínima son entradas del simulador.")

c1, c2 = st.columns(2)
with c1:
    capital = st.number_input("Capital inicial (millones de COP)", min_value=50.0, max_value=20_000.0, step=50.0,
                              value=p.f("inv.capital") / 1e6, help=ui.pie_parametro(T["inv.capital"]))
    anios = st.slider("Horizonte (años)", 1, 30, p.i("inv.horizonte") // 12, help=ui.pie_parametro(T["inv.horizonte"]))
    meta = st.number_input("Meta: patrimonio real al final (millones de pesos de hoy)", min_value=0.0, max_value=50_000.0,
                           step=50.0, value=p.f("inv.meta_real") / 1e6, help=ui.pie_parametro(T["inv.meta_real"]))
with c2:
    liq = st.slider("Liquidez mínima en la cuenta (% del patrimonio)", 0.0, 30.0, p.f("inv.liquidez_min") * 100, 0.5,
                    help=ui.pie_parametro(T["inv.liquidez_min"]))
    modo = st.radio("Si cambia el capital, el inmueble...", ["proporcional", "fijo"],
                    index=0 if p.s("inv.modo_inmueble") == "proporcional" else 1, horizontal=True,
                    format_func=lambda x: "cambia en proporción (43 % o 55 % del capital)" if x == "proporcional" else "mantiene un valor fijo")
    fijo = st.number_input("Valor fijo del inmueble (millones)", min_value=50.0, max_value=5_000.0, step=10.0,
                           value=p.f("inv.valor_inmueble_fijo") / 1e6, disabled=modo == "proporcional",
                           help=ui.pie_parametro(T["inv.valor_inmueble_fijo"]))
    if modo == "fijo" and fijo >= capital:
        st.error("El inmueble no puede costar más que el capital.")

st.markdown("##### Necesidades extraordinarias de liquidez")
st.write("Gastos imprevistos que obligan a conseguir efectivo. Llegan al azar con la frecuencia que se indique (0 = sin necesidades) "
         "y su monto es un porcentaje del capital inicial, en pesos de hoy; el valor mediano es el que se pide aquí y algunos "
         "son bastante mayores.")
c5 = st.columns(4)
meses = {}
for col, (L, nombre) in zip(c5[:3], (("E", "expansión"), ("N", "normal"), ("C", "crisis"))):
    tasa = p.f(f"liq.tasa_{L}")
    with col:
        meses[L] = st.number_input(f"Una necesidad cada ... meses ({nombre})", min_value=0.0, step=1.0,
                                   value=float(round(1 / tasa, 2)) if tasa > 0 else 0.0, key=f"liq_{L}")
with c5[3]:
    monto = st.number_input("Monto mediano (% del capital)", min_value=0.0, max_value=50.0, step=0.5,
                            value=round(p.f("liq.mediana_pct") * 100, 2), help=ui.pie_parametro(T["liq.mediana_pct"]))

st.markdown("##### Experimento")
c3, c4 = st.columns(2)
with c3:
    n = st.number_input("Número de réplicas", 20, 20_000, p.i("exp.replicas"), step=50, help=ui.pie_parametro(T["exp.replicas"]))
with c4:
    semilla = st.number_input("Semilla base", 0, 2**31 - 1, p.i("exp.semilla"), help="La réplica r usa la semilla base + r.")

cambios = {"inv.capital": capital * 1e6, "inv.horizonte": anios * 12, "inv.meta_real": meta * 1e6,
           "inv.liquidez_min": liq / 100, "inv.modo_inmueble": modo, "inv.valor_inmueble_fijo": fijo * 1e6,
           "exp.replicas": int(n), "exp.semilla": int(semilla), "liq.mediana_pct": monto / 100,
           **{f"liq.tasa_{L}": (1 / m if m > 0 else 0.0) for L, m in meses.items()}}
error = estado.probar(cambios)
if error:
    st.error(error)
elif not (modo == "fijo" and fijo >= capital):
    for k, v in cambios.items():
        if abs(float(v) - float(p.valores[k])) > 1e-12 if isinstance(v, (int, float)) else v != p.valores[k]:
            estado.fijar(k, v)

if st.button("Restablecer esta página"):
    estado.restablecer(("inv.", "exp.", "liq."))
    st.rerun()

for a in advertencias(estado.params()):
    st.warning(a)

error_meta = 1.96 * (0.25 / max(int(n), 1)) ** 0.5
st.caption(f"Con {int(n):,} réplicas la probabilidad de meta tiene un error de ±{error_meta * 100:.1f} puntos en el peor caso (p = 50 %).".replace(",", "."))

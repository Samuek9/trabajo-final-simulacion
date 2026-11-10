import streamlit as st

from components import estado, ui
from investment_sim.config.parameters import advertencias

p = estado.params()
T = p.tabla

st.title("Inversionista")
ui.subtitulo("Capital, objetivo y necesidades de liquidez. Los parámetros del experimento se definen en la página Experimento.")

st.markdown("##### Capital y objetivo")
c1, c2 = st.columns(2)
with c1:
    capital = st.number_input("Capital inicial (millones de COP)", min_value=1.0, max_value=100_000.0, step=10.0,
                              value=p.f("inv.capital") / 1e6, help=ui.pie_parametro(T["inv.capital"]))
    anios = st.slider("Horizonte (años)", 1, 30, p.i("inv.horizonte") // 12, help=ui.pie_parametro(T["inv.horizonte"]))
    meta = st.number_input("Meta: patrimonio real final (millones de pesos de hoy)", min_value=0.0, max_value=50_000.0,
                           step=50.0, value=p.f("inv.meta_real") / 1e6, help=ui.pie_parametro(T["inv.meta_real"]))
with c2:
    liq = st.number_input("Liquidez mínima en la cuenta (% del patrimonio)", min_value=0.0, max_value=30.0, step=0.5,
                          value=p.f("inv.liquidez_min") * 100, help=ui.pie_parametro(T["inv.liquidez_min"]))
    modo = st.radio("Valor del inmueble", ["proporcional", "fijo"],
                    index=0 if p.s("inv.modo_inmueble") == "proporcional" else 1, horizontal=True,
                    format_func=lambda x: "Proporcional al capital" if x == "proporcional" else "Valor fijo",
                    help="Proporcional: el inmueble cuesta el porcentaje del capital que cada estrategia le asigna. "
                         "Fijo: todas las estrategias que compran inmueble pagan el mismo valor.")
    fijo = st.number_input("Valor fijo del inmueble (millones de COP)", min_value=1.0, max_value=100_000.0, step=10.0,
                           value=p.f("inv.valor_inmueble_fijo") / 1e6, disabled=modo == "proporcional",
                           help=ui.pie_parametro(T["inv.valor_inmueble_fijo"]))
    if modo == "fijo" and fijo >= capital:
        st.error("El inmueble no puede costar más que el capital.")

st.markdown("##### Necesidades extraordinarias de liquidez")
st.caption("Supuesto del inversionista. 0 meses equivale a sin necesidades.")
c5 = st.columns(4)
meses = {}
for col, (L, nombre) in zip(c5[:3], (("E", "expansión"), ("N", "normal"), ("C", "crisis"))):
    tasa = p.f(f"liq.tasa_{L}")
    with col:
        meses[L] = st.number_input(f"Una cada … meses ({nombre})", min_value=0.0, step=1.0,
                                   value=float(round(1 / tasa, 2)) if tasa > 0 else 0.0, key=f"liq_{L}")
with c5[3]:
    monto = st.number_input("Monto mediano (% del capital)", min_value=0.0, max_value=50.0, step=0.5,
                            value=round(p.f("liq.mediana_pct") * 100, 2), help=ui.pie_parametro(T["liq.mediana_pct"]))

cambios = {"inv.capital": capital * 1e6, "inv.horizonte": anios * 12, "inv.meta_real": meta * 1e6,
           "inv.liquidez_min": liq / 100, "inv.modo_inmueble": modo, "inv.valor_inmueble_fijo": fijo * 1e6,
           "liq.mediana_pct": monto / 100, **{f"liq.tasa_{L}": (1 / m if m > 0 else 0.0) for L, m in meses.items()}}
error = estado.probar(cambios)
if error:
    st.error(error)
elif not (modo == "fijo" and fijo >= capital):
    for k, v in cambios.items():
        if abs(float(v) - float(p.valores[k])) > 1e-12 if isinstance(v, (int, float)) else v != p.valores[k]:
            estado.fijar(k, v)

if st.button("Restablecer valores iniciales"):
    estado.restablecer(("inv.", "liq."))
    st.rerun()
for a in advertencias(estado.params()):
    st.warning(a)

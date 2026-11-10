import streamlit as st

from components import estado, ui
from investment_sim.config.parameters import advertencias

p = estado.params()
T = p.tabla

st.title("Inversionista")
st.write("Ningún valor es fijo: el capital, la meta, el horizonte y la liquidez mínima son entradas del simulador.")

c1, c2 = st.columns(2)
with c1:
    capital = st.number_input("Capital inicial (millones de COP)", min_value=1.0, max_value=100_000.0, step=10.0,
                              value=p.f("inv.capital") / 1e6, help=ui.pie_parametro(T["inv.capital"]))
    anios = st.slider("Horizonte (años)", 1, 30, p.i("inv.horizonte") // 12, help=ui.pie_parametro(T["inv.horizonte"]))
    meta = st.number_input("Meta: patrimonio real al final (millones de pesos de hoy)", min_value=0.0, max_value=50_000.0,
                           step=50.0, value=p.f("inv.meta_real") / 1e6, help=ui.pie_parametro(T["inv.meta_real"]))
with c2:
    liq = st.slider("Liquidez mínima en la cuenta (% del patrimonio)", 0.0, 30.0, p.f("inv.liquidez_min") * 100, 0.5,
                    help=ui.pie_parametro(T["inv.liquidez_min"]))
    modo = st.radio("Valor del inmueble", ["proporcional", "fijo"],
                    index=0 if p.s("inv.modo_inmueble") == "proporcional" else 1, horizontal=True,
                    format_func=lambda x: "proporcional al capital (según el peso del inmueble en cada estrategia)" if x == "proporcional" else "valor fijo que usted indica",
                    help="Con valor proporcional, el inmueble cuesta el porcentaje del capital que cada estrategia le asigna (pestaña Estrategias). "
                         "Con valor fijo, todas las estrategias que compran inmueble pagan lo mismo, sin importar el capital.")
    fijo = st.number_input("Valor fijo del inmueble (millones)", min_value=1.0, max_value=100_000.0, step=10.0,
                           value=p.f("inv.valor_inmueble_fijo") / 1e6, disabled=modo == "proporcional",
                           help=ui.pie_parametro(T["inv.valor_inmueble_fijo"]))
    if modo == "fijo" and fijo >= capital:
        st.error("El inmueble no puede costar más que el capital.")

st.markdown("##### Necesidades extraordinarias de liquidez")
st.write("Supuesto del inversionista, no estimado con datos de monto. Gastos imprevistos que obligan a conseguir efectivo. Llegan al azar con la frecuencia que se indique (0 = sin necesidades) "
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

if st.button("Restablecer valores iniciales de esta página"):
    estado.restablecer(("inv.", "exp.", "liq."))
    st.rerun()

for a in advertencias(estado.params()):
    st.warning(a)

error_meta = 1.96 * (0.25 / max(int(n), 1)) ** 0.5
st.caption(f"Con {int(n):,} réplicas la probabilidad de meta tiene un error de ±{error_meta * 100:.1f} puntos en el peor caso (p = 50 %).".replace(",", "."))

st.markdown("##### Réplicas sugeridas para estos parámetros")
st.write("El número de réplicas necesario depende de qué tan incierta sea la meta con sus datos: cuanto más cerca de 50 % esté la "
         "probabilidad de cumplirla, más réplicas hacen falta. Una corrida piloto de 100 réplicas por estrategia estima esa "
         "probabilidad con el capital, el horizonte, la meta y las estrategias que usted configuró.")
cs1, cs2 = st.columns([1, 2])
with cs1:
    margen = st.number_input("Margen de error buscado (± puntos)", 0.5, 10.0, 3.0, 0.5)
with cs2:
    if st.button("Calcular réplicas sugeridas"):
        from investment_sim.experiments.runner import correr
        from investment_sim.metrics.summary import resumen
        with st.spinner("Corriendo el piloto..."):
            rp = correr(estado.params(), ["E1", "E2", "E3", "E4"], n=100, semilla=int(semilla), procesos=1)
        pm = resumen(rp.tabla).set_index("estrategia").p_meta
        peor = float(min(pm, key=lambda v: abs(v - 0.5)))
        sug = int(1.96 ** 2 * peor * (1 - peor) / (margen / 100) ** 2) + 1
        st.session_state["n_sugerido"] = (sug, peor, margen)
if "n_sugerido" in st.session_state:
    sug, peor, mg = st.session_state["n_sugerido"]
    st.info(f"Sugerido: {sug:,} réplicas para ±{mg:g} puntos. En el piloto, la probabilidad de meta más cercana a 50 % fue {peor * 100:.0f} %. "
            "Si cambia capital, horizonte, meta o estrategias, vuelva a calcularlo.".replace(",", "."))

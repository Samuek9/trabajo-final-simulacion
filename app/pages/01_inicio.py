import streamlit as st

from components import estado, ui
from investment_sim.config.parameters import advertencias

p = estado.params()

st.caption("SIMULACIÓN DE EVENTOS DISCRETOS · MODELO ACADÉMICO")
st.title("Investment Lab")
st.markdown("#### Entiende el riesgo antes de elegir una estrategia.")
st.write(
    "El simulador sigue mes a mes el patrimonio de un inversionista que reparte su capital entre una cuenta "
    "remunerada, CDT a 12 meses, el ETF iCOLCAP y un apartamento para arrendar en Medellín. La economía pasa "
    "por expansiones, años normales y crisis; la inflación mueve la tasa del Banco de la República; el inquilino "
    "puede irse; y en cualquier momento puede aparecer un gasto imprevisto que obligue a vender. Cada resultado "
    "sale de correr el modelo, no de una tabla calculada antes."
)

c = st.columns(4)
with c[0]:
    ui.tarjeta("Capital inicial", ui.cop(p.f("inv.capital")), "se cambia en Inversionista")
with c[1]:
    ui.tarjeta("Horizonte", f"{p.i('inv.horizonte') // 12} años", f"{p.i('inv.horizonte')} cierres de mes")
with c[2]:
    ui.tarjeta("Meta real", ui.cop(p.f("inv.meta_real")), "patrimonio al final, en pesos de hoy")
with c[3]:
    ui.tarjeta("Liquidez mínima", ui.pct(p.f("inv.liquidez_min"), 0), "del patrimonio, siempre en la cuenta")

st.markdown("### Cuatro estrategias")
desc = {
    "E1": "Prioriza conservar el capital y tener el dinero a mano. 80 % en cuenta y CDT; no compra inmueble.",
    "E2": "Reparte entre los cuatro activos para que ninguno domine: arriendo, crecimiento y liquidez.",
    "E3": "Busca crecer: 85 % en inmueble y ETF. Acepta caídas y tener casi todo en un activo lento de vender.",
    "E4": "Empieza como E2 y cada año ajusta cuenta, CDT y ETF al estado de la economía.",
}
cols = st.columns(4)
for col, e in zip(cols, ("E1", "E2", "E3", "E4")):
    w = p.lista(f"{e}.pesos")
    with col:
        ui.tarjeta(ui.NOMBRES[e], " · ".join(f"{x * 100:.0f}" for x in w), desc[e], ui.COLORES[e])
        st.caption("cuenta · CDT · ETF · inmueble (%)")

st.markdown("### Cómo usarlo")
st.markdown(
    "1. **Inversionista**: capital, horizonte, meta, liquidez mínima, necesidades de liquidez y réplicas.\n"
    "2. **Modelos y economía**: el modelo elegido para cada proceso aleatorio. Sus valores están estimados con datos y son fijos.\n"
    "3. **Estrategias**: pesos iniciales y reglas de E4.\n"
    "4. **Experimento**: escenario, estrategias, réplicas y semilla; luego *Ejecutar*.\n"
    "5. **Resultados**, **Sensibilidad y capital-meta** y **Simulación, escenarios y validación** para leer lo que pasó y por qué."
)
for a in advertencias(p):
    st.info(a)
st.caption("Vista de una herramienta académica. Los resultados describen el modelo y sus supuestos; no son una recomendación de inversión.")

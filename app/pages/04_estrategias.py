import streamlit as st

from components import estado, ui

p = estado.params()
T = p.tabla
ACTIVOS = ui.ACTIVOS

st.title("Estrategias")
ui.subtitulo("Reparto inicial entre cuenta, CDT, ETF e inmueble. E1, E2 y E3 mantienen sus pesos; E4 los ajusta según el estado de la economía.")

DESC = {"E1": "Conserva el capital y prioriza la liquidez. No compra inmueble.",
        "E2": "Reparte entre los cuatro activos para que ninguno domine.",
        "E3": "Busca crecimiento: concentra el capital en inmueble y ETF.",
        "E4": "Parte del reparto de E2 y rebalancea cuenta, CDT y ETF según el estado de la economía."}


def texto(xs: list[float]) -> str:
    return ";".join(f"{x:g}" for x in xs)


def editor_pesos(clave: str, etiquetas: tuple[str, ...], k: str) -> None:
    actual = p.lista(clave)
    cols = st.columns(len(etiquetas))
    nuevos = []
    for c, et, v, i in zip(cols, etiquetas, actual, range(len(etiquetas))):
        with c:
            nuevos.append(st.number_input(f"{et} (%)", 0.0, 100.0, round(v * 100, 2), 1.0, key=f"{k}_{i}") / 100)
    suma = sum(nuevos)
    if abs(suma - 1) > 1e-6:
        st.error(f"Suma {suma * 100:.2f} %; debe ser 100 %.")
        return
    if any(abs(a - b) > 1e-9 for a, b in zip(nuevos, actual)):
        estado.fijar(clave, texto(nuevos))
        st.rerun()


cols = st.columns(2)
for i, e in enumerate(("E1", "E2", "E3", "E4")):
    with cols[i % 2]:
        st.markdown(f"<div class='franja' style='border-color:{ui.COLORES[e]}'><b>{ui.NOMBRES[e]}</b>"
                    f"{' · dinámica' if e == 'E4' else ''}</div>", unsafe_allow_html=True)
        ui.barra_composicion(p.lista(f"{e}.pesos"))
        st.caption(DESC[e])

with st.expander("Editar pesos iniciales"):
    for e in ("E1", "E2", "E3", "E4"):
        st.markdown(f"**{ui.NOMBRES[e]}**")
        editor_pesos(f"{e}.pesos", ACTIVOS, f"pesos_{e}")

st.markdown("##### Regla de rebalanceo de E4")
cada_m, banda_p = p.i("E4.cada"), p.f("E4.banda")
r = st.columns(3)
with r[0]:
    ui.tarjeta("Revisión", f"cada {cada_m} meses", "cuenta, CDT y ETF; el inmueble no se toca")
with r[1]:
    ui.tarjeta("Banda de tolerancia", f"{banda_p * 100:g} puntos", "rebalancea si un activo se aleja más")
with r[2]:
    ui.tarjeta("Referencia", "Estado de la economía", "expansión, normal o crisis")
st.write("")
st.markdown("**Objetivo por estado** (% del portafolio financiero)")
st.dataframe({"Estado": ["Expansión", "Normal", "Crisis"],
              **{a: [f"{p.lista(f'E4.obj_{L}')[i] * 100:g} %" for L in "ENC"] for i, a in enumerate(ACTIVOS[:3])}},
             hide_index=True, width="stretch")
st.caption("Las compras de ETF financiadas con CDT esperan a su vencimiento para evitar el castigo por venta anticipada.")

with st.expander("Editar regla de E4"):
    for L, nombre in (("E", "Expansión"), ("N", "Normal"), ("C", "Crisis")):
        st.markdown(f"**Objetivo en {nombre.lower()}**")
        editor_pesos(f"E4.obj_{L}", ACTIVOS[:3], f"obj_{L}")
    c1, c2 = st.columns(2)
    with c1:
        banda = st.number_input("Banda de tolerancia (puntos porcentuales)", 0.0, 50.0, p.f("E4.banda") * 100, 0.5,
                                help=ui.pie_parametro(T["E4.banda"]))
        if abs(banda / 100 - p.f("E4.banda")) > 1e-12:
            estado.fijar("E4.banda", banda / 100)
    with c2:
        cada = st.number_input("Revisión cada … meses", 1, 60, p.i("E4.cada"), 1, help=ui.pie_parametro(T["E4.cada"]))
        if cada != p.i("E4.cada"):
            estado.fijar("E4.cada", int(cada))

if st.button("Restablecer valores iniciales"):
    estado.restablecer(("E1.", "E2.", "E3.", "E4."))
    st.rerun()

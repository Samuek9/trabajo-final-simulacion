import plotly.graph_objects as go
import streamlit as st

from components import charts, estado, ui

p = estado.params()
T = p.tabla
ACTIVOS = ("Cuenta", "CDT", "ETF", "Inmueble")
COL_ACTIVOS = ("#7FA7C9", "#2F6B9A", "#C0612B", "#3C8D5A")

st.title("Estrategias")
st.write("Pesos iniciales de cada estrategia. Cada fila debe sumar 100 %. E1, E2 y E3 compran y mantienen: "
         "solo venden para cubrir liquidez y renuevan el CDT por el mismo capital. E4 empieza como E2 y una vez al año "
         "acomoda cuenta, CDT y ETF al estado de la economía; el inmueble no lo toca.")


def texto(xs: list[float]) -> str:
    return ";".join(f"{x:g}" for x in xs)


def editor_pesos(clave: str, etiquetas: tuple[str, ...], k: str) -> list[float] | None:
    actual = p.lista(clave)
    cols = st.columns(len(etiquetas))
    nuevos = []
    for c, et, v, i in zip(cols, etiquetas, actual, range(len(etiquetas))):
        with c:
            nuevos.append(st.number_input(f"{et} (%)", 0.0, 100.0, round(v * 100, 2), 1.0, key=f"{k}_{i}") / 100)
    suma = sum(nuevos)
    ref = [float(x) for x in str(T[clave].base).split(";")]
    st.caption(f"Valor de referencia: {' · '.join(f'{x * 100:g}' for x in ref)} %")
    if abs(suma - 1) > 1e-6:
        st.error(f"Suma {suma * 100:.2f} %. Ajusta para que dé 100 %.")
        return None
    if any(abs(a - b) > 1e-9 for a, b in zip(nuevos, actual)):
        estado.fijar(clave, texto(nuevos))
        st.rerun()
    return nuevos


for e in ("E1", "E2", "E3", "E4"):
    st.markdown(f"<div class='franja' style='border-color:{ui.COLORES[e]}'><b>{ui.NOMBRES[e]}</b>"
                f"{' (inicio)' if e == 'E4' else ''}</div>", unsafe_allow_html=True)
    editor_pesos(f"{e}.pesos", ACTIVOS, f"pesos_{e}")

st.markdown("### Reglas de E4")
st.write("En cada revisión, E4 compara el reparto de su portafolio financiero (lo que no es inmueble) con el objetivo del "
         "estado vigente. Si algún activo se aleja más que la banda, vende y compra hasta el objetivo. Las compras de ETF "
         "que salgan del CDT esperan a que el CDT venza para no pagar el castigo por venta anticipada.")
for L, nombre in (("E", "Expansión"), ("N", "Normal"), ("C", "Crisis")):
    st.markdown(f"**Objetivo en {nombre.lower()}** (cuenta, CDT y ETF del portafolio financiero)")
    editor_pesos(f"E4.obj_{L}", ACTIVOS[:3], f"obj_{L}")

c1, c2 = st.columns(2)
with c1:
    banda = st.number_input("Banda de tolerancia (puntos porcentuales)", 0.0, 50.0, p.f("E4.banda") * 100, 0.5)
    st.caption(ui.pie_parametro(T["E4.banda"]))
    if abs(banda / 100 - p.f("E4.banda")) > 1e-12:
        estado.fijar("E4.banda", banda / 100)
with c2:
    cada = st.number_input("Revisa cada ... meses", 1, 60, p.i("E4.cada"), 1)
    st.caption(ui.pie_parametro(T["E4.cada"]))
    if cada != p.i("E4.cada"):
        estado.fijar("E4.cada", int(cada))

if st.button("Volver a la referencia"):
    estado.restablecer(("E1.", "E2.", "E3.", "E4."))
    st.rerun()

p = estado.params()
fig = go.Figure()
nombres = [ui.NOMBRES[e] for e in ("E1", "E2", "E3", "E4")] + [f"E4 objetivo {n}" for n in ("expansión", "normal", "crisis")]
for i, (activo, color) in enumerate(zip(ACTIVOS, COL_ACTIVOS)):
    xs = [p.lista(f"{e}.pesos")[i] * 100 for e in ("E1", "E2", "E3", "E4")]
    inm = p.lista("E4.pesos")[3]
    for L in "ENC":
        obj = p.lista(f"E4.obj_{L}")
        xs.append(obj[i] * (1 - inm) * 100 if i < 3 else inm * 100)
    fig.add_trace(go.Bar(y=nombres, x=xs, name=activo, orientation="h", marker_color=color,
                         hovertemplate="%{x:.1f} %<extra>" + activo + "</extra>"))
fig.update_layout(barmode="stack", title="Reparto del capital (%)", xaxis_title="% del patrimonio", height=380,
                  yaxis=dict(autorange="reversed"), **{k: v for k, v in charts.LAYOUT.items() if k != "hovermode"})
st.plotly_chart(fig, width="stretch", config=charts.CONFIG)
st.caption("Las filas de E4 objetivo suponen que el inmueble conserva su peso inicial; en la simulación su peso cambia con la valorización.")

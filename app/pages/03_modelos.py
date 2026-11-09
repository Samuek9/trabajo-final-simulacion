import math

import numpy as np
import pandas as pd
import streamlit as st

from components import estado, ui
from investment_sim.models import inflation, markov

p = estado.params()
T = p.tabla

st.title("Modelos y economía")
st.write("Cada control parte del valor de referencia de la hoja PARAMETROS del Excel maestro. Debajo de cada uno "
         "están la fuente y el estado: **Estimado** sale de los datos, **Supuesto** es una decisión del equipo y "
         "**Pendiente** no tiene datos. Cambiar un valor solo afecta esta sesión; el Excel no se modifica.")


def numero(clave: str, etiqueta: str, escala: float = 1.0, paso: float | None = None, minimo: float | None = None,
           maximo: float | None = None, fmt: str = "%.4f") -> None:
    t = T[clave]
    valor = p.f(clave) * escala
    nuevo = st.number_input(etiqueta, value=float(valor), step=paso, min_value=minimo, max_value=maximo, format=fmt,
                            key=f"w_{clave}")
    st.caption(f"<span class='etiqueta'>{t.estado}</span> Referencia {ui.num(float(t.base) * escala, 4)} · {t.fuente}",
               unsafe_allow_html=True)
    if abs(nuevo - valor) > 1e-12:
        err = estado.probar({clave: nuevo / escala})
        if err:
            st.error(err)
        else:
            estado.fijar(clave, nuevo / escala)


def opcion(clave: str, etiqueta: str, opciones: dict[str, str]) -> None:
    t = T[clave]
    actual = p.s(clave)
    nuevo = st.radio(etiqueta, list(opciones), index=list(opciones).index(actual), format_func=opciones.get,
                     key=f"w_{clave}", horizontal=False)
    st.caption(f"<span class='etiqueta'>{t.estado}</span> Referencia: {opciones[t.base]} · {t.fuente}", unsafe_allow_html=True)
    if nuevo != actual:
        estado.fijar(clave, nuevo)


def restablecer(prefijos: tuple[str, ...], k: str) -> None:
    if st.button("Volver a la referencia", key=f"r_{k}"):
        estado.restablecer(prefijos)
        st.rerun()


with st.expander("Estados económicos: cadena de Markov (punto 8)", expanded=True):
    st.write("Probabilidad de pasar de un estado (fila) a otro (columna) en un trimestre. Estimada con 82 trimestres "
             "del PIB real (DANE, 2006 a 2026): expansión si crece más de 4 %, crisis si crece menos de 1 %.")
    P = markov.matriz(p)
    df = pd.DataFrame(P, index=markov.ESTADOS, columns=markov.ESTADOS)
    edit = st.data_editor(df, key="markov_editor", width="stretch",
                          column_config={c: st.column_config.NumberColumn(format="%.4f", min_value=0.0, max_value=1.0) for c in df})
    filas = edit.to_numpy().sum(axis=1)
    if not np.allclose(filas, 1, atol=1e-6):
        st.error("Cada fila debe sumar 1. Sumas actuales: " + ", ".join(f"{x:.4f}" for x in filas))
    elif not np.allclose(edit.to_numpy(), P):
        for i, o in enumerate("ENC"):
            for j, d in enumerate("ENC"):
                estado.fijar(f"markov.{o}_{d}", float(edit.iat[i, j]))
        st.rerun()
    pi = markov.estacionaria(P)
    dur = 1 / (1 - np.diag(P))
    st.caption("Largo plazo: " + " · ".join(f"{s} {x * 100:.0f} %" for s, x in zip(markov.ESTADOS, pi)) +
               " | Duración media: " + " · ".join(f"{s} {d:.1f} trimestres" for s, d in zip(markov.ESTADOS, dur)))
    estado_ini = st.selectbox("Estado inicial", markov.ESTADOS, index=markov.ESTADOS.index(p.s("eco.estado_inicial")))
    if estado_ini != p.s("eco.estado_inicial"):
        estado.fijar("eco.estado_inicial", estado_ini)
    restablecer(("markov.", "eco."), "markov")

with st.expander("Inflación (7.1)"):
    st.write("Inflación de cada diciembre: ln(π) = a + φ·ln(π anterior) + σ·Z. Entre diciembres, línea recta.")
    c = st.columns(4)
    with c[0]:
        nivel = inflation.nivel_largo_plazo(p.f("inf.c"), p.f("inf.phi"))
        nuevo = st.number_input("Nivel normal (%)", value=float(round(nivel, 3)), step=0.25, min_value=0.5, max_value=30.0)
        st.caption(f"<span class='etiqueta'>Estimado</span> Referencia {ui.num(inflation.nivel_largo_plazo(T['inf.c'].base, T['inf.phi'].base), 2)} % · DANE 2001-2025", unsafe_allow_html=True)
        if abs(nuevo - round(nivel, 3)) > 1e-9:
            estado.fijar("inf.c", (1 - p.f("inf.phi")) * math.log(nuevo))
    with c[1]:
        numero("inf.phi", "Memoria φ", paso=0.05, minimo=0.0, maximo=0.99)
    with c[2]:
        numero("inf.sigma", "σ de las sorpresas", paso=0.05, minimo=0.0)
    with c[3]:
        numero("inf.inicial", "Inflación al iniciar (%)", paso=0.1, minimo=0.1, fmt="%.2f")
    restablecer(("inf.",), "inf")

with st.expander("Tasas: TPM, cuenta y CDT (7.2 a 7.4)"):
    c = st.columns(3)
    with c[0]:
        opcion("junta.modelo", "Decisión de la Junta",
               {"logit": "Logit: tendencia de la inflación y brecha de Taylor", "tendencia": "Tabla por tendencia (7.2)"})
        numero("tpm.inicial", "TPM inicial (%)", paso=0.25, minimo=0.0, fmt="%.2f")
        numero("tpm.max", "TPM máxima (%)", paso=0.5, fmt="%.2f")
    with c[1]:
        numero("cta.alfa", "Cuenta: tasa = α × TPM", paso=0.01, minimo=0.0, maximo=1.0)
        numero("imp.retencion", "Retención sobre intereses del CDT", paso=0.01, minimo=0.0, maximo=0.5)
        numero("imp.retencion_cuenta", "Retención sobre intereses de la cuenta", paso=0.01, minimo=0.0, maximo=0.5)
        numero("cdt.castigo", "Castigo por venta anticipada del CDT", paso=0.005, minimo=0.0, maximo=0.1)
    with c[2]:
        opcion("cdt.modelo", "Margen del CDT", {"m5": "M5: a + b·ΔTPM + ε (residuos observados)", "fase": "Normal por fase de la TPM"})
        numero("cdt.a", "M5: constante a (puntos)", paso=0.1)
        numero("cdt.b", "M5: pendiente b", paso=0.05)
    restablecer(("tpm.", "junta.", "cta.", "cdt.", "imp."), "tasas")

with st.expander("Renta variable: ETF iCOLCAP (7.5)"):
    c = st.columns(3)
    with c[0]:
        opcion("etf.modelo", "Modelo", {"t5_estado": "t de Student con volatilidad del estado", "jf": "t asimétrica de Jones y Faddy (sin estados)"})
        numero("etf.nu", "Grados de libertad ν", paso=0.5, minimo=2.1)
    with c[1]:
        numero("etf.mu", "Rendimiento log medio mensual", paso=0.001)
        numero("etf.comision", "Comisión por operación", paso=0.001, minimo=0.0, maximo=0.05)
    with c[2]:
        numero("etf.vol_E", "Volatilidad mensual en expansión", paso=0.005, minimo=0.0)
        numero("etf.vol_N", "Volatilidad mensual en normal", paso=0.005, minimo=0.0)
        numero("etf.vol_C", "Volatilidad mensual en crisis", paso=0.005, minimo=0.0)
    restablecer(("etf.",), "etf")

with st.expander("Inmueble: valorización, canon y gastos (5.4, 7.6)"):
    c = st.columns(3)
    with c[0]:
        opcion("val.modelo", "Valorización", {"estado": "Normal según el estado (IPVNBR Medellín)", "ar1": "AR(1) del IPVU"})
        numero("val.mu_E", "Valorización log anual en expansión", paso=0.005)
        numero("val.mu_N", "Valorización log anual en normal", paso=0.005)
        numero("val.mu_C", "Valorización log anual en crisis", paso=0.005)
    with c[1]:
        numero("inm.canon_inicial", "Canon de un contrato nuevo (fracción mensual del valor)", paso=0.0005, minimo=0.0, maximo=0.01)
        opcion("inm.reajuste_modelo", "Reajuste del canon", {"dane": "Modelo DANE: a + b·IPC previo + ε, tope IPC previo", "ipc": "Todo el IPC previo"})
        numero("inm.comision_inmobiliaria", "Comisión de la inmobiliaria", paso=0.01, minimo=0.0, maximo=0.3)
    with c[2]:
        numero("inm.ph", "Administración del edificio (COP/mes, pesos de hoy)", paso=5000.0, minimo=0.0, fmt="%.0f")
        numero("inm.predial", "Predial (fracción anual del valor)", paso=0.001, minimo=0.0)
        numero("inm.mantenimiento", "Mantenimiento (fracción anual del valor)", paso=0.001, minimo=0.0)
        numero("inm.descuento_urgente", "Descuento por venta urgente", paso=0.01, minimo=0.0, maximo=0.5)
    restablecer(("val.", "inm."), "inm")

with st.expander("Vacancia (7.7)"):
    st.write("Al vencer cada contrato anual el inquilino no renueva con probabilidad p(estado); la vacancia dura una gamma con la media del estado.")
    c = st.columns(3)
    for col, (L, nombre) in zip(c, (("E", "expansión"), ("N", "normal"), ("C", "crisis"))):
        with col:
            numero(f"vac.p_{L}", f"Probabilidad de no renovar en {nombre}", paso=0.01, minimo=0.0, maximo=1.0)
            numero(f"vac.media_{L}", f"Duración media en {nombre} (meses)", paso=0.25, minimo=0.05, fmt="%.2f")
    numero("vac.forma", "Forma de la gamma", paso=0.1, minimo=0.1, fmt="%.2f")
    restablecer(("vac.",), "vac")

with st.expander("Necesidades extraordinarias de liquidez (punto 12) y reparaciones"):
    c = st.columns(3)
    for col, (L, nombre) in zip(c, (("E", "expansión"), ("N", "normal"), ("C", "crisis"))):
        with col:
            meses = 1 / p.f(f"liq.tasa_{L}") if p.f(f"liq.tasa_{L}") > 0 else 0.0
            nuevo = st.number_input(f"Una necesidad cada ... meses ({nombre})", value=float(round(meses, 2)), min_value=0.0, step=1.0,
                                    help="0 = sin necesidades", key=f"liq_{L}")
            st.caption(f"<span class='etiqueta'>Supuesto</span> Referencia {1 / T[f'liq.tasa_{L}'].base:.0f} meses · Equipo", unsafe_allow_html=True)
            if abs(nuevo - round(meses, 2)) > 1e-9:
                estado.fijar(f"liq.tasa_{L}", 1 / nuevo if nuevo > 0 else 0.0)
    c = st.columns(3)
    with c[0]:
        numero("liq.mediana", "Monto mediano (COP de hoy)", paso=1e6, minimo=0.0, fmt="%.0f")
    with c[1]:
        numero("liq.sigma", "σ del logaritmo del monto", paso=0.1, minimo=0.0)
    with c[2]:
        st.markdown("**Reparaciones** (pendiente: sin datos; para activarlas, primero el costo y luego la frecuencia)")
        numero("rep.mediana", "Costo mediano de una reparación (COP de hoy)", paso=5e5, minimo=0.0, fmt="%.0f")
        numero("rep.sigma", "σ del logaritmo del costo", paso=0.1, minimo=0.0)
        numero("rep.tasa", "Reparaciones por mes", paso=0.01, minimo=0.0)
    restablecer(("liq.", "rep."), "liq")

"""Formato y piezas visuales comunes."""
from __future__ import annotations

import streamlit as st

COLORES = {"E1": "#2F6B9A", "E2": "#3C8D5A", "E3": "#C0612B", "E4": "#7A4E9C"}
NOMBRES = {"E1": "E1 Seguridad", "E2": "E2 Moderada", "E3": "E3 Arriesgada", "E4": "E4 Dinámica"}
ESCENARIOS = {"base": "Base", "adverso": "Adverso", "estres": "Estrés"}
ACTIVOS = ("Cuenta", "CDT", "ETF", "Inmueble")
COL_ACTIVOS = ("#7FA7C9", "#2F6B9A", "#C0612B", "#3C8D5A")

CSS = """
<style>
.block-container {padding-top: 2.2rem; max-width: 1200px;}
h1 {font-size: 1.9rem !important; letter-spacing: -0.01em; margin-bottom: .2rem;}
h2, h3 {letter-spacing: -0.01em;}
h5 {text-transform: uppercase; font-size: .78rem !important; letter-spacing: .06em; opacity: .7;}
[data-testid="stMarkdownContainer"] h5 {padding-top: 1.6rem !important;}
.subtitulo {opacity: .75; margin: -.2rem 0 1.2rem 0; font-size: 1rem;}
.tarjeta {border: 1px solid rgba(128,128,128,.25); border-radius: 8px; padding: 12px 14px; height: 100%; margin-bottom: .5rem; background: rgba(128,128,128,.04);}
.tarjeta .titulo {font-size: .72rem; text-transform: uppercase; letter-spacing: .06em; opacity: .7;}
.tarjeta .valor {font-size: 1.5rem; font-weight: 650; margin-top: 2px; line-height: 1.2;}
.tarjeta .nota {font-size: .78rem; opacity: .75; margin-top: 2px;}
.barra {display: flex; height: 22px; border-radius: 4px; overflow: hidden; margin: 6px 0 4px 0; background: rgba(128,128,128,.15);}
.barra span {display: block; height: 100%;}
.leyenda {font-size: .75rem; opacity: .8; display: flex; gap: 12px; flex-wrap: wrap;}
.leyenda i {display: inline-block; width: 9px; height: 9px; border-radius: 2px; margin-right: 4px;}
.franja {border-left: 4px solid; padding-left: 10px; margin-bottom: 6px;}
[data-testid="stSidebar"] {border-right: 1px solid rgba(128,128,128,.2);}
@media (max-width: 700px) {.block-container {padding-left: 1rem; padding-right: 1rem;} .tarjeta .valor {font-size: 1.25rem;}}
</style>
"""


def estilo() -> None:
    st.markdown(CSS, unsafe_allow_html=True)


def cop(x: float, millones: bool = True) -> str:
    if x is None or x != x:
        return "—"
    if millones:
        s = f"{x / 1e6:,.0f}".replace(",", ".")
        return f"${s} M"
    return "$" + f"{x:,.0f}".replace(",", ".")


def pct(x: float, dec: int = 1) -> str:
    if x is None or x != x:
        return "—"
    return f"{x * 100:.{dec}f} %".replace(".", ",")


def num(x: float, dec: int = 2) -> str:
    if x is None or x != x:
        return "—"
    return f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def esc(texto: str) -> str:
    """Evita que Streamlit lea los signos de pesos como fórmulas."""
    return texto.replace("$", chr(92) + "$")


def tarjeta(titulo: str, valor: str, nota: str = "", color: str | None = None) -> None:
    titulo, valor, nota = (t.replace("$", "&#36;") for t in (titulo, valor, nota))
    borde = f"border-left: 4px solid {color};" if color else ""
    st.markdown(f"<div class='tarjeta' style='{borde}'><div class='titulo'>{titulo}</div>"
                f"<div class='valor'>{valor}</div><div class='nota'>{nota}</div></div>", unsafe_allow_html=True)


def pie_parametro(t) -> str:
    """Texto bajo un control: valor inicial y unidad del parámetro."""
    ref = t.base
    if isinstance(ref, float):
        ref_txt = num(ref, 4) if abs(ref) < 10 else num(ref, 0)
    else:
        ref_txt = str(ref)
    return f"Valor inicial: {ref_txt} {t.unidad}"


def subtitulo(texto: str) -> None:
    st.markdown(f"<div class='subtitulo'>{texto.replace('$', '&#36;')}</div>", unsafe_allow_html=True)


def barra_composicion(pesos, etiquetas=ACTIVOS, leyenda: bool = True) -> None:
    """Barra horizontal apilada con el reparto entre activos (pesos que suman 1)."""
    tramos = "".join(f"<span style='width:{w * 100:.2f}%;background:{c}' title='{a} {w * 100:.0f} %'></span>"
                     for w, a, c in zip(pesos, etiquetas, COL_ACTIVOS) if w > 0)
    ley = "".join(f"<span><i style='background:{c}'></i>{a} {w * 100:.0f} %</span>"
                  for w, a, c in zip(pesos, etiquetas, COL_ACTIVOS) if w > 0)
    st.markdown(f"<div class='barra'>{tramos}</div>" + (f"<div class='leyenda'>{ley}</div>" if leyenda else ""),
                unsafe_allow_html=True)

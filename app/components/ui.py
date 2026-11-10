"""Formato y piezas visuales comunes."""
from __future__ import annotations

import streamlit as st

COLORES = {"E1": "#2F6B9A", "E2": "#3C8D5A", "E3": "#C0612B", "E4": "#7A4E9C"}
NOMBRES = {"E1": "E1 Seguridad", "E2": "E2 Moderada", "E3": "E3 Arriesgada", "E4": "E4 Dinámica"}
ESCENARIOS = {"base": "Base", "adverso": "Adverso", "estres": "Estrés"}

CSS = """
<style>
.block-container {padding-top: 2rem; max-width: 1280px;}
h1, h2, h3 {letter-spacing: -0.01em;}
.tarjeta {border: 1px solid rgba(128,128,128,.25); border-radius: 10px; padding: 14px 16px; height: 100%;}
.tarjeta .titulo {font-size: .8rem; text-transform: uppercase; letter-spacing: .05em; opacity: .7;}
.tarjeta .valor {font-size: 1.6rem; font-weight: 650; margin-top: 2px;}
.tarjeta .nota {font-size: .8rem; opacity: .75; margin-top: 2px;}
.etiqueta {display: inline-block; padding: 1px 8px; border-radius: 999px; font-size: .75rem; margin-right: 4px;
           border: 1px solid rgba(128,128,128,.35);}
.franja {border-left: 4px solid; padding-left: 10px; margin-bottom: 6px;}
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


def tarjeta(titulo: str, valor: str, nota: str = "", color: str | None = None) -> None:
    borde = f"border-left: 4px solid {color};" if color else ""
    st.markdown(f"<div class='tarjeta' style='{borde}'><div class='titulo'>{titulo}</div>"
                f"<div class='valor'>{valor}</div><div class='nota'>{nota}</div></div>", unsafe_allow_html=True)


def pie_parametro(t) -> str:
    """Texto bajo un control: referencia, estado y fuente del parámetro."""
    ref = t.base
    if isinstance(ref, float):
        ref_txt = num(ref, 4) if abs(ref) < 10 else num(ref, 0)
    else:
        ref_txt = str(ref)
    return f"Valor de referencia: {ref_txt} {t.unidad}"

"""Estado de la sesión: cambios del usuario sobre los parámetros y resultados guardados."""
from __future__ import annotations

import sys
from pathlib import Path

import streamlit as st

RAIZ = Path(__file__).resolve().parents[2]
if str(RAIZ / "src") not in sys.path:
    sys.path.insert(0, str(RAIZ / "src"))

from investment_sim.config.parameters import ParametroInvalido, Params, cargar, leer_tabla  # noqa: E402


@st.cache_resource(show_spinner=False)
def tabla():
    return leer_tabla()


def iniciar() -> None:
    st.session_state.setdefault("cambios", {})
    st.session_state.setdefault("resultados", {})
    st.session_state.setdefault("escenario", "base")


def cambios() -> dict:
    iniciar()
    return st.session_state["cambios"]


def params(escenario: str | None = None) -> Params:
    """Parámetros del escenario con los cambios de la sesión. Si algún cambio es
    inválido, lo informa y usa los valores iniciales."""
    iniciar()
    esc = escenario or st.session_state["escenario"]
    try:
        return cargar(esc, cambios(), tabla=tabla())
    except ParametroInvalido as e:
        st.error(f"Hay un parámetro inválido: {e}. Se usan los valores iniciales hasta corregirlo.")
        return cargar(esc, tabla=tabla())


def referencia(escenario: str = "base") -> Params:
    return cargar(escenario, tabla=tabla())


def fijar(clave: str, valor) -> None:
    """Guarda un cambio solo si difiere del valor inicial; si vuelve a él, lo borra."""
    ref = referencia().valores[clave]
    c = cambios()
    igual = (isinstance(ref, (int, float)) and isinstance(valor, (int, float)) and abs(float(ref) - float(valor)) < 1e-12) or ref == valor
    if igual:
        c.pop(clave, None)
    else:
        c[clave] = valor
    st.session_state["resultados"] = {k: v for k, v in st.session_state["resultados"].items() if k == "_ninguno"}


def restablecer(prefijos: tuple[str, ...] | None = None) -> None:
    c = cambios()
    for k in list(c):
        if prefijos is None or k.startswith(prefijos):
            c.pop(k)


def probar(cambio: dict) -> str | None:
    """Devuelve el error si el cambio deja los parámetros inválidos."""
    try:
        cargar("base", {**cambios(), **cambio}, tabla=tabla())
        return None
    except (ParametroInvalido, KeyError, ValueError) as e:
        return str(e)

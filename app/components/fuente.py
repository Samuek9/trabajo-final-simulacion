"""Resultados de la ejecución hecha en esta sesión con los datos que ingresó el usuario."""
from __future__ import annotations

import pandas as pd
import streamlit as st

from investment_sim.metrics.summary import bandas, pareada

PARES = (("E4", "E2"), ("E3", "E2"), ("E2", "E1"), ("E3", "E1"), ("E4", "E1"), ("E4", "E3"))
VARIABLES = ("W_real", "meta", "algun_episodio", "caida_max")


def escenarios() -> list[str]:
    return list(st.session_state.get("resultados", {}))


def datos(esc: str) -> dict:
    """tabla (una fila por réplica), resumen, bandas e iliquidez por mes, comparaciones pareadas."""
    d = st.session_state["resultados"][esc]
    r = d["exp"]
    tabla = r.tabla
    band = {e: bandas(r.trayectorias[e]["W_real"]) for e in r.estrategias if "W_real" in r.trayectorias[e]}
    iliq = {e: r.trayectorias[e]["iliquido"].mean(axis=0) for e in r.estrategias if "iliquido" in r.trayectorias[e]}
    est = set(r.estrategias)
    pares = pd.DataFrame([{"escenario": esc, **pareada(tabla, a, b, v)} for a, b in PARES
                          if a in est and b in est for v in VARIABLES])
    return {"tabla": tabla, "resumen": d["resumen"], "bandas": band, "iliq": iliq, "pares": pares, "params": d["params"],
            "nota": f"{r.n:,} réplicas, semilla {r.semilla}".replace(",", ".")}

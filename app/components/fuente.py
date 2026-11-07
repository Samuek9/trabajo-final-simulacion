"""De dónde salen los resultados que se muestran: la ejecución de esta sesión o la corrida del informe."""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import streamlit as st

from investment_sim.metrics.summary import bandas, pareada

RES = Path(__file__).resolve().parents[2] / "outputs" / "resultados"
PARES = (("E4", "E2"), ("E3", "E2"), ("E2", "E1"), ("E3", "E1"), ("E4", "E1"), ("E4", "E3"))
VARIABLES = ("W_real", "meta", "algun_episodio", "caida_max")


@st.cache_data(show_spinner=False)
def _guardado(nombre: str) -> pd.DataFrame | None:
    f = RES / nombre
    return pd.read_csv(f) if f.exists() else None


def hay_guardados() -> bool:
    return (RES / "resumen.csv").exists()


def corrida() -> dict:
    f = RES / "corrida.json"
    return json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}


def guardado(nombre: str) -> pd.DataFrame | None:
    return _guardado(nombre)


def escenarios(origen: str) -> list[str]:
    if origen == "sesion":
        return list(st.session_state.get("resultados", {}))
    r = _guardado("resumen.csv")
    return list(dict.fromkeys(r.escenario)) if r is not None else []


def datos(origen: str, esc: str) -> dict:
    """tabla (una fila por réplica), resumen, bandas e iliquidez por mes, comparaciones pareadas."""
    if origen == "sesion":
        d = st.session_state["resultados"][esc]
        r = d["exp"]
        tabla = r.tabla
        band = {e: bandas(r.trayectorias[e]["W_real"]) for e in r.estrategias if "W_real" in r.trayectorias[e]}
        iliq = {e: r.trayectorias[e]["iliquido"].mean(axis=0) for e in r.estrategias if "iliquido" in r.trayectorias[e]}
        est = set(r.estrategias)
        pares = pd.DataFrame([{"escenario": esc, **pareada(tabla, a, b, v)} for a, b in PARES
                              if a in est and b in est for v in VARIABLES])
        return {"tabla": tabla, "resumen": d["resumen"], "bandas": band, "iliq": iliq, "pares": pares,
                "params": d["params"], "nota": f"Ejecución de esta sesión: {r.n:,} réplicas, semilla {r.semilla}".replace(",", ".")}
    tabla = _guardado("replicas.csv.gz")
    tabla = tabla[tabla.escenario == esc]
    res = _guardado("resumen.csv")
    b = _guardado("bandas.csv")
    b = b[b.escenario == esc]
    band = {e: g[g.serie == "W_real"].sort_values("mes").reset_index(drop=True) for e, g in b.groupby("estrategia")}
    iliq = {e: g[g.serie == "frac_iliquido"].sort_values("mes").p50.to_numpy() for e, g in b.groupby("estrategia")}
    pares = _guardado("comparaciones_pareadas.csv")
    c = corrida()
    nota = (f"Corrida del informe ({c.get('fecha', 'sin fecha')}): {c.get('replicas', '?')} réplicas, "
            f"semilla {c.get('semilla_base', '?')}, parámetros de referencia")
    return {"tabla": tabla, "resumen": res[res.escenario == esc], "bandas": band, "iliq": iliq,
            "pares": pares[pares.escenario == esc] if pares is not None else pd.DataFrame(), "params": None, "nota": nota}

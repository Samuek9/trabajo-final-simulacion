"""Análisis de sensibilidad (punto 16.3).

Cada parámetro se lleva a un valor bajo y uno alto, con las mismas semillas que la
corrida de referencia (números aleatorios comunes), y se mide el cambio en la
probabilidad de meta y en el patrimonio real mediano. Con eso se arma el tornado.
"""
from __future__ import annotations

from typing import Callable

import numpy as np
import pandas as pd

from ..config.parameters import Params
from .runner import correr


def _escala(p: Params, claves: list[str], factor: float) -> dict:
    return {k: p.f(k) * factor for k in claves}


def factores(p: Params) -> list[dict]:
    """Variaciones del tornado: nombre, fuente de incertidumbre y cambios bajo / alto."""
    phi = p.f("inf.phi")
    nivel = lambda x: (1 - phi) * np.log(x)  # noqa: E731  constante para un nivel de inflación
    vol = ["etf.vol_E", "etf.vol_N", "etf.vol_C"]
    vmu = ["val.mu_E", "val.mu_N", "val.mu_C"]
    vp = ["vac.p_E", "vac.p_N", "vac.p_C"]
    vd = ["vac.media_E", "vac.media_N", "vac.media_C"]
    lt = ["liq.tasa_E", "liq.tasa_N", "liq.tasa_C"]
    return [
        {"factor": "Nivel de la inflación (3,5 % / 6,5 %)", "fuente": "Inflación",
         "bajo": {"inf.c": nivel(3.5)}, "alto": {"inf.c": nivel(6.5)}},
        {"factor": "Volatilidad de la inflación (σ × 0,5 / × 1,5)", "fuente": "Inflación",
         "bajo": {"inf.sigma": p.f("inf.sigma") * 0.5}, "alto": {"inf.sigma": p.f("inf.sigma") * 1.5}},
        {"factor": "Rendimiento medio del ETF (0 / 0,86 % mensual)", "fuente": "Renta variable",
         "bajo": {"etf.mu": 0.0}, "alto": {"etf.mu": 2 * p.f("etf.mu")}},
        {"factor": "Volatilidad del ETF (× 0,75 / × 1,25)", "fuente": "Renta variable",
         "bajo": _escala(p, vol, 0.75), "alto": _escala(p, vol, 1.25)},
        {"factor": "TPM inicial (9 % / 14 %)", "fuente": "Tasas de interés",
         "bajo": {"tpm.inicial": 9.0, "tpm.hace_3m": 9.0}, "alto": {"tpm.inicial": 14.0, "tpm.hace_3m": 14.0}},
        {"factor": "Tasa de la cuenta (α = 0,80 / 0,95)", "fuente": "Tasas de interés",
         "bajo": {"cta.alfa": 0.80}, "alto": {"cta.alfa": 0.95}},
        {"factor": "Margen del CDT (a − 0,5 / + 0,5 puntos)", "fuente": "Tasas de interés",
         "bajo": {"cdt.a": p.f("cdt.a") - 0.5}, "alto": {"cdt.a": p.f("cdt.a") + 0.5}},
        {"factor": "Valorización (−2 / +2 puntos por año)", "fuente": "Valorización",
         "bajo": {k: p.f(k) - 0.02 for k in vmu}, "alto": {k: p.f(k) + 0.02 for k in vmu}},
        {"factor": "Canon inicial (0,50 % / 0,70 % del valor)", "fuente": "Valorización",
         "bajo": {"inm.canon_inicial": 0.005}, "alto": {"inm.canon_inicial": 0.007}},
        {"factor": "Probabilidad de no renovar (× 0,5 / × 2)", "fuente": "Vacancia",
         "bajo": _escala(p, vp, 0.5), "alto": {k: min(p.f(k) * 2, 1.0) for k in vp}},
        {"factor": "Duración de la vacancia (× 0,5 / × 2)", "fuente": "Vacancia",
         "bajo": _escala(p, vd, 0.5), "alto": _escala(p, vd, 2.0)},
        {"factor": "Frecuencia de necesidades de liquidez (× 0,5 / × 1,5)", "fuente": "Liquidez",
         "bajo": _escala(p, lt, 0.5), "alto": _escala(p, lt, 1.5)},
        {"factor": "Monto mediano de las necesidades ($10 M / $30 M)", "fuente": "Liquidez",
         "bajo": {"liq.mediana": 10e6}, "alto": {"liq.mediana": 30e6}},
        {"factor": "Probabilidad de crisis (Normal → Crisis × 0,5 / × 2)", "fuente": "Estados económicos",
         "bajo": {"markov.N_C": p.f("markov.N_C") * 0.5, "markov.N_N": p.f("markov.N_N") + p.f("markov.N_C") * 0.5},
         "alto": {"markov.N_C": p.f("markov.N_C") * 2, "markov.N_N": p.f("markov.N_N") - p.f("markov.N_C")}},
    ]


def modelos_alternativos() -> list[dict]:
    """Sensibilidad a la elección del modelo (cada proceso con su candidato alternativo)."""
    return [
        {"factor": "ETF: Jones-Faddy sin estados (en vez de t por estado)", "cambios": {"etf.modelo": "jf"}},
        {"factor": "Valorización: normal por estado con IPVNBR, vivienda nueva (en vez del AR(1) del IPVU)", "cambios": {"val.modelo": "estado"}},
        {"factor": "CDT: margen por fase (en vez de M5)", "cambios": {"cdt.modelo": "fase"}},
        {"factor": "Canon: sube todo el IPC (en vez del modelo DANE)", "cambios": {"inm.reajuste_modelo": "ipc"}},
        {"factor": "Junta: tabla por tendencia de 7.2 (en vez del logit)", "cambios": {"junta.modelo": "tendencia"}},
        {"factor": "Vacancia de 4 meses en estado normal (Lonja: 121 días en oferta antes de arrendar)",
         "cambios": {"vac.media_N": 4.0}},
        {"factor": "Sin necesidades extraordinarias de liquidez (no tienen fuente)",
         "cambios": {"liq.tasa_E": 0.0, "liq.tasa_N": 0.0, "liq.tasa_C": 0.0}},
    ]


def _medir(p: Params, estrategias, n, semilla, procesos=1) -> pd.DataFrame:
    r = correr(p, estrategias, n=n, semilla=semilla, guardar=(), procesos=procesos)
    g = r.tabla.groupby("estrategia")
    return pd.DataFrame({"p_meta": g.meta.mean(), "W_real_p50": g.W_real.median(),
                         "p_iliquidez": g.algun_episodio.mean(), "rent_real": g.rent_real.mean()})


def tornado(p: Params, estrategias=("E1", "E2", "E3", "E4"), n: int = 400, semilla: int | None = None,
            progreso: Callable[[float], None] | None = None, procesos: int = 1) -> pd.DataFrame:
    ref = _medir(p, estrategias, n, semilla, procesos)
    filas = []
    lista = factores(p)
    for i, f in enumerate(lista):
        bajo = _medir(p.con_dict(f["bajo"]), estrategias, n, semilla, procesos)
        alto = _medir(p.con_dict(f["alto"]), estrategias, n, semilla, procesos)
        for e in estrategias:
            filas.append({"factor": f["factor"], "fuente": f["fuente"], "estrategia": e,
                          "ref_p_meta": ref.loc[e, "p_meta"], "bajo_p_meta": bajo.loc[e, "p_meta"],
                          "alto_p_meta": alto.loc[e, "p_meta"],
                          "ref_W": ref.loc[e, "W_real_p50"], "bajo_W": bajo.loc[e, "W_real_p50"],
                          "alto_W": alto.loc[e, "W_real_p50"],
                          "rango_p_meta": abs(alto.loc[e, "p_meta"] - bajo.loc[e, "p_meta"]),
                          "rango_W": abs(alto.loc[e, "W_real_p50"] - bajo.loc[e, "W_real_p50"])})
        if progreso:
            progreso((i + 1) / len(lista))
    return pd.DataFrame(filas)


def sensibilidad_modelos(p: Params, estrategias=("E1", "E2", "E3", "E4"), n: int = 400,
                         semilla: int | None = None, procesos: int = 1) -> pd.DataFrame:
    ref = _medir(p, estrategias, n, semilla, procesos)
    filas = []
    for m in modelos_alternativos():
        alt = _medir(p.con_dict(m["cambios"]), estrategias, n, semilla, procesos)
        for e in estrategias:
            filas.append({"modelo": m["factor"], "estrategia": e, "ref_p_meta": ref.loc[e, "p_meta"],
                          "alt_p_meta": alt.loc[e, "p_meta"], "ref_W": ref.loc[e, "W_real_p50"],
                          "alt_W": alt.loc[e, "W_real_p50"]})
    return pd.DataFrame(filas)

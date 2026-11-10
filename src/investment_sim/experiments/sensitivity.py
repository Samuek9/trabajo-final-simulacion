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
    lp = np.exp(p.f("inf.c") / (1 - phi))  # nivel de largo plazo actual de la inflación (%)
    nivel = lambda x: (1 - phi) * np.log(x)  # noqa: E731  constante para un nivel de inflación
    vol = ["etf.vol_E", "etf.vol_N", "etf.vol_C"]
    mu_v, phi_v, sig_v = p.lista("val.ar1")
    ar1 = lambda mu: f"{mu};{phi_v};{sig_v}"  # noqa: E731  AR(1) del IPVU con otra media
    vp = ["vac.p_E", "vac.p_N", "vac.p_C"]
    vd = ["vac.media_E", "vac.media_N", "vac.media_C"]
    lt = ["liq.tasa_E", "liq.tasa_N", "liq.tasa_C"]
    return [
        {"factor": "Nivel de la inflación (3,5 % / 6,5 %)", "fuente": "Inflación",
         "bajo": {"inf.c": nivel(3.5)}, "alto": {"inf.c": nivel(6.5)}},
        {"factor": "Memoria de la inflación (φ = 0 / 0,6, mismo nivel)", "fuente": "Inflación",
         "bajo": {"inf.phi": 0.0, "inf.c": np.log(lp)}, "alto": {"inf.phi": 0.6, "inf.c": 0.4 * np.log(lp)}},
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
         "bajo": {"val.ar1": ar1(mu_v - 0.02)}, "alto": {"val.ar1": ar1(mu_v + 0.02)}},
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
        {"factor": "Liquidez mínima en la cuenta (3 % / 10 %)", "fuente": "Inversionista",
         "bajo": {"inv.liquidez_min": 0.03}, "alto": {"inv.liquidez_min": 0.10}},
        {"factor": "Descuento por venta urgente (5 % / 20 %)", "fuente": "Inmueble",
         "bajo": {"inm.descuento_urgente": 0.05}, "alto": {"inm.descuento_urgente": 0.20}},
        {"factor": "Castigo del CDT en la venta anticipada (0 / 2 puntos)", "fuente": "Tasas de interés",
         "bajo": {"cdt.castigo": 0.0}, "alto": {"cdt.castigo": 0.02}},
        {"factor": "Mantenimiento del inmueble (0,2 % / 0,5 % anual)", "fuente": "Inmueble",
         "bajo": {"inm.mantenimiento": 0.002}, "alto": {"inm.mantenimiento": 0.005}},
        {"factor": "Administración del edificio ($165.000 / $245.000)", "fuente": "Inmueble",
         "bajo": {"inm.ph": 165_000}, "alto": {"inm.ph": 245_000}},
    ]


def interaccion(p: Params, estrategias=("E1", "E2", "E3", "E4"), n: int = 400, semilla: int | None = None,
                procesos: int = 1) -> pd.DataFrame:
    """Diseño 2 × 2: nivel de la inflación (3,5 % / 6,5 %) × frecuencia de necesidades (× 0,5 / × 1,5).

    La interacción es (alta-alta − alta-baja) − (baja-alta − baja-baja) sobre P(meta): si es
    cercana a cero, los efectos se suman y el tornado de un factor a la vez basta.
    """
    phi = p.f("inf.phi")
    lt = ["liq.tasa_E", "liq.tasa_N", "liq.tasa_C"]
    res = {}
    for ni, pi in (("baja", 3.5), ("alta", 6.5)):
        for nl, k in (("baja", 0.5), ("alta", 1.5)):
            cambios = {"inf.c": (1 - phi) * np.log(pi), **_escala(p, lt, k)}
            res[(ni, nl)] = _medir(p.con_dict(cambios), estrategias, n, semilla, procesos)
    filas = []
    for e in estrategias:
        bb, ba = res[("baja", "baja")].loc[e, "p_meta"], res[("baja", "alta")].loc[e, "p_meta"]
        ab, aa = res[("alta", "baja")].loc[e, "p_meta"], res[("alta", "alta")].loc[e, "p_meta"]
        filas.append({"estrategia": e, "inf_baja_liq_baja": bb, "inf_baja_liq_alta": ba,
                      "inf_alta_liq_baja": ab, "inf_alta_liq_alta": aa,
                      "efecto_inflacion": (ab + aa - bb - ba) / 2, "efecto_liquidez": (ba + aa - bb - ab) / 2,
                      "interaccion": (aa - ab) - (ba - bb)})
    return pd.DataFrame(filas)


def modelos_alternativos() -> list[dict]:
    """Sensibilidad a la elección del modelo (cada proceso con su candidato alternativo)."""
    return [
        {"factor": "ETF: Jones-Faddy sin estados (en vez de t por estado)", "cambios": {"etf.modelo": "jf"}},
        {"factor": "Valorización: normal por estado con IPVNBR, vivienda nueva (en vez del AR(1) del IPVU)", "cambios": {"val.modelo": "estado"}},
        {"factor": "CDT: margen por fase (en vez de M5)", "cambios": {"cdt.modelo": "fase"}},
        {"factor": "Canon: sube todo el IPC (en vez del modelo DANE)", "cambios": {"inm.reajuste_modelo": "ipc"}},
        {"factor": "Junta: tabla por tendencia de 7.2 (en vez del logit)", "cambios": {"junta.modelo": "tendencia"}},
        {"factor": "Junta: logit estimado solo con 2017-2026 (78 reuniones con el calendario actual)",
         "cambios": {"junta.logit": "1.908741;-0.264993;-1.349915;2.167739",
                     "junta.subidas_logit": "0.25;0.5;0.75;1;1.5", "junta.p_subidas_logit": "0.222222;0.111111;0.111111;0.444444;0.111111",
                     "junta.bajadas_logit": "0.25;0.5;0.75", "junta.p_bajadas_logit": "0.458333;0.458333;0.083333"}},
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

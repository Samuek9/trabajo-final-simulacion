"""Rendimiento mensual del ETF ICOLCAP (7.5 y punto 8)."""
from __future__ import annotations

import math

import numpy as np
from scipy.stats import jf_skew_t
from scipy.stats import t as t_dist

_CLAVE_VOL = {"Expansión": "etf.vol_E", "Normal": "etf.vol_N", "Crisis": "etf.vol_C"}


def choques(u: np.ndarray, p) -> np.ndarray:
    """Convierte uniformes en choques del modelo elegido (inversión de la acumulada).

    t5_estado: T ~ t(ν) estándar; el estado se aplica después, mes a mes.
    jf:        g completo de la t asimétrica de Jones y Faddy (no depende del estado).
    """
    u = np.clip(u, 1e-12, 1 - 1e-12)
    if p.s("etf.modelo") == "jf":
        a, b, loc, esc = p.lista("etf.jf")
        return jf_skew_t.ppf(u, a, b, loc=loc, scale=esc)
    return t_dist.ppf(u, p.f("etf.nu"))


def rendimiento_log(estado: str, choque: float, p) -> float:
    """g = ln(1 + R) del mes.

    t5_estado: g = μ + s(S)·T con s = vol(S)·√((ν − 2)/ν): la desviación de g es la
               volatilidad del estado.
    jf:        el choque ya es g.
    """
    if p.s("etf.modelo") == "jf":
        return float(choque)
    nu = p.f("etf.nu")
    s = p.f(_CLAVE_VOL[estado]) * math.sqrt((nu - 2) / nu)
    return p.f("etf.mu") + s * float(choque)

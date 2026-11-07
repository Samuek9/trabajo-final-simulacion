"""Inflación anual de diciembre: lognormal con memoria (7.1)."""
from __future__ import annotations

import math


def siguiente_diciembre(pi_anterior: float, z: float, c: float, phi: float, sigma: float) -> float:
    """π_dic = exp(c + φ·ln(π anterior) + σ·Z), en % anual."""
    return math.exp(c + phi * math.log(max(pi_anterior, 0.05)) + sigma * z)


def nivel_largo_plazo(c: float, phi: float) -> float:
    return math.exp(c / (1 - phi))


def interpolar(pi_ini: float, pi_fin: float, t_ini: float, t: float, plazo: float = 12.0) -> float:
    """Trayectoria mensual: línea recta entre diciembres."""
    x = min(max((t - t_ini) / plazo, 0.0), 1.0)
    return pi_ini + (pi_fin - pi_ini) * x

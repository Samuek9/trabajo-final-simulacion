"""Necesidad extraordinaria de liquidez (punto 12) y reparaciones (evento 10)."""
from __future__ import annotations

import math

_L = {"Expansión": "E", "Normal": "N", "Crisis": "C"}


def tasa(estado: str, p) -> float:
    """Eventos por mes del proceso de Poisson."""
    return p.f(f"liq.tasa_{_L[estado]}")


def espera(lam: float, u: float) -> float:
    """Tiempo hasta el siguiente evento: Exponencial(λ). Infinito si λ = 0."""
    if lam <= 0:
        return math.inf
    return -math.log(max(1 - u, 1e-300)) / lam


def monto(z: float, indice_precios: float, mult: float, p) -> float:
    """Lognormal con mediana en pesos de hoy, pasada a pesos del momento."""
    return mult * p.f("liq.mediana") * indice_precios * math.exp(p.f("liq.sigma") * z)


def costo_reparacion(z: float, indice_precios: float, p) -> float:
    return p.f("rep.mediana") * indice_precios * math.exp(p.f("rep.sigma") * z)

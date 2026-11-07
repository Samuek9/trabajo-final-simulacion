"""No renovación del contrato y duración de la vacancia (7.7)."""
from __future__ import annotations

from scipy.stats import gamma

_L = {"Expansión": "E", "Normal": "N", "Crisis": "C"}


def no_renueva(estado: str, u: float, p) -> bool:
    return u < p.f(f"vac.p_{_L[estado]}")


def duracion(estado: str, u: float, p) -> float:
    """Gamma con forma fija y media según el estado, en meses."""
    k = p.f("vac.forma")
    media = p.f(f"vac.media_{_L[estado]}")
    return float(gamma.ppf(min(max(u, 1e-12), 1 - 1e-12), k, scale=media / k))

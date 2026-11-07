"""Cadena de Markov de tres estados económicos (punto 8)."""
from __future__ import annotations

import numpy as np

ESTADOS = ("Expansión", "Normal", "Crisis")
LETRA = {"Expansión": "E", "Normal": "N", "Crisis": "C"}


def matriz(p) -> np.ndarray:
    return np.array([[p.f(f"markov.{o}_{d}") for d in "ENC"] for o in "ENC"])


def siguiente(estado: str, u: float, P: np.ndarray) -> str:
    """Inversión de la acumulada de la fila del estado actual."""
    acumulada = np.cumsum(P[ESTADOS.index(estado)])
    k = int(np.searchsorted(acumulada, u, side="right"))
    return ESTADOS[min(k, 2)]


def estacionaria(P: np.ndarray) -> np.ndarray:
    w, v = np.linalg.eig(P.T)
    x = np.real(v[:, np.argmin(np.abs(w - 1))])
    return x / x.sum()

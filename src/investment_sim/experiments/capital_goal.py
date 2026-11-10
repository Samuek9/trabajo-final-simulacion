"""Experimento capital-meta.

Capital inicial × meta × estrategia → probabilidad de cumplimiento. Cada capital se
simula; las metas se evalúan sobre el mismo patrimonio real final, así que cambiar la
meta no exige volver a simular.
"""
from __future__ import annotations

from typing import Callable, Iterable

import numpy as np
import pandas as pd

from ..config.parameters import Params
from ..metrics.summary import ic_proporcion
from .runner import correr


FACTORES_CAPITAL = (4 / 7, 5 / 7, 6 / 7, 1.0, 8 / 7, 9 / 7, 10 / 7)
FACTORES_META = (0.7, 0.8, 1.0, 1.2, 1.5)


def malla(p: Params) -> tuple[list[float], list[float]]:
    """Capitales y metas por defecto: múltiplos del capital y la meta configurados (redondeados al millón)."""
    capital, meta = p.f("inv.capital"), p.f("inv.meta_real")
    return ([round(capital * k / 1e6) * 1e6 for k in FACTORES_CAPITAL],
            [round(meta * k / 1e6) * 1e6 for k in FACTORES_META])


def matriz(p: Params, capitales: Iterable[float], metas: Iterable[float],
           estrategias: Iterable[str] = ("E1", "E2", "E3", "E4"), n: int = 500,
           semilla: int | None = None, progreso: Callable[[float], None] | None = None,
           procesos: int = 1) -> pd.DataFrame:
    capitales, metas, estrategias = list(capitales), list(metas), list(estrategias)
    filas = []
    for i, c in enumerate(capitales):
        pc = p.con(**{"inv.capital": float(c)})
        r = correr(pc, estrategias, n=n, semilla=semilla, guardar=(), procesos=procesos)
        for e, g in r.tabla.groupby("estrategia"):
            w = g["W_real"].to_numpy(float)
            for m in metas:
                prob, lo, hi = ic_proporcion(float((w >= m).sum()), len(w))
                filas.append({"escenario": p.escenario, "capital": c, "meta": m, "estrategia": e,
                              "p_meta": prob, "lo": lo, "hi": hi, "n": len(w)})
        if progreso:
            progreso((i + 1) / len(capitales))
    return pd.DataFrame(filas)


def capital_necesario(tabla: pd.DataFrame, meta: float, estrategia: str, prob_objetivo: float = 0.8) -> float:
    """Interpolación lineal del capital que da la probabilidad objetivo (NaN si no se alcanza)."""
    t = tabla[(tabla.meta == meta) & (tabla.estrategia == estrategia)].sort_values("capital")
    x, y = t.capital.to_numpy(float), t.p_meta.to_numpy(float)
    for k in range(1, len(x)):
        if y[k - 1] < prob_objetivo <= y[k]:
            return float(x[k - 1] + (prob_objetivo - y[k - 1]) * (x[k] - x[k - 1]) / (y[k] - y[k - 1]))
    return float(x[0]) if len(y) and y[0] >= prob_objetivo else float("nan")

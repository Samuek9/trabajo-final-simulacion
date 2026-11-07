"""Réplicas de un experimento (punto 14).

La réplica r usa la semilla base + r para las cuatro estrategias: así las cuatro
enfrentan la misma economía (números aleatorios comunes) y las réplicas son
independientes entre sí.
"""
from __future__ import annotations

import time
from dataclasses import dataclass, field
from typing import Callable, Iterable

import numpy as np
import pandas as pd

from ..config.parameters import Params
from ..data.loader import residuos_cdt
from ..engine.simulation import ResultadoReplica, Simulacion
from ..strategies.rules import Estrategia

SERIES = ("W_real", "W", "cuenta", "etf", "cdt", "inmueble", "iliquido", "estado", "tpm", "inflacion", "ocupado")


@dataclass
class ResultadoExperimento:
    escenario: str
    semilla: int
    n: int
    estrategias: list[str]
    tabla: pd.DataFrame                       # una fila por réplica y estrategia
    trayectorias: dict[str, dict[str, np.ndarray]] = field(default_factory=dict)
    segundos: float = 0.0
    cambios: dict = field(default_factory=dict)


def residuos(p: Params) -> np.ndarray:
    """Residuos del modelo M5 con a y b estimados (valores de referencia de la hoja).

    Si se calcularan con el a y el b del usuario, un cambio en a se cancelaría con
    el corrimiento de los residuos y el margen simulado no se movería."""
    t = p.tabla
    return residuos_cdt(float(t["cdt.a"].base), float(t["cdt.b"].base))


def una_replica(p: Params, estrategia: str, semilla: int, traza: bool = False) -> ResultadoReplica:
    return Simulacion(p, Estrategia(estrategia), semilla, residuos(p), p.fuera_de_crisis(), traza=traza).correr()


def _bloque(p: Params, estrategias: list[str], replicas: list[int], semilla: int, guardar: list[str]):
    res = residuos(p)
    p_fuera = p.fuera_de_crisis()
    filas, series = [], []
    for r in replicas:
        s = semilla + r
        for e in estrategias:
            out = Simulacion(p, Estrategia(e), s, res, p_fuera).correr()
            filas.append({"replica": r, "semilla": s, "estrategia": e, **out.indicadores})
            series.append((r, e, {k: out.trayectoria[k] for k in guardar}))
    return filas, series


def correr(p: Params, estrategias: Iterable[str] = ("E1", "E2", "E3", "E4"), n: int | None = None,
           semilla: int | None = None, guardar: Iterable[str] = SERIES,
           progreso: Callable[[float], None] | None = None, procesos: int = 1) -> ResultadoExperimento:
    """Corre n réplicas de cada estrategia. Con procesos > 1 reparte las réplicas en
    bloques; como cada réplica tiene su semilla, el resultado es el mismo."""
    estrategias = list(estrategias)
    n = int(n if n is not None else p.i("exp.replicas"))
    semilla = int(semilla if semilla is not None else p.i("exp.semilla"))
    guardar = list(guardar)
    H = p.i("inv.horizonte")
    tray = {e: {k: np.zeros((n, H + 1), dtype=np.int8 if k in ("estado", "ocupado") else (bool if k == "iliquido" else float))
                for k in guardar} for e in estrategias}
    filas = []
    t0 = time.perf_counter()

    def recoger(f, series):
        filas.extend(f)
        for r, e, d in series:
            for k, v in d.items():
                tray[e][k][r] = v

    if procesos > 1 and n >= 2 * procesos:
        from concurrent.futures import ProcessPoolExecutor
        bloques = [list(b) for b in np.array_split(np.arange(n), procesos * 4) if len(b)]
        hechos = 0
        with ProcessPoolExecutor(max_workers=procesos) as ex:
            futuros = [ex.submit(_bloque, p, estrategias, [int(x) for x in b], semilla, guardar) for b in bloques]
            for fu in futuros:
                recoger(*fu.result())
                hechos += 1
                if progreso is not None:
                    progreso(hechos / len(futuros))
    else:
        paso = max(n // 50, 1)
        for inicio in range(0, n, paso):
            recoger(*_bloque(p, estrategias, list(range(inicio, min(inicio + paso, n))), semilla, guardar))
            if progreso is not None:
                progreso(min(inicio + paso, n) / n)
    tabla = pd.DataFrame(filas).sort_values(["replica", "estrategia"]).reset_index(drop=True)
    tabla.insert(0, "escenario", p.escenario)
    return ResultadoExperimento(p.escenario, semilla, n, estrategias, tabla, tray,
                                time.perf_counter() - t0, dict(p.cambios))

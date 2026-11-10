"""Resumen estadístico de las réplicas (puntos 14 y 15).

Medias con intervalo t; proporciones con intervalo de Wilson; percentiles y CVaR con
bootstrap de las réplicas; diferencias entre estrategias pareadas réplica a réplica
(números aleatorios comunes).
"""
from __future__ import annotations

import numpy as np
import pandas as pd
from scipy import stats

PERCENTILES = (5, 25, 50, 75, 95)


def ic_media(x: np.ndarray, nivel: float = 0.95) -> tuple[float, float, float]:
    x = np.asarray(x, dtype=float)
    x = x[np.isfinite(x)]
    n = len(x)
    if n == 0:
        return (np.nan, np.nan, np.nan)
    m = x.mean()
    if n < 2:
        return (m, np.nan, np.nan)
    h = stats.t.ppf(0.5 + nivel / 2, n - 1) * x.std(ddof=1) / np.sqrt(n)
    return (m, m - h, m + h)


def ic_proporcion(exitos: float, n: int, nivel: float = 0.95) -> tuple[float, float, float]:
    """Intervalo de Wilson: no se sale de [0, 1] cuando p está cerca de 0 o de 1."""
    if n == 0:
        return (np.nan, np.nan, np.nan)
    z = stats.norm.ppf(0.5 + nivel / 2)
    p = exitos / n
    centro = (p + z * z / (2 * n)) / (1 + z * z / n)
    h = z * np.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / (1 + z * z / n)
    return (p, max(centro - h, 0.0), min(centro + h, 1.0))


def cvar(x: np.ndarray, alfa: float = 0.05) -> float:
    """Promedio del peor alfa de los resultados (cola izquierda)."""
    x = np.sort(np.asarray(x, dtype=float))
    k = max(int(np.ceil(alfa * len(x))), 1)
    return float(x[:k].mean())


def bootstrap(x: np.ndarray, f, b: int = 1000, semilla: int = 12345, nivel: float = 0.95) -> tuple[float, float]:
    rng = np.random.default_rng(semilla)
    x = np.asarray(x, dtype=float)
    idx = rng.integers(0, len(x), size=(b, len(x)))
    muestras = np.apply_along_axis(f, 1, x[idx])
    a = (1 - nivel) / 2
    return float(np.quantile(muestras, a)), float(np.quantile(muestras, 1 - a))


def _percentiles_boot(x: np.ndarray, b: int = 1000, semilla: int = 12345) -> dict[int, tuple[float, float, float]]:
    rng = np.random.default_rng(semilla)
    idx = rng.integers(0, len(x), size=(b, len(x)))
    boot = np.percentile(x[idx], PERCENTILES, axis=1)
    est = np.percentile(x, PERCENTILES)
    return {q: (float(est[i]), float(np.quantile(boot[i], 0.025)), float(np.quantile(boot[i], 0.975)))
            for i, q in enumerate(PERCENTILES)}


def resumen_estrategia(t: pd.DataFrame, b: int = 1000) -> dict[str, float]:
    n = len(t)
    o: dict[str, float] = {"n": n}
    for col in ("W_nominal", "W_real"):
        x = t[col].to_numpy(float)
        m, lo, hi = ic_media(x)
        o[f"{col}_media"], o[f"{col}_media_lo"], o[f"{col}_media_hi"] = m, lo, hi
        o[f"{col}_sd"] = float(x.std(ddof=1)) if n > 1 else np.nan
        for q, (v, a, c) in _percentiles_boot(x, b).items():
            o[f"{col}_p{q}"], o[f"{col}_p{q}_lo"], o[f"{col}_p{q}_hi"] = v, a, c
    x = t["W_real"].to_numpy(float)
    o["cvar5_W_real"] = cvar(x)
    o["cvar5_W_real_lo"], o["cvar5_W_real_hi"] = bootstrap(x, cvar, b)
    for col in ("rent_real", "rent_nominal", "caida_max", "meses_iliquidez", "episodios_iliquidez",
                "liquido_final_real", "participacion_arriendo", "rent_inmueble", "tiempo_vacante",
                "perdida_venta_urgente", "costos_totales", "n_rebalanceos", "n_ventas_inmueble",
                "n_necesidades", "n_coberturas", "acum_necesidades", "deuda_final"):
        if col in t:
            m, lo, hi = ic_media(t[col].to_numpy(float))
            o[f"{col}_media"], o[f"{col}_lo"], o[f"{col}_hi"] = m, lo, hi
    o["caida_max_p50"] = float(t["caida_max"].median())
    o["caida_max_p95"] = float(t["caida_max"].quantile(0.95))
    o["rent_real_sd"] = float(t["rent_real"].std(ddof=1)) if n > 1 else np.nan
    for col, nombre in (("meta", "p_meta"), ("perdida_real", "p_perdida_real"),
                        ("algun_episodio", "p_iliquidez"), ("insolvente", "p_insolvencia"),
                        ("n_ventas_inmueble", "p_venta_urgente")):
        exitos = float((t[col] > 0).sum())
        p, lo, hi = ic_proporcion(exitos, n)
        o[nombre], o[f"{nombre}_lo"], o[f"{nombre}_hi"] = p, lo, hi
    tm = t["t_meta"].dropna().to_numpy(float) / 12
    o["t_meta_media_anios"] = float(tm.mean()) if len(tm) else np.nan
    o["t_meta_p50_anios"] = float(np.median(tm)) if len(tm) else np.nan
    o["t_meta_lo"], o["t_meta_hi"] = ic_media(tm)[1:] if len(tm) > 1 else (np.nan, np.nan)
    d = t["duracion_media_iliquidez"].dropna()
    o["duracion_iliquidez_media"] = float(d.mean()) if len(d) else 0.0
    return o


def resumen(tabla: pd.DataFrame, b: int = 1000) -> pd.DataFrame:
    filas = []
    for (esc, e), g in tabla.groupby(["escenario", "estrategia"], sort=False):
        filas.append({"escenario": esc, "estrategia": e, **resumen_estrategia(g, b)})
    return pd.DataFrame(filas)


def pareada(tabla: pd.DataFrame, a: str, b: str, col: str = "W_real") -> dict[str, float]:
    """Diferencia a − b réplica por réplica. Gracias a los números aleatorios comunes
    la varianza de la diferencia es Var(a) + Var(b) − 2 Cov(a, b)."""
    x = tabla[tabla.estrategia == a].set_index("replica")[col]
    y = tabla[tabla.estrategia == b].set_index("replica")[col]
    x, y = x.align(y, join="inner")
    d = (x - y).to_numpy(float)
    m, lo, hi = ic_media(d)
    var_ind = x.var(ddof=1) + y.var(ddof=1)
    var_par = np.var(d, ddof=1)
    sd = np.std(d, ddof=1)
    p_valor = float(2 * stats.t.sf(abs(m) / (sd / np.sqrt(len(d))), len(d) - 1)) if sd > 0 else (0.0 if m != 0 else 1.0)
    return {"a": a, "b": b, "variable": col, "n": len(d), "diferencia": m, "lo": lo, "hi": hi, "p_valor": p_valor,
            "significativa": bool(lo > 0 or hi < 0),
            "reduccion_varianza": float(1 - var_par / var_ind) if var_ind > 0 else np.nan}


def replicas_necesarias(x: np.ndarray, error_relativo: float = 0.02, nivel: float = 0.95) -> int:
    """n = (z·s / (e·media))², la verificación del punto 14 para el patrimonio real."""
    z = stats.norm.ppf(0.5 + nivel / 2)
    x = np.asarray(x, float)
    return int(np.ceil((z * x.std(ddof=1) / (error_relativo * abs(x.mean()))) ** 2))


def bandas(trayectorias: np.ndarray, qs=(5, 25, 50, 75, 95)) -> pd.DataFrame:
    """Percentiles por mes de una matriz réplicas × meses."""
    return pd.DataFrame(np.percentile(trayectorias, qs, axis=0).T, columns=[f"p{q}" for q in qs])

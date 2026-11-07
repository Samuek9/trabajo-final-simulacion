"""Lectura de las series históricas que el motor o la validación necesitan."""
from __future__ import annotations

from functools import lru_cache
from pathlib import Path

import numpy as np
import pandas as pd

RAIZ = Path(__file__).resolve().parents[3]
CRUDOS = RAIZ / "data" / "raw"


@lru_cache(maxsize=4)
def residuos_cdt(a: float, b: float, ruta: str | None = None) -> np.ndarray:
    """Residuos del modelo M5: margen − (a + b·ΔTPM de 13 semanas), 2008-2026.

    Se sortean como ε al abrir cada CDT (bootstrap de los residuos observados).
    """
    df = pd.read_csv(ruta or CRUDOS / "cdt_semanal_2008.csv")
    r = (df["s"] - (a + b * df["d13"])).to_numpy(dtype=float)
    return np.sort(r[np.isfinite(r)])


def serie(nombre: str) -> pd.DataFrame:
    """Series de data/raw para la validación (inflación, IPVU, ETF, TPM, PIB)."""
    archivos = {
        "inflacion_anual": "inflacion_anual.csv",
        "inflacion_mensual": "inflacion_mensual.csv",
        "ipvu": "ipvu_anual.csv",
        "etf": "etf_estado.csv",
        "pib": "pib_estados.csv",
        "tpm": "banrep_tpm_diaria.csv",
        "cdt": "cdt_semanal_2008.csv",
    }
    return pd.read_csv(CRUDOS / archivos[nombre])

"""Decisión de la Junta: logit ordenado con tendencia de la inflación y brecha de Taylor.

Datos: TPM diaria del Banco de la República (fin de mes) e inflación anual mensual del
DANE, 2001-2026. Una observación por cada mes del calendario actual de la Junta
(enero, marzo, abril, junio, julio, septiembre, octubre y diciembre) en todos los años,
porque el simulador decide con ese calendario. Antes de 2017 la Junta se reunía cada mes;
la decisión de esos años es el cambio acumulado desde el mes de calendario anterior.
Estimar con las reuniones mensuales y simular 8 por año hacía que la TPM se moviera
menos de lo que se movió (la tasa real simulada quedaba en 4,6 puntos). Decisión: baja,
mantiene o sube.

  latente = b_tend·(π(t−1) − π(t−4)) + b_taylor·[TPM(t−1) − (r* + π(t−1) + 0,5·(π(t−1) − meta))]
  P(baja) = logística(k1 − latente); P(sube) = 1 − logística(k2 − latente)

Escribe data/metadata/junta_logit.json y la tabla de decisiones para la hoja TPM_Logit.
Uso: python tools/estimar_junta.py
"""
from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd
from scipy.stats import chi2
from statsmodels.miscmodels.ordinal_model import OrderedModel

RAIZ = Path(__file__).resolve().parents[1]
CRUDOS = RAIZ / "data" / "raw"
R_NEUTRAL = 1.0
MESES_JUNTA = [1, 3, 4, 6, 7, 9, 10, 12]
METAS = {2001: 8, 2002: 6, 2003: 5.5, 2004: 5.5, 2005: 5, 2006: 4.5, 2007: 4, 2008: 4, 2009: 5}


def datos() -> pd.DataFrame:
    tpm = pd.read_csv(CRUDOS / "banrep_tpm_diaria.csv", parse_dates=["Date"]).set_index("Date")["value"].resample("ME").last()
    inf = pd.read_csv(CRUDOS / "inflacion_mensual.csv", parse_dates=["fecha"]).set_index("fecha")["inf"]
    inf.index = inf.index.to_period("M").to_timestamp("M")
    d = pd.DataFrame({"tpm": tpm, "inf": inf}).dropna()
    d["meta"] = [METAS.get(y, 3.0) for y in d.index.year]
    d["dpi3"] = d.inf.shift(1) - d.inf.shift(4)
    pi = d.inf.shift(1)
    d["brecha_taylor"] = d.tpm.shift(1) - (R_NEUTRAL + pi + 0.5 * (pi - d.meta.shift(1)))
    d = d.dropna()
    d = d[d.index.month.isin(MESES_JUNTA)].copy()
    d["cambio"] = d.tpm.diff().round(4)
    d["decision"] = np.sign(d["cambio"])
    return d.dropna()


def estimar() -> dict:
    d = datos()
    y = pd.Categorical(d.decision, categories=[-1, 0, 1], ordered=True)
    m0 = OrderedModel(y, d[["dpi3"]], distr="logit").fit(method="bfgs", disp=False)
    m1 = OrderedModel(y, d[["dpi3", "brecha_taylor"]], distr="logit").fit(method="bfgs", disp=False)
    # statsmodels guarda el segundo umbral como log del incremento.
    k1 = float(m1.params.iloc[2])
    k2 = k1 + float(np.exp(m1.params.iloc[3]))
    lr = 2 * (m1.llf - m0.llf)
    subidas = d.cambio[d.cambio > 0].value_counts(normalize=True).sort_index()
    bajadas = (-d.cambio[d.cambio < 0]).value_counts(normalize=True).sort_index()
    real = d.tpm - d.inf
    return {
        "n": int(len(d)), "desde": str(d.index[0].date()), "hasta": str(d.index[-1].date()),
        "decisiones": {str(int(k)): int(v) for k, v in d.decision.value_counts().items()},
        "b_tend": float(m1.params["dpi3"]), "b_taylor": float(m1.params["brecha_taylor"]),
        "k1": k1, "k2": k2,
        "p_tend": float(m1.pvalues["dpi3"]), "p_taylor": float(m1.pvalues["brecha_taylor"]),
        "aic_solo_tendencia": float(m0.aic), "aic_con_taylor": float(m1.aic),
        "lr": float(lr), "lr_p": float(chi2.sf(lr, 1)), "r_neutral": R_NEUTRAL,
        "subidas": {f"{k:g}": float(v) for k, v in subidas.items()},
        "bajadas": {f"{k:g}": float(v) for k, v in bajadas.items()},
        "tasa_real_media": float(real.mean()), "tasa_real_p5": float(real.quantile(0.05)),
        "tasa_real_p95": float(real.quantile(0.95)),
    }, d


if __name__ == "__main__":
    res, d = estimar()
    (RAIZ / "data" / "metadata").mkdir(parents=True, exist_ok=True)
    (RAIZ / "data" / "metadata" / "junta_logit.json").write_text(json.dumps(res, indent=2, ensure_ascii=False), encoding="utf-8")
    d.reset_index(names="mes").to_csv(RAIZ / "data" / "raw" / "junta_decisiones_2001_2026.csv", index=False)
    print(json.dumps(res, indent=2, ensure_ascii=False))

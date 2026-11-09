"""Decisión de la Junta: logit ordenado con tendencia de la inflación y brecha de Taylor.

Datos: TPM diaria del Banco de la República e inflación anual mensual del DANE, 2001-2026.
Una observación por cada mes del calendario actual de la Junta (enero, marzo, abril, junio,
julio, septiembre, octubre y diciembre) en todos los años, porque el simulador decide con
ese calendario. Antes de 2017 la Junta se reunía cada mes; la decisión de esos años es el
cambio acumulado desde la reunión de calendario anterior.

Fechas: la TPM que decide la reunión del mes t entra en vigencia casi siempre el primer día
hábil del mes t + 1. Por eso el nivel que deja cada reunión es la TPM del día 15 del mes
siguiente, y la decisión es ese nivel menos el que dejó la reunión anterior. La brecha usa
la TPM vigente antes de decidir (la que dejó la reunión anterior) y la inflación conocida
en la reunión (la del mes t − 1). Así ninguna variable explicativa contiene el cambio que
se quiere explicar. Decisión: baja, mantiene o sube.

  latente = b_tend·(π(t−1) − π(t−4)) + b_taylor·[TPM previa − (r* + π(t−1) + 0,5·(π(t−1) − meta))]
  P(baja) = logística(k1 − latente); P(sube) = 1 − logística(k2 − latente)

Escribe data/metadata/junta_logit.json y la tabla de decisiones para la hoja TPM_Logit.
Uso: python tools/estimar_junta.py           (2001-2026, referencia)
     python tools/estimar_junta.py 2017      (solo desde 2017, sensibilidad; no escribe archivos)
"""
from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd
import statsmodels.api as sm
from scipy.stats import chi2
from statsmodels.miscmodels.ordinal_model import OrderedModel

RAIZ = Path(__file__).resolve().parents[1]
CRUDOS = RAIZ / "data" / "raw"
R_NEUTRAL = 1.0
MESES_JUNTA = [1, 3, 4, 6, 7, 9, 10, 12]
METAS = {2001: 8, 2002: 6, 2003: 5.5, 2004: 5.5, 2005: 5, 2006: 4.5, 2007: 4, 2008: 4, 2009: 5}


def datos(desde: int = 2001) -> pd.DataFrame:
    diaria = pd.read_csv(CRUDOS / "banrep_tpm_diaria.csv", parse_dates=["Date"]).set_index("Date")["value"].sort_index()
    inf = pd.read_csv(CRUDOS / "inflacion_mensual.csv", parse_dates=["fecha"]).set_index("fecha")["inf"]
    inf.index = inf.index.to_period("M")
    filas = []
    for per in pd.period_range("2001-01", diaria.index[-1].to_period("M"), freq="M"):
        if per.month not in MESES_JUNTA:
            continue
        dia = (per + 1).to_timestamp() + pd.Timedelta(days=14)
        if dia > diaria.index[-1] or (per - 4) not in inf.index or (per - 1) not in inf.index:
            continue
        filas.append({"mes": per.to_timestamp("M"), "tpm": float(diaria.asof(dia)),
                      "inf": float(inf.get(per, np.nan)), "pi_conocida": float(inf[per - 1]),
                      "dpi3": float(inf[per - 1] - inf[per - 4]), "meta": METAS.get(per.year, 3.0)})
    d = pd.DataFrame(filas).set_index("mes")
    d["tpm_previa"] = d.tpm.shift(1)
    pi = d.pi_conocida
    d["brecha_taylor"] = d.tpm_previa - (R_NEUTRAL + pi + 0.5 * (pi - d.meta))
    d["cambio"] = (d.tpm - d.tpm_previa).round(4)
    d["decision"] = np.sign(d["cambio"])
    d = d.dropna(subset=["tpm_previa", "dpi3", "brecha_taylor"])
    return d[d.index.year >= desde]


def estimar(desde: int = 2001) -> dict:
    d = datos(desde)
    y = pd.Categorical(d.decision, categories=[-1, 0, 1], ordered=True)
    m0 = OrderedModel(y, d[["dpi3"]], distr="logit").fit(method="bfgs", disp=False)
    m1 = OrderedModel(y, d[["dpi3", "brecha_taylor"]], distr="logit").fit(method="bfgs", disp=False)
    # statsmodels guarda el segundo umbral como log del incremento.
    k1 = float(m1.params.iloc[2])
    k2 = k1 + float(np.exp(m1.params.iloc[3]))
    lr = 2 * (m1.llf - m0.llf)
    subidas = d.cambio[d.cambio > 0].value_counts(normalize=True).sort_index()
    bajadas = (-d.cambio[d.cambio < 0]).value_counts(normalize=True).sort_index()
    real = (d.tpm - d.inf).dropna()
    # Odds proporcionales: logit binario en cada corte; las pendientes deberían coincidir (Brant aproximado,
    # sin la covarianza entre cortes).
    X = sm.add_constant(d[["dpi3", "brecha_taylor"]])
    c1 = sm.Logit((d.decision >= 0).astype(int), X).fit(disp=False)
    c2 = sm.Logit((d.decision >= 1).astype(int), X).fit(disp=False)
    dif = (c1.params - c2.params)[["dpi3", "brecha_taylor"]].to_numpy()
    var = (c1.bse ** 2 + c2.bse ** 2)[["dpi3", "brecha_taylor"]].to_numpy()
    brant = float(np.sum(dif ** 2 / var))
    # Calibración: probabilidades predichas frente a frecuencias observadas por tercil del índice latente.
    lat = m1.params["dpi3"] * d.dpi3 + m1.params["brecha_taylor"] * d.brecha_taylor
    logis = lambda z: 1 / (1 + np.exp(-z))  # noqa: E731
    cal = pd.DataFrame({"p_baja": logis(k1 - lat), "p_sube": 1 - logis(k2 - lat), "obs_baja": d.decision == -1,
                        "obs_sube": d.decision == 1, "tercil": pd.qcut(lat, 3, labels=["bajo", "medio", "alto"])})
    cal = cal.groupby("tercil", observed=True).mean()
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
        "brant_w": brant, "brant_p": float(chi2.sf(brant, 2)),
        "calibracion": {k: {c: round(float(v), 3) for c, v in fila.items()} for k, fila in cal.iterrows()},
    }, d


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        res, _ = estimar(int(sys.argv[1]))
        print(json.dumps(res, indent=2, ensure_ascii=False))
        raise SystemExit
    res, d = estimar()
    (RAIZ / "data" / "metadata").mkdir(parents=True, exist_ok=True)
    (RAIZ / "data" / "metadata" / "junta_logit.json").write_text(json.dumps(res, indent=2, ensure_ascii=False), encoding="utf-8")
    d.reset_index(names="mes").to_csv(RAIZ / "data" / "raw" / "junta_decisiones_2001_2026.csv", index=False)
    print(json.dumps(res, indent=2, ensure_ascii=False))

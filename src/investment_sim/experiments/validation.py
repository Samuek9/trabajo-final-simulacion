"""Validación (16.2): lo que genera el modelo frente a lo observado.

Se comparan las series simuladas (escenario base) con los datos de data/raw y con
referencias del documento. No reemplaza a la verificación: aquí se juzgan los
supuestos económicos, no el código.
"""
from __future__ import annotations

import numpy as np
import pandas as pd
from scipy import stats

from ..config.parameters import Params
from ..data.loader import CRUDOS, serie
from ..models import equity, markov
from .runner import correr


def _duraciones(estados: np.ndarray, k: int, paso: int = 3) -> list[float]:
    """Duración de cada racha del estado k en una trayectoria trimestral (en trimestres)."""
    x = estados[::paso]
    out, n = [], 0
    for v in x:
        if v == k:
            n += 1
        elif n:
            out.append(n)
            n = 0
    return out


def validar(p: Params, n: int = 500, semilla: int | None = None, procesos: int = 1) -> pd.DataFrame:
    r = correr(p, ["E3"], n=n, semilla=semilla, procesos=procesos,
               guardar=("inflacion", "tpm", "estado", "etf", "inmueble", "ocupado", "cuenta"))
    tr = r.trayectorias["E3"]
    filas = []

    def fila(proceso, medida, simulado, observado, referencia, comentario="", tipo="Reproduce insumos"):
        # "Reproduce insumos": se compara con los mismos datos con que se estimó el proceso (verifica los
        # generadores). "Contraste": se compara con algo que el modelo no usó al estimarse (validación).
        filas.append({"proceso": proceso, "medida": medida, "tipo": tipo, "simulado": simulado, "observado": observado,
                      "referencia": referencia, "lectura": comentario})

    # Inflación de diciembre
    inf_dic = tr["inflacion"][:, 12::12]
    obs = serie("inflacion_anual")
    obs_v = obs.iloc[:, 1].to_numpy(float)
    fila("Inflación", "Mediana de la inflación de diciembre (%)", np.median(inf_dic), np.median(obs_v), "DANE 2001-2025")
    fila("Inflación", "Percentil 5 y 95 (%)", f"{np.percentile(inf_dic, 5):.2f} y {np.percentile(inf_dic, 95):.2f}",
         f"{np.percentile(obs_v, 5):.2f} y {np.percentile(obs_v, 95):.2f}", "DANE 2001-2025")
    sim_ac = np.mean([np.corrcoef(np.log(f[:-1]), np.log(f[1:]))[0, 1] for f in inf_dic if np.std(f) > 0])
    fila("Inflación", "Autocorrelación de ln(π) de un año al siguiente", sim_ac,
         np.corrcoef(np.log(obs_v[:-1]), np.log(obs_v[1:]))[0, 1], "DANE 2001-2025",
         "Con 10 datos por réplica la autocorrelación muestral sale sesgada hacia abajo.")

    # TPM y tasa real
    tpm = tr["tpm"]
    real = (tpm - tr["inflacion"])[:, 12::12]
    hist = pd.read_csv(CRUDOS / "junta_decisiones_2001_2026.csv")
    real_obs = (hist.tpm - hist.inf).to_numpy(float)
    fila("TPM", "TPM mediana en el año 10 (%)", np.median(tpm[:, -1]), np.median(hist.tpm), "BanRep 2001-2026",
         tipo="Contraste")
    fila("TPM", "Tasa real media: TPM − inflación (puntos)", real.mean(), real_obs.mean(), "BanRep y DANE 2001-2026",
         "El modelo parte de una tasa real de 6 puntos (12,25 % con inflación de 6,3 %) y la baja despacio.",
         tipo="Contraste")
    tardia = (tpm - tr["inflacion"])[:, 61:]
    pi_tardia = float(np.median(tr["inflacion"][:, 61:]))
    fila("TPM", "Tasa real media en los años 6 a 10 (puntos)", tardia.mean(), real_obs.mean(), "BanRep y DANE 2001-2026",
         "Sin el arranque la brecha se reduce. Con la inflación simulada de esos años ("
         + f"{pi_tardia:.1f} %), la regla de Taylor del modelo pide {1 + 0.5 * (pi_tardia - 3):.1f} puntos; "
         .replace(".", ",") + "la tasa todavía va bajando hacia ahí.", tipo="Contraste")
    fila("TPM", "Réplicas con la TPM en el tope de 15 % al año 10", float((tpm[:, -1] >= p.f("tpm.max") - 1e-9).mean()), 0.0,
         "Máximo histórico desde 2001: 13,25 %", tipo="Contraste")

    # ETF: rendimientos que genera el modelo sobre los estados simulados.
    estados = tr["estado"][:, 1:]
    rng = np.random.default_rng(7)
    T = equity.choques(rng.random(estados.size), p).reshape(estados.shape)
    g = np.vectorize(lambda k, t: equity.rendimiento_log(markov.ESTADOS[k], t, p))(estados, T)
    obs_g = serie("etf")["g"].to_numpy(float)
    fila("ETF", "Rendimiento log mensual medio", float(g.mean()), float(obs_g.mean()), "BlackRock NAV 2011-2026")
    fila("ETF", "Desviación mensual (todos los estados)", float(g.std()), float(obs_g.std()), "BlackRock NAV 2011-2026")
    fila("ETF", "Desviación mensual en crisis", float(g[estados == 2].std()), 0.08783, "Hoja MARKOV (24 meses)")
    fila("ETF", "Asimetría", float(stats.skew(g.ravel())), float(stats.skew(obs_g)), "BlackRock NAV 2011-2026",
         "La t(5) es simétrica; la asimetría negativa de los datos solo la recoge Jones-Faddy.", tipo="Contraste")
    fila("ETF", "Curtosis en exceso", float(stats.kurtosis(g.ravel())), float(stats.kurtosis(obs_g)), "BlackRock NAV 2011-2026",
         tipo="Contraste")
    fila("ETF", "Peor mes en 10 años (mediana de las réplicas)", float(np.median(g.min(axis=1))), float(obs_g.min()),
         "Marzo de 2020: −31 % en logaritmo", tipo="Contraste")

    # Estados económicos
    est = tr["estado"][:, 3::3]
    for k, nombre in enumerate(markov.ESTADOS):
        fila("Estados", f"Proporción del tiempo en {nombre}", float((est == k).mean()),
             [0.366, 0.488, 0.146][k], "DANE PIB 2006-2026")
    dur_c = [d for f in tr["estado"] for d in _duraciones(f, 2)]
    fila("Estados", "Duración media de una crisis (trimestres)", float(np.mean(dur_c)) if dur_c else np.nan, 4.0,
         "DANE PIB 2006-2026", "El 4,0 observado es 1 / (1 − 0,75), la misma matriz; las rachas que tocan el final del "
         "horizonte quedan truncadas.")

    # Inmueble
    inm = tr["inmueble"]
    with np.errstate(divide="ignore", invalid="ignore"):
        rv = np.log(inm[:, 24::12] / inm[:, 12:-1:12])
    rv = rv[np.isfinite(rv)]
    ipvu = serie("ipvu")
    r_obs = ipvu[ipvu.fecha >= 2001].rn.to_numpy(float)
    fila("Valorización", "Valorización nominal media (ln anual)", float(rv.mean()), float(r_obs.mean()), "IPVU 2001-2025",
         "El primer año parte de la variación de 2025 (0,0809) y se acerca a la media del AR(1).")
    fila("Valorización", "Desviación anual", float(rv.std()), float(r_obs.std(ddof=1)), "IPVU 2001-2025")
    # Tiempo desocupado medido en continuo, el mismo indicador de resumen.csv.
    fila("Vacancia", "Fracción del tiempo desocupado", float(r.tabla.tiempo_vacante.mean()), 0.006,
         "Ciencuadras 2026: 0,6 % de vacancia en Medellín",
         "El modelo es más conservador a propósito: 10 % de no renovación en estado normal.", tipo="Contraste")
    return pd.DataFrame(filas)


def margen_cdt_simulado(p: Params, n: int = 20000, semilla: int = 1) -> dict[str, float]:
    """Compara la distribución del margen sorteado por el modelo con la observada."""
    from ..data.loader import residuos_cdt
    from ..models import interest_rates as tasas
    rng = np.random.default_rng(semilla)
    obs = serie("cdt")
    d13 = rng.choice(obs["d13"].to_numpy(float), n)
    res = residuos_cdt(float(p.tabla["cdt.a"].base), float(p.tabla["cdt.b"].base))
    m = np.array([tasas.margen_cdt(d, z, u, p, res) for d, z, u in zip(d13, rng.standard_normal(n), rng.random(n))])
    s = obs["s"].to_numpy(float)
    return {"media_sim": float(m.mean()), "media_obs": float(s.mean()), "sd_sim": float(m.std()), "sd_obs": float(s.std()),
            "p5_sim": float(np.percentile(m, 5)), "p5_obs": float(np.percentile(s, 5)),
            "p95_sim": float(np.percentile(m, 95)), "p95_obs": float(np.percentile(s, 95))}

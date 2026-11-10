"""Corre todos los experimentos del informe y guarda los resultados en outputs/resultados.

  python tools/correr_experimentos.py [procesos]

1. Las 12 configuraciones (4 estrategias × 3 escenarios) con las réplicas de PARAMETROS.
2. Verificación del número de réplicas (piloto de 200) y precisión con 100 a 2.400.
3. Diferencias pareadas entre estrategias.
4. Experimento capital-meta (base y estrés).
5. Sensibilidad (tornado) y sensibilidad a la elección de modelos.
6. Validación frente a los datos.
7. Traza de eventos de una réplica de ejemplo.
"""
from __future__ import annotations

import json
import sys
import time
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / "src"))

from investment_sim import __version__  # noqa: E402
from investment_sim.config.parameters import cargar  # noqa: E402
from investment_sim.experiments import capital_goal, sensitivity, validation  # noqa: E402
from investment_sim.experiments.runner import correr, una_replica  # noqa: E402
from investment_sim.metrics.summary import bandas, ic_proporcion, pareada, replicas_necesarias, resumen  # noqa: E402

SALIDA = RAIZ / "outputs" / "resultados"
ESTRATEGIAS = ["E1", "E2", "E3", "E4"]


def main(procesos: int) -> None:
    SALIDA.mkdir(parents=True, exist_ok=True)
    t0 = time.perf_counter()
    meta = {"version_modelo": __version__, "fecha": datetime.now().isoformat(timespec="seconds"), "procesos": procesos}
    base = cargar("base")
    n = base.i("exp.replicas")
    semilla = base.i("exp.semilla")
    meta.update({"replicas": n, "semilla_base": semilla})

    # 2a. Piloto de 200 réplicas para revisar si alcanza n = 1.068 (punto 14).
    piloto = correr(base, ESTRATEGIAS, n=200, semilla=semilla + 10_000_000, guardar=(), procesos=procesos)
    meta["replicas_necesarias_W_real"] = {e: replicas_necesarias(g.W_real.to_numpy())
                                          for e, g in piloto.tabla.groupby("estrategia")}
    print("piloto", meta["replicas_necesarias_W_real"], flush=True)

    # 1. Doce configuraciones.
    tablas, resumenes, pares, bandas_todas = [], [], [], []
    for esc in ("base", "adverso", "estres"):
        p = cargar(esc)
        r = correr(p, ESTRATEGIAS, n=n, semilla=semilla, procesos=procesos,
                   guardar=("W_real", "W", "cuenta", "iliquido", "estado", "inmueble"))
        tablas.append(r.tabla)
        resumenes.append(resumen(r.tabla))
        for a, b in (("E4", "E2"), ("E3", "E2"), ("E2", "E1"), ("E3", "E1"), ("E4", "E1"), ("E4", "E3")):
            for col in ("W_real", "meta", "algun_episodio", "caida_max"):
                pares.append({"escenario": esc, **pareada(r.tabla, a, b, col)})
        for e in ESTRATEGIAS:
            for serie in ("W_real", "cuenta"):
                b = bandas(r.trayectorias[e][serie])
                b.insert(0, "mes", np.arange(len(b)))
                b.insert(0, "serie", serie)
                b.insert(0, "estrategia", e)
                b.insert(0, "escenario", esc)
                bandas_todas.append(b)
            il = r.trayectorias[e]["iliquido"].mean(axis=0)
            bandas_todas.append(pd.DataFrame({"escenario": esc, "estrategia": e, "serie": "frac_iliquido",
                                              "mes": np.arange(len(il)), "p50": il}))
        print(esc, f"{r.segundos:.0f} s", flush=True)
    tabla = pd.concat(tablas, ignore_index=True)
    tabla.to_csv(SALIDA / "replicas.csv.gz", index=False, compression="gzip")
    pd.concat(resumenes, ignore_index=True).to_csv(SALIDA / "resumen.csv", index=False)
    pd.DataFrame(pares).to_csv(SALIDA / "comparaciones_pareadas.csv", index=False)
    pd.concat(bandas_todas, ignore_index=True).to_csv(SALIDA / "bandas.csv", index=False)

    # 2b. Precisión de la probabilidad de meta según el número de réplicas (E2 base).
    prec = []
    e2 = tabla[(tabla.escenario == "base") & (tabla.estrategia == "E2")]
    extra = correr(base, ["E2"], n=2400, semilla=semilla, guardar=(), procesos=procesos).tabla
    for k in (100, 500, 1068, 2400):
        x = extra.meta.to_numpy()[:k]
        p_, lo, hi = ic_proporcion(float(x.sum()), k)
        prec.append({"replicas": k, "p_meta": p_, "lo": lo, "hi": hi, "semiancho": (hi - lo) / 2,
                     "semiancho_teorico_max": 1.96 * np.sqrt(0.25 / k)})
    pd.DataFrame(prec).to_csv(SALIDA / "precision_replicas.csv", index=False)
    assert np.allclose(e2.W_real.to_numpy(), extra.W_real.to_numpy()[:len(e2)])

    # 4. Capital-meta.
    caps = [400e6, 500e6, 600e6, 700e6, 800e6, 900e6, 1000e6]
    metas = [700e6, 800e6, 1000e6, 1200e6, 1500e6]
    cm = []
    for esc in ("base", "estres"):
        cm.append(capital_goal.matriz(cargar(esc), caps, metas, ESTRATEGIAS, n=500, semilla=semilla, procesos=procesos))
    cm = pd.concat(cm, ignore_index=True)
    cm.to_csv(SALIDA / "capital_meta.csv", index=False)
    # Banda del capital necesario: dónde cruzan 80 % los extremos del IC de Wilson de cada punto de la malla.
    cm_hi, cm_lo = cm.assign(p_meta=cm.hi), cm.assign(p_meta=cm.lo)
    necesario = [{"escenario": esc, "estrategia": e, "meta": m,
                  "capital_80": capital_goal.capital_necesario(cm[cm.escenario == esc], m, e, 0.8),
                  "capital_80_lo": capital_goal.capital_necesario(cm_hi[cm_hi.escenario == esc], m, e, 0.8),
                  "capital_80_hi": capital_goal.capital_necesario(cm_lo[cm_lo.escenario == esc], m, e, 0.8),
                  "capital_50": capital_goal.capital_necesario(cm[cm.escenario == esc], m, e, 0.5)}
                 for esc in ("base", "estres") for e in ESTRATEGIAS for m in metas]
    pd.DataFrame(necesario).to_csv(SALIDA / "capital_necesario.csv", index=False)
    print("capital-meta listo", flush=True)

    # 5. Sensibilidad.
    sensitivity.tornado(base, ESTRATEGIAS, n=400, semilla=semilla, procesos=procesos).to_csv(SALIDA / "tornado.csv", index=False)
    sensitivity.sensibilidad_modelos(base, ESTRATEGIAS, n=400, semilla=semilla, procesos=procesos).to_csv(
        SALIDA / "sensibilidad_modelos.csv", index=False)
    sensitivity.interaccion(base, ESTRATEGIAS, n=400, semilla=semilla, procesos=procesos).to_csv(
        SALIDA / "interaccion.csv", index=False)
    print("sensibilidad lista", flush=True)

    # 6. Validación.
    validation.validar(base, n=500, semilla=semilla, procesos=procesos).to_csv(SALIDA / "validacion.csv", index=False)
    (SALIDA / "validacion_cdt.json").write_text(json.dumps(validation.margen_cdt_simulado(base), indent=2), encoding="utf-8")

    # 7. Traza de ejemplo (E4, base, réplica 0) y efecto de mirar la meta en pesos nominales.
    ej = una_replica(base, "E4", semilla, traza=True)
    pd.DataFrame(ej.traza).to_csv(SALIDA / "traza_E4_replica0.csv", index=False)
    nominal = []
    for (esc, e), g in tabla.groupby(["escenario", "estrategia"]):
        nominal.append({"escenario": esc, "estrategia": e, "p_meta_real": g.meta.mean(),
                        "p_meta_nominal": float((g.W_nominal >= base.f("inv.meta_real")).mean()),
                        "p_perdida_real": g.perdida_real.mean(),
                        "p_perdida_nominal": float((g.W_nominal < base.f("inv.capital")).mean()),
                        "inflacion_acumulada_p50": g.inflacion_acumulada.median()})
    pd.DataFrame(nominal).to_csv(SALIDA / "nominal_vs_real.csv", index=False)
    meta["segundos"] = round(time.perf_counter() - t0, 1)
    (SALIDA / "corrida.json").write_text(json.dumps(meta, indent=2, ensure_ascii=False), encoding="utf-8")
    print("listo", meta["segundos"], "s")


if __name__ == "__main__":
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 1)

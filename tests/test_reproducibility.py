"""Semillas, independencia entre réplicas y números aleatorios comunes (punto 14)."""
import numpy as np
import pytest

from investment_sim.engine.simulation import generadores
from investment_sim.experiments.runner import correr, una_replica
from investment_sim.metrics.summary import ic_proporcion, pareada


def test_misma_semilla_mismo_resultado(base):
    a = una_replica(base, "E4", 123)
    b = una_replica(base, "E4", 123)
    assert a.indicadores == b.indicadores or all(
        (np.isnan(a.indicadores[k]) and np.isnan(b.indicadores[k])) or a.indicadores[k] == b.indicadores[k]
        for k in a.indicadores)
    assert np.array_equal(a.trayectoria["W"], b.trayectoria["W"])


def test_semillas_distintas_dan_resultados_distintos(base):
    assert una_replica(base, "E3", 1).indicadores["W_real"] != una_replica(base, "E3", 2).indicadores["W_real"]


def test_flujos_por_proceso_son_independientes():
    g = generadores(99)
    x = {k: v.random(5000) for k, v in g.items()}
    claves = list(x)
    for i in range(len(claves)):
        for j in range(i + 1, len(claves)):
            assert abs(np.corrcoef(x[claves[i]], x[claves[j]])[0, 1]) < 0.06


def test_numeros_comunes_las_estrategias_ven_la_misma_economia(base):
    r = correr(base, n=5, semilla=40)
    tr = r.trayectorias
    for k in ("estado", "tpm", "inflacion"):
        assert np.array_equal(tr["E1"][k], tr["E3"][k])
        assert np.array_equal(tr["E2"][k], tr["E4"][k])


def test_resultado_no_depende_del_numero_de_procesos(base):
    a = correr(base, ["E2", "E4"], n=12, semilla=5, guardar=())
    b = correr(base, ["E2", "E4"], n=12, semilla=5, guardar=(), procesos=2)
    assert np.allclose(a.tabla.W_real.to_numpy(), b.tabla.W_real.to_numpy())


def test_numeros_comunes_reducen_la_varianza(base):
    r = correr(base, ["E2", "E4"], n=60, semilla=50, guardar=())
    d = pareada(r.tabla, "E4", "E2")
    assert d["reduccion_varianza"] > 0.5


def test_wilson_dentro_de_cero_y_uno():
    p, lo, hi = ic_proporcion(0, 100)
    assert p == 0 and lo == pytest.approx(0, abs=1e-12) and hi > 0
    p, lo, hi = ic_proporcion(100, 100)
    assert hi == 1 and lo < 1

"""Reglas de decisión, liquidez y casos extremos (16.1)."""
import numpy as np
import pytest

from investment_sim.engine.simulation import Simulacion
from investment_sim.experiments.runner import correr, residuos, una_replica
from investment_sim.strategies.rules import Estrategia


@pytest.mark.parametrize("e,pesos", [("E1", (0.20, 0.60, 0.20, 0)), ("E2", (0.10, 0.26, 0.21, 0.43)),
                                     ("E3", (0.07, 0.08, 0.30, 0.55))])
def test_asignacion_inicial(quieto, e, pesos):
    sim = Simulacion(quieto, Estrategia(e), 1, residuos(quieto), quieto)
    sim.iniciar()
    s = sim.s
    C = quieto.f("inv.capital")
    assert s.cuenta == pytest.approx(pesos[0] * C)
    assert s.valor_cdts(0) == pytest.approx(pesos[1] * C)
    assert s.etf == pytest.approx(pesos[2] * C)
    assert (s.inmueble.valor if s.inmueble.tiene else 0) == pytest.approx(pesos[3] * C)


def test_e4_se_mueve_hacia_el_objetivo_del_estado(quieto):
    """En crisis persistente, E4 lleva el ETF a 20 % del portafolio financiero."""
    p = quieto.con(**{"eco.estado_inicial": "Crisis", "markov.C_C": 1.0, "markov.C_N": 0.0})
    sim = Simulacion(p, Estrategia("E4"), 1, residuos(p), p)
    r = sim.correr()
    F = r.trayectoria["cuenta"][-1] + r.trayectoria["cdt"][-1] + r.trayectoria["etf"][-1]
    assert r.trayectoria["etf"][-1] / F == pytest.approx(0.20, abs=0.051)
    assert r.indicadores["n_rebalanceos"] >= 1
    e2 = una_replica(p, "E2", 1)
    assert e2.indicadores["n_rebalanceos"] == 0


def test_e4_cambia_con_el_estado(quieto):
    exp = quieto.con(**{"eco.estado_inicial": "Expansión", "markov.E_E": 1.0, "markov.E_N": 0.0})
    r = una_replica(exp, "E4", 2)
    F = r.trayectoria["cuenta"][-1] + r.trayectoria["cdt"][-1] + r.trayectoria["etf"][-1]
    assert r.trayectoria["etf"][-1] / F == pytest.approx(0.55, abs=0.051)


def test_sin_liquidez_suficiente_hay_iliquidez_y_deuda(quieto):
    """E1 sin ETF ni CDT y necesidades gigantes: no hay de dónde sacar, queda deuda."""
    p = quieto.con(**{"E1.pesos": "1;0;0;0", "liq.tasa_N": 0.2, "liq.mediana": 200e6, "liq.sigma": 0.0})
    r = una_replica(p, "E1", 3)
    assert r.indicadores["episodios_iliquidez"] >= 1
    assert r.indicadores["deuda_final"] > 0
    assert r.indicadores["insolvente"] == 1.0


def test_orden_de_cobertura_etf_antes_que_cdt(quieto):
    p = quieto.con(**{"liq.tasa_N": 0.05, "liq.tasa_E": 0.05, "liq.tasa_C": 0.05, "liq.mediana": 30e6, "liq.sigma": 0.0})
    r = una_replica(p, "E1", 4)
    a = r.indicadores
    assert a["acum_cubierto_etf"] > 0
    if a["acum_cubierto_cdt"] > 0:
        assert a["acum_cubierto_etf"] >= 0.2 * 700e6 * 0.98  # el ETF se agotó antes de tocar el CDT


def test_patrimonio_alto_con_iliquidez(base):
    """E3 con necesidades grandes: el patrimonio sigue alto (inmueble) aunque falte efectivo."""
    p = base.con(**{"liq.mediana": 120e6, "liq.sigma": 0.3, "liq.tasa_N": 1 / 6, "liq.tasa_E": 1 / 6, "liq.tasa_C": 1 / 6})
    r = correr(p, ["E3"], n=30, semilla=10, guardar=("iliquido", "W", "inmueble"))
    t = r.tabla
    assert (t.algun_episodio > 0).any()
    ilq = r.trayectorias["E3"]["iliquido"]
    inm = r.trayectorias["E3"]["inmueble"]
    assert (inm[ilq] > 0).any(), "debe haber meses ilíquidos con el inmueble aún en el patrimonio"


def test_caida_extrema_del_etf(base):
    p = base.con(**{"estres.activo": 1, "etf.caida_estres": 0.9})
    sim = Simulacion(p, Estrategia("E3"), 5, residuos(p), base, traza=True)
    r = sim.correr()
    caidas = [x for x in r.traza if x["n"] == 16]
    assert len(caidas) == 1
    assert r.indicadores["caida_max"] > 0.1


def test_crisis_persistente(base):
    p = base.con(**{"eco.estado_inicial": "Crisis", "markov.C_C": 1.0, "markov.C_N": 0.0})
    r = una_replica(p, "E2", 6)
    assert (r.trayectoria["estado"][1:] == 2).all()


def test_sin_vacancia(base):
    p = base.con(**{"vac.p_E": 0.0, "vac.p_N": 0.0, "vac.p_C": 0.0})
    r = una_replica(p, "E3", 7)
    assert r.indicadores["n_vacancias"] == 0
    assert r.indicadores["tiempo_vacante"] == 0


def test_sin_eventos_extraordinarios(base):
    p = base.con(**{"liq.tasa_E": 0.0, "liq.tasa_N": 0.0, "liq.tasa_C": 0.0})
    r = una_replica(p, "E2", 8)
    assert r.indicadores["n_necesidades"] == 0
    assert r.indicadores["acum_reparaciones"] == 0


def test_reparaciones_cuando_se_activan(base):
    p = base.con(**{"rep.tasa": 1 / 12, "rep.mediana": 2e6, "rep.sigma": 0.5})
    r = una_replica(p, "E2", 9)
    assert r.indicadores["n_reparaciones"] > 0


def test_estres_fija_la_crisis_y_vuelve_a_normal(tabla):
    from investment_sim.config.parameters import cargar
    p = cargar("estres", tabla=tabla)
    sim = Simulacion(p, Estrategia("E2"), 12, residuos(p), p.fuera_de_crisis(), traza=True)
    r = sim.correr()
    ini = [x["t"] for x in r.traza if x["n"] == 16][0]
    fin = [x["t"] for x in r.traza if x["n"] == 17]
    est = r.trayectoria["estado"]
    hasta = int(np.floor(fin[0])) if fin else len(est) - 1
    assert (est[int(np.ceil(ini)):hasta] == 2).all()

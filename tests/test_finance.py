"""Verificación de los cálculos financieros con casos de resultado conocido (16.1)."""
import math

import numpy as np
import pytest

from investment_sim.engine.simulation import Simulacion
from investment_sim.experiments.runner import residuos, una_replica
from investment_sim.models import interest_rates as tasas
from investment_sim.strategies.rules import Estrategia

C = 700e6


@pytest.mark.parametrize("e", ["E1", "E2", "E3", "E4"])
def test_mundo_quieto_conserva_el_patrimonio(quieto, e):
    """Rendimientos, inflación, costos y eventos en cero: el patrimonio no cambia."""
    r = una_replica(quieto, e, 1)
    W = r.trayectoria["W"]
    assert np.allclose(W, C, rtol=1e-9)
    assert math.isclose(r.indicadores["W_real"], C, rel_tol=1e-9)
    assert r.indicadores["rent_real"] == pytest.approx(0, abs=1e-9)
    assert r.indicadores["caida_max"] == pytest.approx(0, abs=1e-9)


def test_inflacion_cero_real_igual_a_nominal(quieto):
    p = quieto.con(**{"etf.mu": 0.004, "etf.vol_N": 0.05, "etf.vol_E": 0.05, "etf.vol_C": 0.08})
    r = una_replica(p, "E3", 2)
    assert np.allclose(r.trayectoria["W"], r.trayectoria["W_real"], rtol=1e-9)


def test_cuenta_sola_crece_al_interes_compuesto(quieto):
    """Todo en la cuenta, TPM fija de 10 % y α = 1: crece a 10 % anual neto de retención."""
    p = quieto.con(**{"E1.pesos": "1;0;0;0", "tpm.inicial": 10.0, "tpm.hace_3m": 10.0, "cta.alfa": 1.0,
                      "imp.retencion_cuenta": 0.07})
    r = una_replica(p, "E1", 3)
    tasa_mes = 1.1 ** (1 / 12) - 1
    esperado = C * (1 + tasa_mes * 0.93) ** 120
    assert r.indicadores["W_nominal"] == pytest.approx(esperado, rel=1e-9)


def test_cdt_paga_tpm_mas_margen_cada_anio(quieto):
    p = quieto.con(**{"E1.pesos": "0;1;0;0", "tpm.inicial": 10.0, "tpm.hace_3m": 10.0, "inv.liquidez_min": 0.0,
                      "cdt.fase_estable": "1;0", "cdt.margen_min": -5.0, "cdt.margen_max": 5.0})
    r = una_replica(p, "E1", 4)
    # E1 renueva por el mismo capital: los intereses (11 % al año) se acumulan en la cuenta, que no rinde.
    assert r.indicadores["W_nominal"] == pytest.approx(C * (1 + 0.11 * 10), rel=1e-9)


def test_etf_sin_volatilidad_crece_exp_mu(quieto):
    p = quieto.con(**{"E1.pesos": "0;0;1;0", "etf.mu": 0.005, "inv.liquidez_min": 0.0})
    r = una_replica(p, "E1", 5)
    assert r.indicadores["W_nominal"] == pytest.approx(C * math.exp(0.005 * 120), rel=1e-9)


def test_comision_del_etf_se_cobra_al_comprar(quieto):
    p = quieto.con(**{"E1.pesos": "0;0;1;0", "etf.comision": 0.005, "inv.liquidez_min": 0.0})
    r = una_replica(p, "E1", 5)
    assert r.indicadores["W_nominal"] == pytest.approx(C * 0.995, rel=1e-9)


def test_inmueble_siempre_ocupado_cobra_todos_los_arriendos(quieto):
    p = quieto.con(**{"inm.canon_inicial": 0.006, "inm.comision_inmobiliaria": 0.0, "inm.reajuste_modelo": "ipc"})
    r = una_replica(p, "E3", 6)
    V = 0.55 * C
    assert r.indicadores["acum_arriendo_bruto"] == pytest.approx(0.006 * V * 120, rel=1e-9)
    assert r.indicadores["tiempo_vacante"] == 0.0
    assert r.trayectoria["ocupado"][1:].all()


def test_costos_de_compra_y_venta(quieto):
    p = quieto.con(**{"inm.costo_compra": 0.025, "inm.costo_venta": 0.0484})
    r = una_replica(p, "E2", 7)
    V = 0.43 * C
    assert r.indicadores["W_nominal"] == pytest.approx(C - 0.025 * V - 0.0484 * V, rel=1e-9)


def test_4x1000_sobre_lo_que_pasa_la_exencion(quieto):
    p = quieto.con(**{"imp.gmf": 0.004, "imp.gmf_exencion": 10e6, "E1.pesos": "0.8;0;0.2;0"})
    r = una_replica(p, "E1", 8)
    assert r.indicadores["acum_gmf"] == pytest.approx(0.004 * (0.2 * C - 10e6), rel=1e-9)



def test_renovar_un_cdt_no_paga_4x1000(quieto):
    # El capital que se renueva no sale del banco: solo la apertura inicial paga 4x1000.
    p = quieto.con(**{"imp.gmf": 0.004, "imp.gmf_exencion": 10e6, "E1.pesos": "0.4;0.6;0;0"})
    r = una_replica(p, "E1", 8)
    assert r.indicadores["acum_gmf"] == pytest.approx(0.004 * (0.6 * C - 10e6), rel=1e-9)

def test_venta_anticipada_de_cdt_formula():
    v = tasas.valor_venta_anticipada(100.0, 0.10, 0.10, 0.0, 1.0, 1.0)
    assert v == pytest.approx(110.0)
    v0 = tasas.valor_venta_anticipada(100.0, 0.10, 0.10, 0.0, 1.0, 0.0)
    assert v0 == pytest.approx(100.0)
    castigo = tasas.valor_venta_anticipada(100.0, 0.10, 0.10, 0.01, 1.0, 0.5)
    sin = tasas.valor_venta_anticipada(100.0, 0.10, 0.10, 0.0, 1.0, 0.5)
    assert castigo < sin


def test_canon_respeta_tope_e_ipc(base):
    from investment_sim.models.real_estate import reajuste_canon
    assert reajuste_canon(1e6, 1e9, 5.0, 10.0, base) == pytest.approx(1e6 * 1.05)  # tope: IPC previo
    assert reajuste_canon(1.2e6, 1e8, 5.0, 0.0, base) == pytest.approx(1e6)        # tope: 1 % del valor
    ipc = base.con(**{"inm.reajuste_modelo": "ipc"})
    assert reajuste_canon(1e6, 1e9, 4.0, 0.0, ipc) == pytest.approx(1.04e6)


def test_patrimonio_nominal_y_real_consistentes(base):
    r = una_replica(base, "E2", 9)
    tr = r.trayectoria
    assert np.allclose(tr["W_real"], tr["W"] / tr["indice"])
    assert np.allclose(tr["W"], tr["cuenta"] + tr["cdt"] + tr["etf"] + tr["inmueble"] - tr["deuda"])


def test_capital_fijo_o_proporcional(base):
    fijo = base.con(**{"inv.modo_inmueble": "fijo", "inv.capital": 500e6})
    sim = Simulacion(fijo, Estrategia("E2"), 1, residuos(fijo), fijo)
    sim.iniciar()
    assert sim.s.inmueble.valor == pytest.approx(300e6)
    prop = base.con(**{"inv.capital": 500e6})
    sim2 = Simulacion(prop, Estrategia("E2"), 1, residuos(prop), prop)
    sim2.iniciar()
    assert sim2.s.inmueble.valor == pytest.approx(0.43 * 500e6)


def test_subir_la_constante_del_cdt_sube_el_margen(base):
    """Los residuos salen de a y b estimados; si no, un cambio en a se cancelaría solo."""
    alto = base.con(**{"cdt.a": base.f("cdt.a") + 1.0})
    for u in (0.1, 0.5, 0.9):
        m0 = tasas.margen_cdt(0.0, 0.0, u, base, residuos(base))
        m1 = tasas.margen_cdt(0.0, 0.0, u, alto, residuos(alto))
        assert m1 - m0 == pytest.approx(1.0, abs=1e-9)

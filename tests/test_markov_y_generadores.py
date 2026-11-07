"""Verificación de la cadena de Markov y de los generadores aleatorios (16.1)."""
import numpy as np
import pytest
from scipy import stats

from investment_sim.config.parameters import ParametroInvalido, cargar
from investment_sim.models import equity, inflation, liquidity_shocks as liq, markov, vacancy


@pytest.mark.parametrize("esc", ["base", "adverso", "estres"])
def test_filas_de_la_matriz_suman_uno(tabla, esc):
    P = markov.matriz(cargar(esc, tabla=tabla))
    assert np.allclose(P.sum(axis=1), 1)
    assert (P >= 0).all()


def test_matriz_invalida_se_rechaza(base):
    with pytest.raises(ParametroInvalido):
        base.con(**{"markov.N_C": 0.5})


def test_frecuencias_simuladas_igualan_la_matriz(base):
    P = markov.matriz(base)
    rng = np.random.default_rng(0)
    for i, s in enumerate(markov.ESTADOS):
        u = rng.random(200_000)
        destino = np.array([markov.ESTADOS.index(markov.siguiente(s, x, P)) for x in u[:20000]])
        frec = np.bincount(destino, minlength=3) / len(destino)
        assert np.allclose(frec, P[i], atol=0.01)


def test_estacionaria_parecida_a_lo_observado(base):
    pi = markov.estacionaria(markov.matriz(base))
    assert np.allclose(pi, [0.32, 0.52, 0.16], atol=0.01)


def test_choques_del_etf_tienen_la_volatilidad_del_estado(base):
    rng = np.random.default_rng(1)
    T = equity.choques(rng.random(200_000), base)
    for estado, clave in (("Normal", "etf.vol_N"), ("Crisis", "etf.vol_C")):
        g = np.array([equity.rendimiento_log(estado, t, base) for t in T[:50_000]])
        assert g.std() == pytest.approx(base.f(clave), rel=0.05)
        assert g.mean() == pytest.approx(base.f("etf.mu"), abs=0.001)
    assert stats.kurtosis(T) > 3  # colas pesadas de la t(5)


def test_jones_faddy_reproduce_media_y_asimetria(base):
    p = base.con(**{"etf.modelo": "jf"})
    g = equity.choques(np.random.default_rng(2).random(200_000), p)
    assert g.mean() == pytest.approx(0.0043, abs=0.001)
    assert stats.skew(g) < 0


def test_inflacion_converge_a_su_nivel(base):
    rng = np.random.default_rng(3)
    pi = 6.29
    xs = []
    for _ in range(20000):
        pi = inflation.siguiente_diciembre(pi, rng.standard_normal(), base.f("inf.c"), base.f("inf.phi"), base.f("inf.sigma"))
        xs.append(pi)
    mediana = np.median(xs)
    assert mediana == pytest.approx(inflation.nivel_largo_plazo(base.f("inf.c"), base.f("inf.phi")), rel=0.03)
    assert min(xs) > 0


def test_duracion_de_vacancia_tiene_la_media_del_estado(base):
    u = np.random.default_rng(4).random(40_000)
    for estado, clave in (("Normal", "vac.media_N"), ("Crisis", "vac.media_C")):
        d = np.array([vacancy.duracion(estado, x, base) for x in u])
        assert d.mean() == pytest.approx(base.f(clave), rel=0.03)
        assert (d > 0).all()


def test_poisson_de_liquidez(base):
    u = np.random.default_rng(5).random(100_000)
    lam = liq.tasa("Normal", base)
    esperas = np.array([liq.espera(lam, x) for x in u])
    assert esperas.mean() == pytest.approx(24, rel=0.02)
    assert liq.espera(0.0, 0.5) == float("inf")
    z = np.random.default_rng(6).standard_normal(100_000)
    montos = np.array([liq.monto(x, 1.0, 1.0, base) for x in z])
    assert np.median(montos) == pytest.approx(20e6, rel=0.02)


def test_parametros_fuera_de_dominio(base):
    for cambio in ({"vac.p_N": 1.5}, {"etf.nu": 2.0}, {"inv.capital": -1}, {"E2.pesos": "0.5;0.5;0.5;0"},
                   {"inf.phi": 1.0}, {"etf.modelo": "normal"}):
        with pytest.raises(ParametroInvalido):
            base.con_dict(cambio)

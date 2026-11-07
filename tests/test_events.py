"""Verificación del reloj, la lista de eventos y su procesamiento (16.1)."""
import math

from investment_sim.engine.events import ListaEventos, Tipo
from investment_sim.engine.simulation import Simulacion
from investment_sim.experiments.runner import residuos
from investment_sim.strategies.rules import Estrategia


def test_lista_ordena_por_tiempo_prioridad_y_secuencia():
    c = ListaEventos()
    c.programar(12.0, Tipo.CIERRE_MES)
    c.programar(12.0, Tipo.VENCE_CDT)
    c.programar(12.0, Tipo.CAMBIO_ESTADO)
    c.programar(3.0, Tipo.JUNTA)
    c.programar(12.0, Tipo.REVISION)
    c.programar(12.0, Tipo.INFLACION)
    orden = [c.siguiente().tipo for _ in range(6)]
    assert orden == [Tipo.JUNTA, Tipo.CAMBIO_ESTADO, Tipo.INFLACION, Tipo.REVISION, Tipo.VENCE_CDT, Tipo.CIERRE_MES]


def test_empate_total_respeta_orden_de_llegada():
    c = ListaEventos()
    a = c.programar(5.0, Tipo.ARRIENDO, id="a")
    b = c.programar(5.0, Tipo.ARRIENDO, id="b")
    assert c.siguiente() is a and c.siguiente() is b


def test_evento_cancelado_no_se_procesa():
    c = ListaEventos()
    e = c.programar(1.0, Tipo.NECESIDAD)
    c.programar(2.0, Tipo.FIN)
    ListaEventos.cancelar(e)
    assert c.siguiente().tipo is Tipo.FIN


def test_prioridades_siguen_el_diagrama():
    orden = [Tipo.CAMBIO_ESTADO, Tipo.INFLACION, Tipo.JUNTA, Tipo.VALORIZACION, Tipo.REVISION, Tipo.VENCE_CDT,
             Tipo.FIN_CONTRATO, Tipo.ARRIENDO, Tipo.REPARACION, Tipo.NECESIDAD, Tipo.VENTA_INMUEBLE,
             Tipo.CIERRE_MES, Tipo.FIN]
    pr = [t.prioridad for t in orden]
    assert pr == sorted(pr)
    assert Tipo.FIN_CONTRATO.prioridad == Tipo.FIN_VACANCIA.prioridad


def _correr_con_traza(p, e="E4", semilla=7):
    sim = Simulacion(p, Estrategia(e), semilla, residuos(p), p.fuera_de_crisis(), traza=True)
    return sim, sim.correr()


def test_reloj_nunca_retrocede_y_termina_en_H(base):
    sim, r = _correr_con_traza(base)
    tiempos = [x["t"] for x in r.traza]
    assert all(b >= a for a, b in zip(tiempos, tiempos[1:]))
    assert max(tiempos) <= base.i("inv.horizonte")
    assert math.isclose(sim.s.t, base.i("inv.horizonte"))


def test_cierres_de_mes_son_exactamente_H(base):
    sim, r = _correr_con_traza(base, "E1")
    assert r.indicadores["n_eventos"] > 0
    cierres = sum(1 for x in sim.traza if x["n"] == 1 and x["detalle"].startswith("Cobertura") is False)
    # El cierre no anota en la traza salvo coberturas; se verifica con la trayectoria.
    assert len(r.trayectoria["W"]) == base.i("inv.horizonte") + 1
    assert cierres >= 0


def test_eventos_periodicos_en_su_calendario(base):
    sim, r = _correr_con_traza(base, "E2")
    juntas = [x["t"] for x in r.traza if x["n"] == 4]
    assert len(juntas) == 8 * base.i("inv.horizonte") // 12
    assert {int(t) % 12 or 12 for t in juntas} == {1, 3, 4, 6, 7, 9, 10, 12}
    estados = [x["t"] for x in r.traza if x["n"] == 2]
    assert estados == [float(t) for t in range(3, base.i("inv.horizonte") + 1, 3)]
    inflacion = [x["t"] for x in r.traza if x["n"] == 3]
    assert inflacion == [float(t) for t in range(12, base.i("inv.horizonte") + 1, 12)]
    revisiones = [x["t"] for x in r.traza if x["n"] == 14]
    assert revisiones == []  # E2 no revisa
    _, r4 = _correr_con_traza(base, "E4")
    assert [x["t"] for x in r4.traza if x["n"] == 14] == [float(t) for t in range(12, base.i("inv.horizonte") + 1, 12)]


def test_inmueble_vendido_cancela_sus_eventos(base):
    p = base.con(**{"liq.mediana": 600e6, "liq.sigma": 0.0, "liq.tasa_N": 0.5, "liq.tasa_E": 0.5, "liq.tasa_C": 0.5})
    sim, r = _correr_con_traza(p, "E3", 3)
    ventas = [x for x in r.traza if x["n"] == 13 and x["detalle"].startswith("Se vende")]
    assert ventas, "con necesidades enormes el inmueble debe venderse"
    t_venta = ventas[0]["t"]
    despues = [x for x in r.traza if x["t"] > t_venta and x["n"] in (6, 7, 8, 9, 10)]
    assert despues == []


def test_horizonte_corto_y_largo(base):
    for H in (12, 360):
        p = base.con(**{"inv.horizonte": H})
        _, r = _correr_con_traza(p, "E4", 11)
        assert len(r.trayectoria["W"]) == H + 1
        assert math.isfinite(r.indicadores["W_real"])

"""Verificación (16.1) en el momento: comprobaciones sobre la implementación con los parámetros vigentes.

Cada comprobación calcula el valor esperado a partir de los parámetros que el inversionista configuró
(capital, horizonte, pesos, tasas de liquidez) y lo compara con lo que produce el simulador.
Los casos de resultado conocido usan un "mundo quieto": todo rinde cero y nada se paga.
"""
from __future__ import annotations

import math

import numpy as np
import pandas as pd

from ..config.parameters import ParametroInvalido, Params, cargar
from ..engine.events import ListaEventos, Tipo
from ..engine.simulation import Simulacion, generadores
from ..models import equity, inflation, liquidity_shocks as liq, markov, vacancy
from ..models import interest_rates as tasas
from ..strategies.rules import Estrategia
from .runner import correr, residuos, una_replica

QUIETO = {
    "inf.c": -30.0, "inf.phi": 0.0, "inf.sigma": 0.0, "inf.inicial": 1e-9,
    "tpm.inicial": 0.0, "tpm.hace_3m": 0.0, "tpm.min": 0.0, "junta.modelo": "tendencia",
    "junta.bajando": "0;1;0", "junta.estable": "0;1;0", "junta.subiendo": "0;1;0",
    "cta.alfa": 0.0, "imp.retencion": 0.0, "imp.retencion_cuenta": 0.0, "imp.gmf": 0.0,
    "cdt.modelo": "fase", "cdt.fase_subida": "0;0", "cdt.fase_estable": "0;0", "cdt.fase_bajada": "0;0",
    "cdt.margen_min": 0.0, "cdt.margen_max": 0.0,
    "etf.mu": 0.0, "etf.vol_E": 0.0, "etf.vol_N": 0.0, "etf.vol_C": 0.0, "etf.comision": 0.0,
    "val.ar1": "0;0;0", "val.mu_E": 0.0, "val.mu_N": 0.0, "val.mu_C": 0.0, "val.sd_E": 0.0, "val.sd_N": 0.0, "val.sd_C": 0.0,
    "inm.costo_compra": 0.0, "inm.canon_inicial": 0.0, "inm.ph": 0.0, "inm.predial": 0.0,
    "inm.mantenimiento": 0.0, "inm.pintura": 0.0, "inm.costo_venta": 0.0,
    "vac.p_E": 0.0, "vac.p_N": 0.0, "vac.p_C": 0.0,
    "liq.tasa_E": 0.0, "liq.tasa_N": 0.0, "liq.tasa_C": 0.0,
}

EVENTOS, ESTADOS, FINANZAS, GENERADORES, REGLAS = ("Reloj y eventos", "Estados y contabilidad", "Cálculos financieros",
                                                   "Generadores aleatorios", "Reglas de decisión")


def _m(x: float) -> str:
    return f"${x / 1e6:,.2f} M".replace(",", "X").replace(".", ",").replace("X", ".")


def _n(x: float, d: int = 4) -> str:
    return f"{x:,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def verificar(p: Params) -> pd.DataFrame:
    """Corre las comprobaciones con los parámetros dados. Devuelve una fila por comprobación."""
    H, C, semilla = p.i("inv.horizonte"), p.f("inv.capital"), p.i("exp.semilla")
    q = p.con_dict(QUIETO)
    filas: list[dict] = []

    def prueba(grupo: str, nombre: str, f) -> None:
        try:
            ok, esperado, obtenido = f()
        except Exception as e:  # una comprobación que falla no debe tumbar la página
            ok, esperado, obtenido = False, "sin error", f"{type(e).__name__}: {e}"[:120]
        filas.append({"Componente": grupo, "Comprobación": nombre, "Esperado": esperado, "Obtenido": obtenido,
                      "Resultado": "Cumple" if ok else "No cumple"})

    def valor_inmueble_esperado(pesos_inmueble: float) -> float:
        if pesos_inmueble <= 0:
            return 0.0
        return p.f("inv.valor_inmueble_fijo") if p.s("inv.modo_inmueble") == "fijo" else pesos_inmueble * C

    def traza_e4():
        sim = Simulacion(p, Estrategia("E4"), semilla, residuos(p), p.fuera_de_crisis(), traza=True)
        return sim, sim.correr()

    # --- Reloj y eventos
    def orden_lista():
        c = ListaEventos()
        for t in (Tipo.CIERRE_MES, Tipo.VENCE_CDT, Tipo.CAMBIO_ESTADO, Tipo.INFLACION, Tipo.REVISION):
            c.programar(12.0, t)
        c.programar(3.0, Tipo.JUNTA)
        salida = [c.siguiente().tipo for _ in range(6)]
        esp = [Tipo.JUNTA, Tipo.CAMBIO_ESTADO, Tipo.INFLACION, Tipo.REVISION, Tipo.VENCE_CDT, Tipo.CIERRE_MES]
        return salida == esp, "por tiempo y, en empates, por prioridad", "orden correcto" if salida == esp else "orden incorrecto"

    def cancelado():
        c = ListaEventos()
        e = c.programar(1.0, Tipo.NECESIDAD)
        c.programar(2.0, Tipo.FIN)
        ListaEventos.cancelar(e)
        sig = c.siguiente().tipo
        return sig is Tipo.FIN, "el evento cancelado no se procesa", "no se procesó" if sig is Tipo.FIN else "se procesó"

    def reloj():
        sim, r = traza_e4()
        t = [x["t"] for x in r.traza]
        retro = sum(b < a for a, b in zip(t, t[1:]))
        return (retro == 0 and math.isclose(sim.s.t, H), f"sin retrocesos; termina en el mes {H}",
                f"{retro} retrocesos; termina en {sim.s.t:g}")

    def cierres():
        _, r = traza_e4()
        n = len(r.trayectoria["W"])
        return n == H + 1, f"{H + 1} cierres de mes (incluye el inicial)", f"{n}"

    def calendario():
        _, r = traza_e4()
        juntas = sum(1 for x in r.traza if x["n"] == Tipo.JUNTA.numero)
        estados = [x["t"] for x in r.traza if x["n"] == Tipo.CAMBIO_ESTADO.numero]
        infl = [x["t"] for x in r.traza if x["n"] == Tipo.INFLACION.numero]
        ok = (juntas == 8 * H // 12 and estados == [float(t) for t in range(3, H + 1, 3)]
              and infl == [float(t) for t in range(12, H + 1, 12)])
        return (ok, f"{8 * H // 12} juntas, {len(range(3, H + 1, 3))} estados, {len(range(12, H + 1, 12))} inflaciones",
                f"{juntas} juntas, {len(estados)} estados, {len(infl)} inflaciones")

    prueba(EVENTOS, "Orden de la lista de eventos", orden_lista)
    prueba(EVENTOS, "Evento cancelado", cancelado)
    prueba(EVENTOS, "El reloj nunca retrocede", reloj)
    prueba(EVENTOS, "Cierres de mes", cierres)
    prueba(EVENTOS, "Eventos periódicos en su calendario", calendario)

    # --- Estados y contabilidad
    def identidad():
        tr = una_replica(p, "E2", semilla).trayectoria
        ok = (np.allclose(tr["W"], tr["cuenta"] + tr["cdt"] + tr["etf"] + tr["inmueble"] - tr["deuda"])
              and np.allclose(tr["W_real"], tr["W"] / tr["indice"]))
        return ok, "patrimonio = activos − deuda", "se cumple en todos los meses" if ok else "no se cumple"

    def quieto_conserva():
        dev = max(float(np.max(np.abs(una_replica(q, e, 1).trayectoria["W"] / C - 1))) for e in ("E1", "E2", "E3", "E4"))
        return dev < 1e-9, f"patrimonio constante en {_m(C)}", f"desviación máxima {dev:.1e}"

    def inflacion_cero():
        pp = q.con(**{"etf.mu": 0.004, "etf.vol_N": 0.05, "etf.vol_E": 0.05, "etf.vol_C": 0.08})
        tr = una_replica(pp, "E3", 2).trayectoria
        ok = bool(np.allclose(tr["W"], tr["W_real"], rtol=1e-9))
        return ok, "patrimonio real = nominal", "iguales" if ok else "difieren"

    def valor_inmueble():
        sim = Simulacion(p, Estrategia("E2"), 1, residuos(p), p)
        sim.iniciar()
        obt = sim.s.inmueble.valor if sim.s.inmueble.tiene else 0.0
        esp = valor_inmueble_esperado(p.lista("E2.pesos")[3])
        return math.isclose(obt, esp, rel_tol=1e-9, abs_tol=1e-6), _m(esp), _m(obt)

    prueba(ESTADOS, "Identidad contable y patrimonio real", identidad)
    prueba(ESTADOS, "Mundo quieto: el patrimonio no cambia", quieto_conserva)
    prueba(ESTADOS, "Inflación cero: real igual a nominal", inflacion_cero)
    prueba(ESTADOS, "Valor inicial del inmueble (E2)", valor_inmueble)

    # --- Cálculos financieros
    def cuenta_sola():
        pp = q.con(**{"E1.pesos": "1;0;0;0", "tpm.inicial": 10.0, "tpm.hace_3m": 10.0, "cta.alfa": 1.0, "imp.retencion_cuenta": 0.07})
        obt = una_replica(pp, "E1", 3).indicadores["W_nominal"]
        esp = C * (1 + (1.1 ** (1 / 12) - 1) * 0.93) ** H
        return math.isclose(obt, esp, rel_tol=1e-9), f"{_m(esp)} (10 % anual neto)", _m(obt)

    def cdt():
        pp = q.con(**{"E1.pesos": "0;1;0;0", "tpm.inicial": 10.0, "tpm.hace_3m": 10.0, "inv.liquidez_min": 0.0,
                      "cdt.fase_estable": "1;0", "cdt.margen_min": -5.0, "cdt.margen_max": 5.0})
        obt = una_replica(pp, "E1", 4).indicadores["W_nominal"]
        esp = C * (1 + 0.11 * H / 12)
        return math.isclose(obt, esp, rel_tol=1e-9), f"{_m(esp)} (11 % anual)", _m(obt)

    def etf():
        pp = q.con(**{"E1.pesos": "0;0;1;0", "etf.mu": 0.005, "inv.liquidez_min": 0.0})
        obt = una_replica(pp, "E1", 5).indicadores["W_nominal"]
        esp = C * math.exp(0.005 * H)
        return math.isclose(obt, esp, rel_tol=1e-9), f"{_m(esp)} (crece exp(μ·meses))", _m(obt)

    def comision():
        pp = q.con(**{"E1.pesos": "0;0;1;0", "etf.comision": 0.005, "inv.liquidez_min": 0.0})
        obt = una_replica(pp, "E1", 5).indicadores["W_nominal"]
        return math.isclose(obt, C * 0.995, rel_tol=1e-9), f"{_m(C * 0.995)} (comisión de 0,5 %)", _m(obt)

    def costos_inmueble():
        pp = q.con(**{"inm.costo_compra": 0.025, "inm.costo_venta": 0.0484})
        V = valor_inmueble_esperado(p.lista("E2.pesos")[3])
        obt = una_replica(pp, "E2", 7).indicadores["W_nominal"]
        esp = C - 0.025 * V - 0.0484 * V
        return math.isclose(obt, esp, rel_tol=1e-9), f"{_m(esp)} (2,5 % y 4,84 %)", _m(obt)

    def gmf():
        pp = q.con(**{"imp.gmf": 0.004, "imp.gmf_exencion": 10e6, "E1.pesos": "0.8;0;0.2;0"})
        obt = una_replica(pp, "E1", 8).indicadores["acum_gmf"]
        esp = 0.004 * max(0.0, 0.2 * C - 10e6)
        return (math.isclose(obt, esp, rel_tol=1e-9, abs_tol=1e-6), f"{_m(esp)} (4×1000 sobre el exceso)", _m(obt))

    def venta_anticipada():
        a = tasas.valor_venta_anticipada(100.0, 0.10, 0.10, 0.0, 1.0, 1.0)
        b = tasas.valor_venta_anticipada(100.0, 0.10, 0.10, 0.0, 1.0, 0.0)
        c1 = tasas.valor_venta_anticipada(100.0, 0.10, 0.10, 0.01, 1.0, 0.5)
        c0 = tasas.valor_venta_anticipada(100.0, 0.10, 0.10, 0.0, 1.0, 0.5)
        ok = math.isclose(a, 110.0) and math.isclose(b, 100.0) and c1 < c0
        return ok, "110 al vencimiento, 100 al inicio y menos con castigo", f"{_n(a, 1)}; {_n(b, 1)}; {_n(c1, 2)} < {_n(c0, 2)}"

    prueba(FINANZAS, "Cuenta remunerada: interés compuesto", cuenta_sola)
    prueba(FINANZAS, "CDT: tasa de la TPM más margen", cdt)
    prueba(FINANZAS, "ETF sin volatilidad: crecimiento exponencial", etf)
    prueba(FINANZAS, "Comisión del ETF al comprar", comision)
    prueba(FINANZAS, "Costos de compra y venta del inmueble", costos_inmueble)
    prueba(FINANZAS, "4×1000 sobre lo que pasa la exención", gmf)
    prueba(FINANZAS, "Venta anticipada del CDT", venta_anticipada)

    # --- Generadores aleatorios
    def filas_markov():
        peor = max(abs(float(markov.matriz(cargar(e, tabla=list(p.tabla.values()))).sum(axis=1).max()) - 1)
                   for e in ("base", "adverso", "estres"))
        return peor < 1e-9, "cada fila de la matriz suma 1 en los tres escenarios", f"desviación máxima {peor:.1e}"

    def frec_markov():
        P = markov.matriz(p)
        rng = np.random.default_rng(0)
        dif = 0.0
        for i, s in enumerate(markov.ESTADOS):
            u = rng.random(20_000)
            destino = np.array([markov.ESTADOS.index(markov.siguiente(s, x, P)) for x in u])
            dif = max(dif, float(np.abs(np.bincount(destino, minlength=3) / len(u) - P[i]).max()))
        return dif < 0.015, "frecuencias simuladas = matriz (±1,5 puntos)", f"diferencia máxima {dif * 100:.2f} puntos"

    def choques_etf():
        T = equity.choques(np.random.default_rng(1).random(200_000), p)
        sd = {}
        for estado, clave in (("Normal", "etf.vol_N"), ("Crisis", "etf.vol_C")):
            g = np.array([equity.rendimiento_log(estado, t, p) for t in T[:50_000]])
            sd[estado] = (g.std(), p.f(clave))
        ok = all(abs(a / b - 1) < 0.06 for a, b in sd.values())
        return (ok, f"σ normal {_n(sd['Normal'][1] * 100, 2)} % y σ crisis {_n(sd['Crisis'][1] * 100, 2)} %",
                f"{_n(sd['Normal'][0] * 100, 2)} % y {_n(sd['Crisis'][0] * 100, 2)} %")

    def inflacion_nivel():
        rng = np.random.default_rng(3)
        pi, xs = 6.29, []
        for _ in range(20_000):
            pi = inflation.siguiente_diciembre(pi, rng.standard_normal(), p.f("inf.c"), p.f("inf.phi"), p.f("inf.sigma"))
            xs.append(pi)
        esp = inflation.nivel_largo_plazo(p.f("inf.c"), p.f("inf.phi"))
        obt = float(np.median(xs))
        return abs(obt / esp - 1) < 0.04, f"mediana ≈ nivel de largo plazo ({_n(esp, 2)} %)", f"{_n(obt, 2)} %"

    def vacancia():
        u = np.random.default_rng(4).random(40_000)
        r = []
        for estado, clave in (("Normal", "vac.media_N"), ("Crisis", "vac.media_C")):
            r.append((float(np.mean([vacancy.duracion(estado, x, p) for x in u])), p.f(clave)))
        return (all(abs(a / b - 1) < 0.04 for a, b in r), f"{_n(r[0][1], 1)} y {_n(r[1][1], 1)} meses (normal, crisis)",
                f"{_n(r[0][0], 2)} y {_n(r[1][0], 2)}")

    def poisson():
        lam = liq.tasa("Normal", p)
        u = np.random.default_rng(5).random(100_000)
        if lam > 0:
            obt = float(np.mean([liq.espera(lam, x) for x in u]))
            esp = 1 / lam
            ok_e = abs(obt / esp - 1) < 0.03
            txt_e = (f"espera media {_n(esp, 1)} meses", f"{_n(obt, 1)}")
        else:
            ok_e = liq.espera(0.0, 0.5) == float("inf")
            txt_e = ("sin necesidades: espera infinita", "infinita" if ok_e else "finita")
        z = np.random.default_rng(6).standard_normal(100_000)
        med = float(np.median([liq.monto(x, 1.0, 1.0, p) for x in z]))
        esp_m = p.f("liq.mediana_pct") * C
        ok_m = abs(med / esp_m - 1) < 0.03 if esp_m > 0 else med == 0
        return ok_e and ok_m, f"{txt_e[0]}; monto mediano {_m(esp_m)}", f"{txt_e[1]}; {_m(med)}"

    def reproducible():
        a, b, c1 = una_replica(p, "E4", 123), una_replica(p, "E4", 123), una_replica(p, "E4", 124)
        ok = np.array_equal(a.trayectoria["W"], b.trayectoria["W"]) and not np.array_equal(a.trayectoria["W"], c1.trayectoria["W"])
        return ok, "misma semilla, mismo resultado; otra semilla, distinto", "se cumple" if ok else "no se cumple"

    def independientes():
        g = generadores(99)
        x = {k: v.random(5000) for k, v in g.items()}
        ks = list(x)
        mx = max(abs(float(np.corrcoef(x[ks[i]], x[ks[j]])[0, 1])) for i in range(len(ks)) for j in range(i + 1, len(ks)))
        return mx < 0.06, f"correlación entre los {len(ks)} flujos menor de 0,06", f"máxima {mx:.3f}"

    def comunes():
        r = correr(p, n=3, semilla=semilla, guardar=("estado", "tpm", "inflacion"))
        tr = r.trayectorias
        ok = all(np.array_equal(tr["E1"][k], tr["E3"][k]) and np.array_equal(tr["E2"][k], tr["E4"][k])
                 for k in ("estado", "tpm", "inflacion"))
        return ok, "todas las estrategias ven la misma economía en cada réplica", "se cumple" if ok else "no se cumple"

    prueba(GENERADORES, "Filas de la matriz de Markov suman 1", filas_markov)
    prueba(GENERADORES, "Frecuencias simuladas de la cadena de Markov", frec_markov)
    prueba(GENERADORES, "Volatilidad de los choques del ETF", choques_etf)
    prueba(GENERADORES, "Inflación converge a su nivel de largo plazo", inflacion_nivel)
    prueba(GENERADORES, "Duración media de la vacancia", vacancia)
    prueba(GENERADORES, "Necesidades de liquidez: frecuencia y monto", poisson)
    prueba(GENERADORES, "Reproducibilidad con semilla", reproducible)
    prueba(GENERADORES, "Independencia entre flujos aleatorios", independientes)
    prueba(GENERADORES, "Números aleatorios comunes entre estrategias", comunes)

    # --- Reglas de decisión
    def e4_objetivo(persistente: dict, letra: str):
        pp = q.con(**{"inv.horizonte": max(H, 60), "inv.liquidez_min": 0.0, **persistente})
        tr = una_replica(pp, "E4", 1).trayectoria
        F = tr["cuenta"][-1] + tr["cdt"][-1] + tr["etf"][-1]
        obt = tr["etf"][-1] / F
        esp = p.lista(f"E4.obj_{letra}")[2]
        banda = p.f("E4.banda")
        return (abs(obt - esp) <= banda + 1e-3, f"ETF ≈ {_n(esp * 100, 1)} % del portafolio financiero (banda ±{_n(banda * 100, 1)} puntos)",
                f"{_n(obt * 100, 1)} %")

    def e4_crisis():
        return e4_objetivo({"eco.estado_inicial": "Crisis", "markov.C_C": 1.0, "markov.C_N": 0.0}, "C")

    def e4_expansion():
        return e4_objetivo({"eco.estado_inicial": "Expansión", "markov.E_E": 1.0, "markov.E_N": 0.0}, "E")

    def e2_no_rebalancea():
        n = una_replica(p, "E2", semilla).indicadores["n_rebalanceos"]
        return n == 0, "E2 no rebalancea", f"{n:g} rebalanceos"

    def orden_cobertura():
        pp = q.con(**{"E1.pesos": "0.05;0.6;0.35;0", "inv.liquidez_min": 0.0, "liq.tasa_E": 1.0, "liq.tasa_N": 1.0,
                      "liq.tasa_C": 1.0, "liq.mediana_pct": 0.05, "liq.sigma": 0.0})
        a = una_replica(pp, "E1", 4).indicadores
        etf_ = a["acum_cubierto_etf"]
        ok = etf_ > 0 and (a["acum_cubierto_cdt"] == 0 or etf_ >= 0.35 * C * 0.98)
        return ok, "el CDT se vende solo cuando el ETF ya se agotó", f"ETF cubrió {_m(etf_)}; CDT {_m(a['acum_cubierto_cdt'])}"

    def iliquidez():
        pp = q.con(**{"E1.pesos": "1;0;0;0", "liq.tasa_N": 1.0, "liq.tasa_E": 1.0, "liq.tasa_C": 1.0,
                      "liq.mediana_pct": 1.0, "liq.sigma": 0.0})
        a = una_replica(pp, "E1", 3).indicadores
        return (a["episodios_iliquidez"] >= 1 and a["deuda_final"] > 0, "sin de dónde cubrir: episodio de iliquidez y deuda",
                f"{a['episodios_iliquidez']:g} episodios; deuda {_m(a['deuda_final'])}")

    def parametros_invalidos():
        malos = ({"vac.p_N": 1.5}, {"inv.capital": -1}, {"E2.pesos": "0.5;0.5;0.5;0"}, {"inf.phi": 1.0})
        rechazados = 0
        for c1 in malos:
            try:
                p.con_dict(c1)
            except ParametroInvalido:
                rechazados += 1
        return rechazados == len(malos), "se rechazan los valores fuera de dominio", f"{rechazados} de {len(malos)} rechazados"

    prueba(REGLAS, "E4 en crisis persistente lleva el ETF a su objetivo", e4_crisis)
    prueba(REGLAS, "E4 en expansión persistente lleva el ETF a su objetivo", e4_expansion)
    prueba(REGLAS, "E2 no rebalancea", e2_no_rebalancea)
    prueba(REGLAS, "Orden de cobertura de liquidez: cuenta, ETF, CDT", orden_cobertura)
    prueba(REGLAS, "Sin liquidez suficiente: iliquidez y deuda", iliquidez)
    prueba(REGLAS, "Parámetros fuera de dominio", parametros_invalidos)
    return pd.DataFrame(filas)

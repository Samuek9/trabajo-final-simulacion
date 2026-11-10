"""Procesador de eventos de una réplica (puntos 9 y 10).

Una réplica = una estrategia, un escenario y una semilla. El reloj salta de evento
en evento; cada evento actualiza variables de estado, sortea sus números aleatorios
con su propio generador y programa eventos futuros, como en el diagrama.

Reglas que el diagrama deja abiertas y que se resuelven aquí (ver docs/decision_log.md):
- La inflación de cada diciembre se sortea al empezar el año (t = 0, 12, 24, ...),
  porque la trayectoria mensual es una recta entre diciembres y la Junta la usa durante el año.
- El arriendo que vence el mismo mes en que cambia el contrato se paga con el canon anterior.
- El 4x1000 se cobra sobre los débitos de la cuenta que pasan la exención mensual y sobre
  el producto de las ventas de ETF.
- Un faltante que ninguna fuente cubre queda como deuda (obligación sin pagar) y abre un
  episodio de iliquidez; se paga con la primera entrada de dinero.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Any

import numpy as np

from ..models import equity, inflation, interest_rates as tasas, liquidity_shocks as liq
from ..models import markov, real_estate, vacancy
from ..strategies.rules import Estrategia
from .events import Evento, ListaEventos, Tipo
from .state import CDT, Estado

FLUJOS = ("estado", "inflacion", "junta", "cdt", "etf", "valorizacion", "vacancia", "canon",
          "reparaciones", "liquidez_tiempo", "liquidez_monto", "crisis", "venta")
TOL = 1.0  # pesos: diferencias menores se consideran cero


def generadores(semilla: int) -> dict[str, np.random.Generator]:
    """Un generador independiente por proceso, derivado de la semilla de la réplica."""
    hijos = np.random.SeedSequence(semilla).spawn(len(FLUJOS))
    return {n: np.random.default_rng(s) for n, s in zip(FLUJOS, hijos)}


@dataclass
class ResultadoReplica:
    estrategia: str
    semilla: int
    indicadores: dict[str, float]
    trayectoria: dict[str, np.ndarray] = field(default_factory=dict)
    traza: list[dict[str, Any]] = field(default_factory=list)


class Simulacion:
    def __init__(self, p, estrategia: Estrategia, semilla: int, residuos_cdt: np.ndarray | None = None,
                 p_base=None, traza: bool = False) -> None:
        self.p = p
        # Fuera de la crisis forzada el escenario de estrés usa los valores base.
        self.p_fuera = p_base if p_base is not None else p
        self.e = estrategia
        self.semilla = semilla
        self.res_cdt = residuos_cdt
        self.con_traza = traza
        self.H = p.i("inv.horizonte")
        self.ret = p.f("imp.retencion")
        self.ret_cta = p.f("imp.retencion_cuenta")
        self.gmf = p.f("imp.gmf")
        self.exencion = p.f("imp.gmf_exencion")
        self.com_etf = p.f("etf.comision")
        self.liq_min = p.f("inv.liquidez_min")
        self.meta = p.f("inv.meta_real")
        self.capital = p.f("inv.capital")
        self.meses_junta = [int(m) for m in p.lista("tpm.meses_junta")]
        self.rng = generadores(semilla)
        anios = self.H // 12 + 3
        g = self.rng
        self.u_estado = g["estado"].random(self.H // 3 + 3)
        self.z_inf = g["inflacion"].standard_normal(anios)
        self.u_junta = g["junta"].random((8 * anios, 2))
        self.z_cdt = g["cdt"].standard_normal(anios + 2)
        self.u_cdt = g["cdt"].random(anios + 2)
        self.choque_etf = equity.choques(g["etf"].random(self.H + 2), p)
        self.z_val = g["valorizacion"].standard_normal(anios + 2)
        self.u_crisis = g["crisis"].random(2)
        self.P = markov.matriz(p)
        self.P_fuera = markov.matriz(self.p_fuera)
        self.cola = ListaEventos()
        self.s = Estado()
        self.ev_necesidad: Evento | None = None
        self.junta_k = 0
        self.compra_etf_pendiente = 0.0
        self.renovar_cdt_en: float | None = None
        self.n = self.H + 1
        self.tray = {k: np.zeros(self.n) for k in (
            "W", "W_real", "cuenta", "cdt", "etf", "inmueble", "deuda", "tpm", "inflacion", "indice")}
        self.tray["estado"] = np.zeros(self.n, dtype=np.int8)
        self.tray["iliquido"] = np.zeros(self.n, dtype=bool)
        self.tray["ocupado"] = np.zeros(self.n, dtype=np.int8)
        self.traza: list[dict[str, Any]] = []

    # ------------------------------------------------------------------ utilidades
    @property
    def pp(self):
        """Parámetros vigentes: los de la crisis forzada solo mientras dura."""
        return self.p if self.s.crisis_forzada else self.p_fuera

    def W(self) -> float:
        return self.s.patrimonio(self.ret)

    def pi_actual(self, t: float) -> float:
        s = self.s
        if t < 0:
            return self.p.f("inf.inicial")
        return inflation.interpolar(s.pi_ini, s.pi_fin, s.t_pi, t)

    def tpm_hace(self, meses: float) -> float:
        objetivo = self.s.t - meses
        valor = self.p.f("tpm.hace_3m") if objetivo < 0 else self.s.historial_tpm[0][1]
        for t, v in self.s.historial_tpm:
            if t <= objetivo + 1e-9:
                valor = v
            else:
                break
        return valor

    def anotar(self, ev_tipo: Tipo, texto: str, **extra: Any) -> None:
        if self.con_traza:
            self.traza.append({"t": round(self.s.t, 4), "n": ev_tipo.numero, "evento": ev_tipo.nombre,
                               "detalle": texto, "estado": self.s.estado_eco,
                               "cuenta": self.s.cuenta, "W": self.W(), **extra})

    def debitar(self, monto: float, concepto: str) -> None:
        """Saca dinero de la cuenta y cobra el 4x1000 sobre lo que pase la exención del mes."""
        if monto <= 0:
            return
        s = self.s
        antes = max(s.debitos_mes - self.exencion, 0.0)
        s.debitos_mes += monto
        despues = max(s.debitos_mes - self.exencion, 0.0)
        impuesto = (despues - antes) * self.gmf
        s.cuenta -= monto + impuesto
        s.acum["gmf"] += impuesto

    def costo_debito(self, monto: float) -> float:
        antes = max(self.s.debitos_mes - self.exencion, 0.0)
        despues = max(self.s.debitos_mes + monto - self.exencion, 0.0)
        return monto + (despues - antes) * self.gmf

    def acreditar(self, monto: float) -> None:
        self.s.cuenta += monto
        self.pagar_deuda()

    def pagar_deuda(self) -> None:
        s = self.s
        if s.deuda > TOL and s.cuenta > TOL:
            pago = min(s.deuda, s.cuenta / (1 + self.gmf))
            self.debitar(pago, "deuda")
            s.deuda -= pago

    def pagar(self, monto: float, concepto: str, mantener_minimo: bool = False) -> None:
        """Paga una obligación. Si la cuenta no alcanza (o queda bajo el mínimo cuando
        mantener_minimo es verdadero) pide cobertura de liquidez (evento 12) por el faltante.
        Lo que no se cubre queda como deuda."""
        if monto <= 0:
            return
        s = self.s
        necesario = self.costo_debito(monto)
        piso = self.liq_min * max(self.W() - monto, 0.0) if mantener_minimo else 0.0
        if s.cuenta - necesario < piso - TOL:
            self.cobertura(necesario + piso - s.cuenta, concepto)
        disponible = max(s.cuenta, 0.0)
        if disponible + TOL >= self.costo_debito(monto):
            self.debitar(monto, concepto)
        else:
            pagado = disponible / (1 + self.gmf) if disponible > 0 else 0.0
            pagado = min(pagado, monto)
            self.debitar(pagado, concepto)
            s.deuda += monto - pagado
        if s.cuenta < 0 and abs(s.cuenta) < TOL:
            s.cuenta = 0.0

    # ------------------------------------------------------------------ evento 12
    def cobertura(self, faltante: float, motivo: str) -> float:
        """Cobertura de liquidez: ETF, venta anticipada de CDT y, al final, venta urgente
        del inmueble. Devuelve lo que entró a la cuenta de inmediato."""
        s = self.s
        if faltante <= TOL:
            return 0.0
        s.contadores["coberturas"] += 1
        pendiente = faltante
        entro = 0.0
        # (2) ETF: se vende en el mismo mes, con comisión y 4x1000.
        if s.etf > TOL and pendiente > TOL:
            neto_por_peso = (1 - self.com_etf) * (1 - self.gmf)
            bruto = min(s.etf, pendiente / neto_por_peso)
            frac = bruto / s.etf if s.etf > 0 else 0.0
            costo = s.etf_costo * frac
            s.etf -= bruto
            s.etf_costo -= costo
            comision = bruto * self.com_etf
            impuesto = (bruto - comision) * self.gmf
            neto = bruto - comision - impuesto
            s.acum["comision_etf"] += comision
            s.acum["gmf"] += impuesto
            s.acum["perdida_realizada_etf"] += max(costo - (bruto - comision), 0.0)
            s.acum["cubierto_etf"] += neto
            s.contadores["ventas_etf"] += 1
            s.cuenta += neto
            pendiente -= neto
            entro += neto
            if s.etf < TOL:
                s.etf, s.etf_costo = 0.0, 0.0
        # (3) CDT: venta anticipada a precio de mercado más castigo.
        if pendiente > TOL and s.cdts:
            r_mkt = (s.tpm + tasas.margen_esperado(s.tpm - self.tpm_hace(3), self.pp)) / 100
            for cdt in sorted(list(s.cdts), key=lambda c: c.abierto):
                if pendiente <= TOL:
                    break
                tau = (s.t - cdt.abierto) / 12
                plazo = (cdt.vence - cdt.abierto) / 12
                bruto_total = tasas.valor_venta_anticipada(cdt.capital, cdt.tasa, r_mkt, self.p.f("cdt.castigo"), plazo, tau)
                ganancia = max(bruto_total - cdt.capital, 0.0)
                neto_total = bruto_total - ganancia * self.ret
                frac = min(1.0, pendiente / neto_total) if neto_total > 0 else 1.0
                causado = cdt.capital * ((1 + cdt.tasa) ** tau - 1)
                s.acum["intereses_perdidos"] += max(causado - (bruto_total - cdt.capital), 0.0) * frac
                s.acum["retencion"] += ganancia * self.ret * frac
                neto = neto_total * frac
                s.acum["cubierto_cdt"] += neto
                s.contadores["ventas_cdt"] += 1
                s.cuenta += neto
                pendiente -= neto
                entro += neto
                if frac >= 1 - 1e-12:
                    ListaEventos.cancelar(cdt.evento)
                    s.cdts.remove(cdt)
                else:
                    cdt.capital *= (1 - frac)
        # (4) Inmueble: venta urgente, el dinero llega en 3 a 6 meses.
        inm = s.inmueble
        if pendiente > TOL and inm.tiene and not inm.en_venta:
            inm.en_venta = True
            espera = self.rng["venta"].uniform(self.p.f("inm.venta_min"), self.p.f("inm.venta_max"))
            self.cola.programar(s.t + espera, Tipo.VENTA_INMUEBLE)
            self.anotar(Tipo.VENTA_INMUEBLE, f"Se decide la venta urgente; el dinero llega en {espera:.1f} meses")
        self.anotar(Tipo.CIERRE_MES, f"Cobertura de liquidez ({motivo}): faltante {faltante:,.0f}, entró {entro:,.0f}")
        return entro

    # ------------------------------------------------------------------ eventos
    def iniciar(self) -> None:
        p, s = self.p, self.s
        s.t = 0.0
        s.estado_eco = p.s("eco.estado_inicial")
        s.tpm = p.f("tpm.inicial")
        s.historial_tpm = [(0.0, s.tpm)]
        s.pi_dic = p.f("inf.inicial")
        s.pi_ini = s.pi_dic
        s.t_pi = 0.0
        s.pi_fin = inflation.siguiente_diciembre(s.pi_dic, self.z_inf[0], p.f("inf.c"), p.f("inf.phi"), p.f("inf.sigma"))
        s.indice = 1.0
        s.cuenta = self.capital
        w = self.e.pesos_iniciales(p)
        # Inmueble: proporcional al capital o de valor fijo.
        if w["inmueble"] > 0:
            if p.s("inv.modo_inmueble") == "fijo":
                V = p.f("inv.valor_inmueble_fijo")
                resto = self.capital - V
                suma = w["cuenta"] + w["cdt"] + w["etf"]
                base = {k: resto * w[k] / suma for k in ("cuenta", "cdt", "etf")}
            else:
                V = w["inmueble"] * self.capital
                base = {k: w[k] * self.capital for k in ("cuenta", "cdt", "etf")}
            costo = V * p.f("inm.costo_compra")
            self.debitar(V + costo, "compra del inmueble")
            s.acum["costo_compra"] += costo
            inm = s.inmueble
            inm.tiene, inm.valor, inm.valor_compra, inm.ocupado = True, V, V, True
            inm.canon = p.f("inm.canon_inicial") * V
            inm.canon_anterior = inm.canon
            inm.contrato_hasta = 12.0
            inm.eventos["arriendo"] = self.cola.programar(1.0, Tipo.ARRIENDO)
            inm.eventos["contrato"] = self.cola.programar(12.0, Tipo.FIN_CONTRATO)
            inm.eventos["valorizacion"] = self.cola.programar(12.0, Tipo.VALORIZACION, k=1)
            self._programar_reparacion()
        else:
            base = {k: w[k] * self.capital for k in ("cuenta", "cdt", "etf")}
        if base["etf"] > 0:
            self.comprar_etf(base["etf"])
        if base["cdt"] > 0:
            self.abrir_cdt(base["cdt"], k=0)
        # Programa los eventos iniciales.
        self.cola.programar(1.0, Tipo.CIERRE_MES)
        self.cola.programar(float(p.i("eco.paso")), Tipo.CAMBIO_ESTADO, k=0)
        self.cola.programar(12.0, Tipo.INFLACION, k=1)
        self._programar_junta(0.0)
        self._programar_necesidad()
        if self.e.dinamica:
            self.cola.programar(float(p.i("E4.cada")), Tipo.REVISION)
        if p.i("estres.activo") == 1:
            ini = p.f("estres.inicio_min") + self.u_crisis[0] * (self.H - p.f("estres.margen_final") - p.f("estres.inicio_min"))
            dur = p.f("estres.dur_min") + self.u_crisis[1] * (p.f("estres.dur_max") - p.f("estres.dur_min"))
            self.cola.programar(ini, Tipo.CRISIS_INICIO, duracion=dur)
            self.cola.programar(ini + dur, Tipo.CRISIS_FIN)
        self.cola.programar(float(self.H), Tipo.FIN)
        s.pico_real = self.W() / s.indice
        self._registrar(0)
        self.anotar(Tipo.INICIO, f"Compra con los pesos de {self.e.nombre}; TPM {s.tpm:.2f} %; inflación {s.pi_dic:.2f} %")

    def comprar_etf(self, monto: float) -> None:
        """Gasta 'monto' de la cuenta en ETF; la comisión sale del monto."""
        s = self.s
        monto = min(monto, max(s.cuenta, 0.0) / (1 + self.gmf))
        if monto <= TOL:
            return
        self.debitar(monto, "compra de ETF")
        comision = monto * self.com_etf
        s.acum["comision_etf"] += comision
        s.etf += monto - comision
        s.etf_costo += monto

    def abrir_cdt(self, monto: float, k: int) -> None:
        s, p = self.s, self.p
        monto = min(monto, max(s.cuenta, 0.0) / (1 + self.gmf))
        if monto <= TOL:
            return
        delta = s.tpm - self.tpm_hace(3)
        margen = tasas.margen_cdt(delta, self.z_cdt[k], self.u_cdt[k], self.pp, self.res_cdt)
        tasa = (s.tpm + margen) / 100
        self.debitar(monto, "apertura de CDT")
        plazo = p.f("cdt.plazo")
        cdt = CDT(monto, tasa, s.t, s.t + plazo)
        if s.t + plazo <= self.H + 1e-9:
            cdt.evento = self.cola.programar(s.t + plazo, Tipo.VENCE_CDT, cdt=cdt, k=k + 1)
        s.cdts.append(cdt)

    def _programar_junta(self, desde: float) -> None:
        anio = int(desde // 12)
        mes_actual = desde - 12 * anio
        for a in (anio, anio + 1):
            for m in self.meses_junta:
                t = 12 * a + m
                if t > desde + 1e-9 and (a > anio or m > mes_actual + 1e-9):
                    if t <= self.H:
                        self.cola.programar(float(t), Tipo.JUNTA)
                    return

    def _programar_necesidad(self) -> None:
        lam = liq.tasa(self.s.estado_eco, self.pp)
        espera = liq.espera(lam, self.rng["liquidez_tiempo"].random())
        ListaEventos.cancelar(self.ev_necesidad)
        self.ev_necesidad = None
        if self.s.t + espera <= self.H:
            self.ev_necesidad = self.cola.programar(self.s.t + espera, Tipo.NECESIDAD, lam=lam)

    def _programar_reparacion(self) -> None:
        lam = self.p.f("rep.tasa")
        if lam <= 0:
            return
        espera = liq.espera(lam, self.rng["reparaciones"].random())
        if self.s.t + espera <= self.H:
            self.s.inmueble.eventos["reparacion"] = self.cola.programar(self.s.t + espera, Tipo.REPARACION)

    def e_cambio_estado(self, ev: Evento) -> None:
        s = self.s
        k = ev.datos["k"]
        u = self.u_estado[k]
        antes = s.estado_eco
        lam_antes = liq.tasa(antes, self.pp)
        if not s.crisis_forzada:
            s.estado_eco = markov.siguiente(antes, u, self.P_fuera)
        if liq.tasa(s.estado_eco, self.pp) != lam_antes:
            self._programar_necesidad()
        self.anotar(ev.tipo, f"{antes} → {s.estado_eco}" + (" (crisis forzada: la matriz no opera)" if s.crisis_forzada else ""))
        siguiente = s.t + self.p.i("eco.paso")
        if siguiente <= self.H:
            self.cola.programar(siguiente, Tipo.CAMBIO_ESTADO, k=k + 1)

    def e_crisis_inicio(self, ev: Evento) -> None:
        s = self.s
        s.crisis_forzada = True
        s.estado_eco = "Crisis"
        caida = self.p.f("etf.caida_estres")
        if s.etf > 0 and caida > 0:
            s.etf *= (1 - caida)
        self._programar_necesidad()
        self.anotar(ev.tipo, f"Empieza una crisis de {ev.datos['duracion']:.1f} meses; el ETF cae {caida:.0%}")

    def e_crisis_fin(self, ev: Evento) -> None:
        s = self.s
        s.crisis_forzada = False
        s.estado_eco = "Normal"
        self._programar_necesidad()
        self.anotar(ev.tipo, "Termina la crisis forzada; la economía pasa a Normal")

    def e_inflacion(self, ev: Evento) -> None:
        s, p = self.s, self.p
        k = ev.datos["k"]
        s.pi_dic = s.pi_fin
        s.pi_ini = s.pi_fin
        s.t_pi = s.t
        s.pi_fin = inflation.siguiente_diciembre(s.pi_dic, self.z_inf[k], p.f("inf.c"), p.f("inf.phi"), p.f("inf.sigma"))
        self.anotar(ev.tipo, f"Inflación de diciembre {s.pi_dic:.2f} %; la del próximo será {s.pi_fin:.2f} %")
        if s.t + 12 <= self.H:
            self.cola.programar(s.t + 12, Tipo.INFLACION, k=k + 1)

    def e_junta(self, ev: Evento) -> None:
        s, p = self.s, self.p
        cambio_inf = self.pi_actual(s.t) - self.pi_actual(s.t - 3)
        tend = tasas.tendencia(cambio_inf, p.f("tpm.umbral_tendencia"))
        u1, u2 = self.u_junta[self.junta_k]
        self.junta_k += 1
        if p.s("junta.modelo") == "logit":
            delta = tasas.decidir_logit(cambio_inf, s.tpm, self.pi_actual(s.t), u1, u2, p, bajadas_estres=s.crisis_forzada)
        else:
            delta = tasas.decidir(tend, u1, u2, p, bajadas_estres=s.crisis_forzada)
        antes = s.tpm
        s.tpm = tasas.limitar_tpm(s.tpm + delta, p)
        s.historial_tpm.append((s.t, s.tpm))
        self.anotar(ev.tipo, f"Inflación {tend.lower()}: TPM {antes:.2f} % → {s.tpm:.2f} %")
        self._programar_junta(s.t)

    def e_valorizacion(self, ev: Evento) -> None:
        s, p = self.s, self.p
        inm = s.inmueble
        if not inm.tiene:
            return
        k = ev.datos["k"]
        r = real_estate.valorizacion_log(s.estado_eco, self.z_val[k], inm.r_anterior, self.pp)
        inm.r_anterior = r
        inm.valor *= math.exp(r)
        predial = p.f("inm.predial") * inm.valor
        mant = p.f("inm.mantenimiento") * inm.valor
        s.acum["predial"] += predial
        s.acum["mantenimiento"] += mant
        self.pagar(predial + mant, "predial y mantenimiento")
        self.anotar(ev.tipo, f"El inmueble cambia {math.exp(r) - 1:+.2%}; predial y mantenimiento {predial + mant:,.0f}")
        if s.t + 12 <= self.H:
            inm.eventos["valorizacion"] = self.cola.programar(s.t + 12, Tipo.VALORIZACION, k=k + 1)

    def e_revision(self, ev: Evento) -> None:
        """E4: pesos objetivo del portafolio financiero según el estado; banda de 5 puntos."""
        s, p = self.s, self.p
        obj = self.e.objetivo(s.estado_eco, p)
        banda = p.f("E4.banda")
        cdt_val = s.valor_cdts(self.ret)
        F = s.cuenta + cdt_val + s.etf
        if F <= TOL:
            return
        hechos = []
        meta_etf = obj["etf"] * F
        if abs(s.etf / F - obj["etf"]) > banda:
            if s.etf > meta_etf:
                exceso = s.etf - meta_etf
                frac = exceso / s.etf
                costo = s.etf_costo * frac
                s.etf -= exceso
                s.etf_costo -= costo
                comision = exceso * self.com_etf
                impuesto = (exceso - comision) * self.gmf
                s.acum["comision_etf"] += comision
                s.acum["gmf"] += impuesto
                s.acum["perdida_realizada_etf"] += max(costo - (exceso - comision), 0.0)
                self.acreditar(exceso - comision - impuesto)
                hechos.append(f"vende ETF por {exceso:,.0f}")
            else:
                self.compra_etf_pendiente = meta_etf - s.etf
                hechos.append(f"comprará ETF por {self.compra_etf_pendiente:,.0f}")
        meta_cdt = obj["cdt"] * F
        if abs(cdt_val / F - obj["cdt"]) > banda:
            self.renovar_cdt_en = meta_cdt
            hechos.append(f"renovará CDT por {meta_cdt:,.0f}")
        else:
            self.renovar_cdt_en = None
        if hechos:
            s.contadores["rebalanceos"] += 1
        # Si este mes no vence ningún CDT, la compra de ETF se hace ya.
        if not any(abs(c.vence - s.t) < 1e-9 for c in s.cdts):
            self._ejecutar_compra_pendiente()
            if self.renovar_cdt_en is not None and self.renovar_cdt_en > cdt_val + TOL:
                self.abrir_cdt(min(self.renovar_cdt_en - cdt_val, self._disponible()), k=int(round(s.t / 12)))
            self.renovar_cdt_en = None
        self.anotar(ev.tipo, f"Estado {s.estado_eco}: objetivo cuenta {obj['cuenta']:.1%}, CDT {obj['cdt']:.1%}, ETF {obj['etf']:.1%}; " + ("; ".join(hechos) or "sin ajustes"))
        if s.t + p.i("E4.cada") <= self.H:
            self.cola.programar(s.t + p.i("E4.cada"), Tipo.REVISION)

    def _disponible(self) -> float:
        """Lo que se puede sacar de la cuenta sin bajar de la liquidez mínima."""
        return max(self.s.cuenta - self.liq_min * self.W(), 0.0) / (1 + self.gmf)

    def _ejecutar_compra_pendiente(self) -> None:
        if self.compra_etf_pendiente > TOL:
            self.comprar_etf(min(self.compra_etf_pendiente, self._disponible()))
        self.compra_etf_pendiente = 0.0

    def e_vence_cdt(self, ev: Evento) -> None:
        s = self.s
        cdt: CDT = ev.datos["cdt"]
        if cdt not in s.cdts:
            return
        interes = cdt.capital * ((1 + cdt.tasa) ** ((cdt.vence - cdt.abierto) / 12) - 1)
        s.acum["intereses_cdt"] += interes * (1 - self.ret)
        s.acum["retencion"] += interes * self.ret
        s.cdts.remove(cdt)
        self.acreditar(cdt.capital + interes * (1 - self.ret))
        monto = cdt.capital
        if self.e.dinamica and self.renovar_cdt_en is not None:
            monto = self.renovar_cdt_en
            self.renovar_cdt_en = None
        if self.e.dinamica:
            monto = min(monto, cdt.capital + self._disponible())
        k = ev.datos["k"]
        self.abrir_cdt(monto, k=k)
        nuevo = s.cdts[-1] if s.cdts and abs(s.cdts[-1].abierto - s.t) < 1e-9 else None
        self.anotar(ev.tipo, f"Vence CDT de {cdt.capital:,.0f} al {cdt.tasa:.2%}; " +
                    (f"se renueva por {nuevo.capital:,.0f} al {nuevo.tasa:.2%}" if nuevo else "no se renueva"))
        if self.e.dinamica:
            self._ejecutar_compra_pendiente()

    def e_fin_contrato(self, ev: Evento) -> None:
        s, p = self.s, self.p
        inm = s.inmueble
        if not inm.tiene:
            return
        u, u2 = self.rng["vacancia"].random(2)
        z = self.rng["canon"].standard_normal()
        inm.canon_anterior = inm.canon
        inm.t_cambio_canon = s.t
        if vacancy.no_renueva(s.estado_eco, u, self.pp):
            d = vacancy.duracion(s.estado_eco, u2, self.pp)
            inm.ocupado = False
            inm.contrato_hasta = s.t
            inm.vacancia_desde = s.t
            s.contadores["vacancias"] += 1
            pintura = p.f("inm.pintura") * s.indice
            s.acum["pintura"] += pintura
            self.pagar(pintura, "pintura")
            inm.eventos["vacancia"] = self.cola.programar(s.t + d, Tipo.FIN_VACANCIA)
            self.anotar(ev.tipo, f"El inquilino no renueva; vacancia de {d:.1f} meses")
        else:
            antes = inm.canon
            inm.canon = real_estate.reajuste_canon(inm.canon, inm.valor, s.pi_dic, z, p)
            inm.contrato_hasta = s.t + 12
            s.contadores["renovaciones"] += 1
            if s.t + 12 <= self.H:
                inm.eventos["contrato"] = self.cola.programar(s.t + 12, Tipo.FIN_CONTRATO)
            self.anotar(ev.tipo, f"Renueva: canon {antes:,.0f} → {inm.canon:,.0f} (IPC previo {s.pi_dic:.2f} %)")

    def e_fin_vacancia(self, ev: Evento) -> None:
        s, p = self.s, self.p
        inm = s.inmueble
        if not inm.tiene:
            return
        s.acum["meses_vacante"] += s.t - (inm.vacancia_desde or s.t)
        inm.vacancia_desde = None
        inm.ocupado = True
        inm.canon = p.f("inm.canon_inicial") * inm.valor
        inm.canon_anterior = 0.0
        inm.t_cambio_canon = s.t
        inm.contrato_hasta = s.t + 12
        if s.t + 1 <= self.H:
            inm.eventos["arriendo"] = self.cola.programar(s.t + 1, Tipo.ARRIENDO)
        if s.t + 12 <= self.H:
            inm.eventos["contrato"] = self.cola.programar(s.t + 12, Tipo.FIN_CONTRATO)
        self.anotar(ev.tipo, f"Nuevo inquilino; canon {inm.canon:,.0f}")

    def e_arriendo(self, ev: Evento) -> None:
        s, p = self.s, self.p
        inm = s.inmueble
        if not inm.tiene:
            return
        # El mes que termina en t se paga con el canon que regía en ese mes: si el
        # contrato cambió justo en t, con el anterior.
        if abs(inm.t_cambio_canon - s.t) < 1e-9:
            canon = inm.canon_anterior
        elif inm.ocupado:
            canon = inm.canon
        else:
            canon = 0.0
        if canon > 0:
            comision = canon * p.f("inm.comision_inmobiliaria")
            s.acum["arriendo_bruto"] += canon
            s.acum["comision_inmobiliaria"] += comision
            self.acreditar(canon - comision)
        if inm.ocupado and s.t + 1 <= self.H:
            inm.eventos["arriendo"] = self.cola.programar(s.t + 1, Tipo.ARRIENDO)

    def e_reparacion(self, ev: Evento) -> None:
        s = self.s
        if not s.inmueble.tiene:
            return
        costo = liq.costo_reparacion(self.rng["reparaciones"].standard_normal(), s.indice, self.p)
        s.acum["reparaciones"] += costo
        s.contadores["reparaciones"] += 1
        self.pagar(costo, "reparación")
        self.anotar(ev.tipo, f"Reparación por {costo:,.0f}")
        self._programar_reparacion()

    def e_necesidad(self, ev: Evento) -> None:
        s = self.s
        z = self.rng["liquidez_monto"].standard_normal()
        M = liq.monto(z, s.indice, self.pp.f("liq.mult_monto"), self.p)
        s.acum["necesidades"] += M
        s.contadores["necesidades"] += 1
        cuenta_antes = s.cuenta
        otras = s.acum["cubierto_etf"] + s.acum["cubierto_cdt"] + s.deuda
        self.pagar(M, "necesidad de liquidez", mantener_minimo=True)
        de_otras = s.acum["cubierto_etf"] + s.acum["cubierto_cdt"] + s.deuda - otras
        s.acum["cubierto_cuenta"] += max(M - de_otras, 0.0)
        self.anotar(ev.tipo, f"Necesidad de {M:,.0f}; la cuenta pasa de {cuenta_antes:,.0f} a {s.cuenta:,.0f}")
        self.ev_necesidad = None
        self._programar_necesidad()

    def e_venta_inmueble(self, ev: Evento) -> None:
        s, p = self.s, self.p
        inm = s.inmueble
        if not inm.tiene:
            return
        precio = inm.valor * (1 - p.f("inm.descuento_urgente"))
        costos = precio * p.f("inm.costo_venta")
        s.acum["descuento_urgente"] += inm.valor - precio
        s.acum["costo_venta"] += costos
        s.acum["cubierto_inmueble"] += precio - costos
        s.contadores["ventas_inmueble"] += 1
        if inm.vacancia_desde is not None:
            s.acum["meses_vacante"] += s.t - inm.vacancia_desde
        for e in inm.eventos.values():
            ListaEventos.cancelar(e)
        valor = inm.valor
        inm.tiene, inm.valor, inm.ocupado, inm.en_venta, inm.vacancia_desde = False, 0.0, False, False, None
        self.acreditar(precio - costos)
        self.anotar(ev.tipo, f"Se vende el inmueble de {valor:,.0f}: llegan {precio - costos:,.0f}")

    def e_cierre_mes(self, ev: Evento) -> None:
        s, p = self.s, self.p
        m = int(round(s.t))
        # ETF: rendimiento total del mes (dividendos incluidos).
        g = equity.rendimiento_log(s.estado_eco, self.choque_etf[m], self.pp)
        s.etf *= math.exp(g)
        # Cuenta: interés mensual neto de retención.
        if s.cuenta > 0:
            interes = s.cuenta * ((1 + p.f("cta.alfa") * s.tpm / 100) ** (1 / 12) - 1)
            s.acum["intereses_cuenta"] += interes * (1 - self.ret_cta)
            s.acum["retencion"] += interes * self.ret_cta
            s.cuenta += interes * (1 - self.ret_cta)
        # Inflación del mes e índice de precios.
        pi = self.pi_actual(s.t)
        s.indice *= (1 + pi / 100) ** (1 / 12)
        # Gastos fijos del inmueble.
        if s.inmueble.tiene:
            s.acum["meses_con_inmueble"] += 1
            ph = p.f("inm.ph") * s.indice
            s.acum["ph"] += ph
            self.pagar(ph, "administración")
        self.pagar_deuda()
        # Regla de liquidez mínima.
        W = self.W()
        if s.cuenta < self.liq_min * W - TOL:
            self.cobertura(self.liq_min * W - s.cuenta, "cuenta bajo el mínimo")
        W = self.W()
        Wr = W / s.indice
        if Wr > s.pico_real:
            s.pico_real = Wr
        if s.pico_real > 0:
            s.caida_max = max(s.caida_max, 1 - Wr / s.pico_real)
        # La meta se mide igual que al final: el inmueble neto de los costos de venta.
        neto_inm = s.inmueble.valor * p.f("inm.costo_venta") if s.inmueble.tiene else 0.0
        if s.t_meta is None and (W - neto_inm) / s.indice >= self.meta:
            s.t_meta = s.t
        iliquido = s.deuda > TOL or s.cuenta < self.liq_min * W - 1000.0
        if iliquido:
            if not s.en_episodio:
                s.en_episodio, s.duracion_actual = True, 0
            s.duracion_actual += 1
        elif s.en_episodio:
            s.episodios.append(s.duracion_actual)
            s.en_episodio, s.duracion_actual = False, 0
        self._registrar(m, iliquido)
        s.debitos_mes = 0.0
        if s.t + 1 <= self.H:
            self.cola.programar(s.t + 1, Tipo.CIERRE_MES)

    def _registrar(self, m: int, iliquido: bool = False) -> None:
        s, tr = self.s, self.tray
        W = self.W()
        tr["W"][m] = W
        tr["W_real"][m] = W / s.indice
        tr["cuenta"][m] = s.cuenta
        tr["cdt"][m] = s.valor_cdts(self.ret)
        tr["etf"][m] = s.etf
        tr["inmueble"][m] = s.inmueble.valor if s.inmueble.tiene else 0.0
        tr["deuda"][m] = s.deuda
        tr["tpm"][m] = s.tpm
        tr["inflacion"][m] = self.pi_actual(s.t)
        tr["indice"][m] = s.indice
        tr["estado"][m] = markov.ESTADOS.index(s.estado_eco)
        tr["iliquido"][m] = iliquido
        tr["ocupado"][m] = int(s.inmueble.tiene and s.inmueble.ocupado)

    # ------------------------------------------------------------------ ciclo principal
    def correr(self) -> ResultadoReplica:
        manejadores = {
            Tipo.CRISIS_INICIO: self.e_crisis_inicio, Tipo.CRISIS_FIN: self.e_crisis_fin,
            Tipo.CAMBIO_ESTADO: self.e_cambio_estado, Tipo.INFLACION: self.e_inflacion,
            Tipo.JUNTA: self.e_junta, Tipo.VALORIZACION: self.e_valorizacion,
            Tipo.REVISION: self.e_revision, Tipo.VENCE_CDT: self.e_vence_cdt,
            Tipo.FIN_CONTRATO: self.e_fin_contrato, Tipo.FIN_VACANCIA: self.e_fin_vacancia,
            Tipo.ARRIENDO: self.e_arriendo, Tipo.REPARACION: self.e_reparacion,
            Tipo.NECESIDAD: self.e_necesidad, Tipo.VENTA_INMUEBLE: self.e_venta_inmueble,
            Tipo.CIERRE_MES: self.e_cierre_mes,
        }
        self.iniciar()
        reloj = 0.0
        while True:
            ev = self.cola.siguiente()
            if ev is None:
                raise RuntimeError("La lista de eventos quedó vacía antes del fin")
            if ev.tiempo < reloj - 1e-9:
                raise RuntimeError(f"El reloj retrocedió: {ev.tiempo} < {reloj}")
            reloj = ev.tiempo
            self.s.t = ev.tiempo
            self.s.contadores["eventos"] += 1
            if ev.tipo is Tipo.FIN:
                return self.finalizar()
            if ev.tiempo > self.H + 1e-9:
                raise RuntimeError("Se procesó un evento después del horizonte")
            manejadores[ev.tipo](ev)

    def finalizar(self) -> ResultadoReplica:
        s, p = self.s, self.p
        inm = s.inmueble
        if inm.tiene and inm.vacancia_desde is not None:
            s.acum["meses_vacante"] += s.t - inm.vacancia_desde
        if s.en_episodio:
            s.episodios.append(s.duracion_actual)
        inm_neto = inm.valor * (1 - p.f("inm.costo_venta")) if inm.tiene else 0.0
        cdt = s.valor_cdts(self.ret)
        W = s.cuenta + cdt + s.etf + inm_neto - s.deuda
        Wr = W / s.indice
        anios = self.H / 12
        a = s.acum
        arriendo_neto = a["arriendo_bruto"] - a["comision_inmobiliaria"]
        gastos_inm = a["ph"] + a["predial"] + a["mantenimiento"] + a["pintura"] + a["reparaciones"]
        intereses = a["intereses_cuenta"] + a["intereses_cdt"]
        V0 = inm.valor_compra if inm.valor_compra else 0.0
        if V0 > 0:
            final_inm = inm_neto + (a["cubierto_inmueble"] if not inm.tiene else 0.0)
            total = final_inm + arriendo_neto - gastos_inm
            rent_inm = (total / (V0 * (1 + p.f("inm.costo_compra")))) ** (1 / anios) - 1 if total > 0 else -1.0
        else:
            rent_inm = float("nan")
        ind = {
            "W_nominal": W, "W_real": Wr,
            "rent_nominal": (W / self.capital) ** (1 / anios) - 1 if W > 0 else -1.0,
            "rent_real": (Wr / self.capital) ** (1 / anios) - 1 if Wr > 0 else -1.0,
            "caida_max": s.caida_max,
            "perdida_real": float(Wr < self.capital),
            "meta": float(Wr >= self.meta),
            "t_meta": s.t_meta if s.t_meta is not None else float("nan"),
            "episodios_iliquidez": float(len(s.episodios)),
            "meses_iliquidez": float(sum(s.episodios)),
            "duracion_max_iliquidez": float(max(s.episodios) if s.episodios else 0),
            "duracion_media_iliquidez": float(np.mean(s.episodios)) if s.episodios else float("nan"),
            "algun_episodio": float(bool(s.episodios)),
            "liquido_final": s.cuenta + cdt + s.etf - s.deuda,
            "liquido_final_real": (s.cuenta + cdt + s.etf - s.deuda) / s.indice,
            "deuda_final": s.deuda, "insolvente": float(s.deuda > TOL),
            "inflacion_acumulada": s.indice - 1,
            "arriendo_neto": arriendo_neto, "gastos_inmueble": gastos_inm,
            "participacion_arriendo": arriendo_neto / (arriendo_neto + intereses) if arriendo_neto + intereses > 0 else 0.0,
            "rent_inmueble": rent_inm,
            "tiempo_vacante": a["meses_vacante"] / a["meses_con_inmueble"] if a["meses_con_inmueble"] else float("nan"),
            "perdida_venta_urgente": a["descuento_urgente"] + a["costo_venta"],
            "costos_totales": a["comision_etf"] + a["gmf"] + a["retencion"] + a["intereses_perdidos"] + a["costo_compra"] + a["comision_inmobiliaria"]
                              + a["descuento_urgente"] + a["costo_venta"],
            "estado_final": float(markov.ESTADOS.index(s.estado_eco)),
            "tpm_final": s.tpm, "inflacion_final": self.pi_actual(s.t),
        }
        for k, v in a.items():
            ind[f"acum_{k}"] = v
        for k, v in s.contadores.items():
            ind[f"n_{k}"] = float(v)
        return ResultadoReplica(self.e.codigo, self.semilla, ind, self.tray, self.traza)

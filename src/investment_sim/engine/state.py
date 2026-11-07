"""Variables de estado de una réplica (punto 9.2)."""
from __future__ import annotations

from dataclasses import dataclass, field

from .events import Evento


@dataclass
class CDT:
    capital: float
    tasa: float          # fracción anual efectiva
    abierto: float       # mes de apertura
    vence: float         # mes de vencimiento
    evento: Evento | None = None

    def valor(self, t: float, retencion: float) -> float:
        """Capital más intereses causados, netos de retención."""
        anios = max(min(t, self.vence) - self.abierto, 0.0) / 12.0
        interes = self.capital * ((1 + self.tasa) ** anios - 1)
        return self.capital + interes * (1 - retencion)


@dataclass
class Inmueble:
    tiene: bool = False
    valor: float = 0.0
    valor_compra: float = 0.0
    ocupado: bool = False
    canon: float = 0.0
    canon_anterior: float = 0.0
    t_cambio_canon: float = -1.0
    contrato_hasta: float = 0.0
    en_venta: bool = False
    vacancia_desde: float | None = None
    r_anterior: float | None = None
    eventos: dict[str, Evento | None] = field(default_factory=dict)


@dataclass
class Estado:
    t: float = 0.0
    cuenta: float = 0.0
    deuda: float = 0.0
    cdts: list[CDT] = field(default_factory=list)
    etf: float = 0.0
    etf_costo: float = 0.0         # base de costo para pérdidas realizadas
    inmueble: Inmueble = field(default_factory=Inmueble)
    estado_eco: str = "Normal"
    crisis_forzada: bool = False
    tpm: float = 0.0
    historial_tpm: list[tuple[float, float]] = field(default_factory=list)
    pi_ini: float = 0.0            # inflación anual al inicio del tramo
    pi_fin: float = 0.0            # inflación del próximo diciembre
    t_pi: float = 0.0
    pi_dic: float = 0.0            # último diciembre conocido (reajuste del canon)
    indice: float = 1.0            # índice de precios, 1 al inicio
    debitos_mes: float = 0.0       # para la exención del 4x1000
    # Acumulados
    acum: dict[str, float] = field(default_factory=lambda: {k: 0.0 for k in (
        "arriendo_bruto", "comision_inmobiliaria", "ph", "predial", "mantenimiento", "pintura",
        "reparaciones", "costo_compra", "comision_etf", "gmf", "retencion", "intereses_cuenta",
        "intereses_cdt", "intereses_perdidos", "perdida_realizada_etf", "descuento_urgente",
        "costo_venta", "necesidades", "cubierto_cuenta", "cubierto_etf", "cubierto_cdt",
        "cubierto_inmueble", "meses_vacante", "meses_con_inmueble")})
    contadores: dict[str, int] = field(default_factory=lambda: {k: 0 for k in (
        "necesidades", "coberturas", "ventas_etf", "ventas_cdt", "ventas_inmueble", "vacancias",
        "renovaciones", "rebalanceos", "reparaciones", "eventos")})
    # Iliquidez
    en_episodio: bool = False
    duracion_actual: int = 0
    episodios: list[int] = field(default_factory=list)
    # Seguimiento del patrimonio
    pico_real: float = 0.0
    caida_max: float = 0.0
    t_meta: float | None = None

    def valor_cdts(self, retencion: float) -> float:
        return sum(c.valor(self.t, retencion) for c in self.cdts)

    def patrimonio(self, retencion: float) -> float:
        """W = cuenta + CDT + ETF + inmueble − deudas (inmueble a precio de mercado)."""
        inm = self.inmueble.valor if self.inmueble.tiene else 0.0
        return self.cuenta + self.valor_cdts(retencion) + self.etf + inm - self.deuda

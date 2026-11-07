"""Lectura y validación de los parámetros del simulador.

La única fuente de los valores es la hoja PARAMETROS de data/master_datos.xlsx.
Cada fila tiene una clave (por ejemplo "etf.mu"), un valor base y, si cambia,
un valor para el escenario adverso y otro para el de estrés.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

ESCENARIOS = ("base", "adverso", "estres")
ESTADOS_VALIDOS = {"Configurable", "Verificado", "Estimado", "Calibrado", "Supuesto", "Pendiente", "Calculado"}
RAIZ = Path(__file__).resolve().parents[3]
EXCEL_POR_DEFECTO = RAIZ / "data" / "master_datos.xlsx"
COLUMNAS = ("ID", "Clave", "Grupo", "Parámetro", "Base", "Adverso", "Estrés", "Unidad", "Estado", "Fuente", "Soporte")


class ParametroInvalido(ValueError):
    pass


@dataclass(frozen=True)
class Parametro:
    id: str
    clave: str
    grupo: str
    nombre: str
    base: Any
    adverso: Any
    estres: Any
    unidad: str
    estado: str
    fuente: str
    soporte: str

    def valor(self, escenario: str) -> Any:
        if escenario == "adverso" and self.adverso is not None:
            return self.adverso
        if escenario == "estres" and self.estres is not None:
            return self.estres
        return self.base


def leer_tabla(ruta: str | Path = EXCEL_POR_DEFECTO) -> list[Parametro]:
    """Lee la hoja PARAMETROS. Solo valores, nunca fórmulas."""
    from openpyxl import load_workbook

    ruta = Path(ruta)
    if not ruta.exists():
        raise FileNotFoundError(f"No existe el libro de parámetros: {ruta}")
    wb = load_workbook(ruta, read_only=True, data_only=True)
    if "PARAMETROS" not in wb.sheetnames:
        raise ParametroInvalido(f"{ruta.name} no tiene la hoja PARAMETROS")
    ws = wb["PARAMETROS"]
    filas = list(ws.iter_rows(values_only=True))
    wb.close()
    # El encabezado es la primera fila cuya primera celda es "ID".
    inicio = next(i for i, f in enumerate(filas) if f and f[0] == "ID")
    encabezado = [str(c).strip() if c is not None else "" for c in filas[inicio]]
    idx = {c: encabezado.index(c) for c in COLUMNAS}
    salida = []
    for f in filas[inicio + 1:]:
        if not f or f[idx["Clave"]] in (None, ""):
            continue
        g = lambda c: f[idx[c]] if idx[c] < len(f) else None  # noqa: E731
        salida.append(Parametro(
            id=str(g("ID")), clave=str(g("Clave")).strip(), grupo=str(g("Grupo") or ""),
            nombre=str(g("Parámetro") or ""), base=_limpio(g("Base")), adverso=_limpio(g("Adverso")),
            estres=_limpio(g("Estrés")), unidad=str(g("Unidad") or ""), estado=str(g("Estado") or ""),
            fuente=str(g("Fuente") or ""), soporte=str(g("Soporte") or ""),
        ))
    claves = [p.clave for p in salida]
    repetidas = {c for c in claves if claves.count(c) > 1}
    if repetidas:
        raise ParametroInvalido(f"Claves repetidas en PARAMETROS: {sorted(repetidas)}")
    return salida


def _limpio(v: Any) -> Any:
    if isinstance(v, str):
        v = v.strip()
        if v in ("", "—", "-"):
            return None
    return v


@dataclass
class Params:
    """Parámetros de una corrida: escenario + cambios del usuario."""

    valores: dict[str, Any]
    escenario: str = "base"
    tabla: dict[str, Parametro] = field(default_factory=dict)
    cambios: dict[str, Any] = field(default_factory=dict)

    def __getitem__(self, clave: str) -> Any:
        try:
            return self.valores[clave]
        except KeyError as e:
            raise KeyError(f"Parámetro desconocido: {clave}") from e

    def f(self, clave: str) -> float:
        return float(self[clave])

    def i(self, clave: str) -> int:
        return int(round(float(self[clave])))

    def s(self, clave: str) -> str:
        return str(self[clave])

    def lista(self, clave: str) -> list[float]:
        v = self[clave]
        if isinstance(v, (list, tuple)):
            return [float(x) for x in v]
        if isinstance(v, (int, float)):
            return [float(v)]
        return [float(x) for x in str(v).replace(",", ".").split(";") if x.strip() != ""]

    def con(self, **cambios: Any) -> "Params":
        """Copia con cambios. Acepta claves con punto usando dict: p.con(**{"etf.mu": 0})."""
        nuevos = dict(self.valores)
        for k, v in cambios.items():
            if k not in nuevos:
                raise KeyError(f"Parámetro desconocido: {k}")
            nuevos[k] = v
        p = Params(nuevos, self.escenario, self.tabla, {**self.cambios, **cambios})
        validar(p)
        return p

    def con_dict(self, cambios: dict[str, Any]) -> "Params":
        return self.con(**cambios)

    def fuera_de_crisis(self) -> "Params":
        """Valores base con los mismos cambios del usuario.

        En el escenario de estrés, los valores de la columna Estrés rigen solo mientras
        dura la crisis forzada; el resto del tiempo rigen estos.
        """
        if self.escenario != "estres":
            return self
        base = {k: t.base for k, t in self.tabla.items()}
        base.update(self.cambios)
        base["estres.activo"] = self.valores["estres.activo"]
        return Params(base, "base", self.tabla, dict(self.cambios))


def cargar(escenario: str = "base", cambios: dict[str, Any] | None = None,
           ruta: str | Path = EXCEL_POR_DEFECTO, tabla: Iterable[Parametro] | None = None) -> Params:
    if escenario not in ESCENARIOS:
        raise ParametroInvalido(f"Escenario desconocido: {escenario}")
    tabla = list(tabla) if tabla is not None else leer_tabla(ruta)
    valores = {p.clave: p.valor(escenario) for p in tabla}
    p = Params(valores, escenario, {t.clave: t for t in tabla})
    validar(p)
    return p.con_dict(cambios) if cambios else p


def _prob(x: float, nombre: str) -> None:
    if not (0.0 <= x <= 1.0):
        raise ParametroInvalido(f"{nombre} debe estar entre 0 y 1 (vale {x})")


def validar(p: Params) -> None:
    """Rechaza combinaciones imposibles antes de simular."""
    if p.f("inv.capital") <= 0:
        raise ParametroInvalido("El capital inicial debe ser positivo")
    if p.i("inv.horizonte") < 12:
        raise ParametroInvalido("El horizonte debe ser de al menos 12 meses")
    _prob(p.f("inv.liquidez_min"), "La liquidez mínima")
    if p.s("inv.modo_inmueble") not in ("proporcional", "fijo"):
        raise ParametroInvalido("inv.modo_inmueble debe ser 'proporcional' o 'fijo'")
    for o in "ENC":
        fila = [p.f(f"markov.{o}_{d}") for d in "ENC"]
        for x in fila:
            _prob(x, f"Una probabilidad de transición desde {o}")
        if abs(sum(fila) - 1) > 1e-6:
            raise ParametroInvalido(f"La fila {o} de la matriz de transición suma {sum(fila):.4f}, no 1")
    for k in ("junta.bajando", "junta.estable", "junta.subiendo", "junta.p_subidas", "junta.p_bajadas"):
        v = p.lista(k)
        if any(x < 0 for x in v) or abs(sum(v) - 1) > 1e-3:
            raise ParametroInvalido(f"{k} debe tener probabilidades que sumen 1")
    if len(p.lista("junta.subidas")) != len(p.lista("junta.p_subidas")):
        raise ParametroInvalido("junta.subidas y junta.p_subidas deben tener el mismo largo")
    if len(p.lista("junta.bajadas")) != len(p.lista("junta.p_bajadas")):
        raise ParametroInvalido("junta.bajadas y junta.p_bajadas deben tener el mismo largo")
    if not (0 <= p.f("inf.phi") < 1):
        raise ParametroInvalido("La memoria de la inflación φ debe estar en [0, 1)")
    for k in ("inf.sigma", "etf.vol_E", "etf.vol_N", "etf.vol_C", "val.sd_E", "val.sd_N", "val.sd_C"):
        if p.f(k) < 0:
            raise ParametroInvalido(f"{k} no puede ser negativo")
    for k in ("vac.forma", "vac.media_E", "vac.media_N", "vac.media_C"):
        if p.f(k) <= 0:
            raise ParametroInvalido(f"{k} debe ser positivo")
    if p.f("etf.nu") <= 2:
        raise ParametroInvalido("Los grados de libertad de la t deben ser mayores que 2")
    for k in ("vac.p_E", "vac.p_N", "vac.p_C", "cta.alfa", "etf.caida_estres", "inm.descuento_urgente"):
        _prob(p.f(k), k)
    for k in ("liq.tasa_E", "liq.tasa_N", "liq.tasa_C", "rep.tasa"):
        if p.f(k) < 0:
            raise ParametroInvalido(f"{k} no puede ser negativa")
    if p.f("rep.tasa") > 0 and p.f("rep.mediana") <= 0:
        raise ParametroInvalido("Si hay reparaciones, su costo mediano debe ser positivo")
    if p.f("inm.venta_min") > p.f("inm.venta_max"):
        raise ParametroInvalido("El tiempo mínimo de venta no puede superar el máximo")
    if p.f("estres.dur_min") > p.f("estres.dur_max"):
        raise ParametroInvalido("La duración mínima de la crisis no puede superar la máxima")
    if p.f("tpm.min") >= p.f("tpm.max"):
        raise ParametroInvalido("La TPM mínima debe ser menor que la máxima")
    for e in ("E1", "E2", "E3", "E4"):
        w = p.lista(f"{e}.pesos")
        if len(w) != 4 or any(x < 0 for x in w) or abs(sum(w) - 1) > 1e-6:
            raise ParametroInvalido(f"Los pesos de {e} deben ser 4 valores no negativos que sumen 100 %")
    for s in "ENC":
        w = p.lista(f"E4.obj_{s}")
        if len(w) != 3 or any(x < 0 for x in w) or abs(sum(w) - 1) > 1e-6:
            raise ParametroInvalido(f"Los pesos objetivo de E4 en {s} deben sumar 100 %")
    if p.s("cdt.modelo") not in ("m5", "fase"):
        raise ParametroInvalido("cdt.modelo debe ser 'm5' o 'fase'")
    if p.s("etf.modelo") not in ("t5_estado", "jf"):
        raise ParametroInvalido("etf.modelo debe ser 't5_estado' o 'jf'")
    if p.s("val.modelo") not in ("estado", "ar1"):
        raise ParametroInvalido("val.modelo debe ser 'estado' o 'ar1'")
    if p.s("inm.reajuste_modelo") not in ("dane", "ipc"):
        raise ParametroInvalido("inm.reajuste_modelo debe ser 'dane' o 'ipc'")
    if p.s("eco.estado_inicial") not in ("Expansión", "Normal", "Crisis"):
        raise ParametroInvalido("El estado inicial debe ser Expansión, Normal o Crisis")


def advertencias(p: Params) -> list[str]:
    """Avisos que no impiden simular, pero que el usuario debe ver."""
    avisos = []
    n = p.i("exp.replicas")
    if n < 1068:
        error = 1.96 * (0.25 / max(n, 1)) ** 0.5 * 100
        avisos.append(f"Con {n} réplicas el error de la probabilidad de meta puede llegar a ±{error:.1f} puntos (con 1.068 es ±3).")
    for e in ("E1", "E2", "E3", "E4"):
        w = p.lista(f"{e}.pesos")
        costo = w[3] * p.f("inm.costo_compra")
        if w[0] - costo < p.f("inv.liquidez_min"):
            avisos.append(f"{e}: después de pagar la compra del inmueble la cuenta queda bajo la liquidez mínima; el simulador venderá ETF desde el primer mes.")
    pend = [t.nombre for t in p.tabla.values() if t.estado == "Pendiente"]
    if pend:
        avisos.append("Parámetros pendientes, sin datos (el proceso queda apagado): " + "; ".join(pend))
    return avisos

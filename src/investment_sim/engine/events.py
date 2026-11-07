"""Tipos de evento y lista de eventos futuros.

Cada evento lleva el número que tiene en el diagrama de eventos. La prioridad
resuelve los empates de tiempo con la regla del diagrama de flujo (página 1):
2 cambio de estado → 3 inflación → 4 Junta → 9 valorización → 14 revisión (E4)
→ 5 vencimiento de CDT → 7 y 8 contrato / vacancia → 6 arriendo → 10 reparación
→ 11 necesidad de liquidez → 13 venta del inmueble → 1 cierre de mes → 15 fin.
La cobertura de liquidez (12) no va en la lista: la llama el evento que la necesita.
"""
from __future__ import annotations

import heapq
import itertools
from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class Tipo(Enum):
    # (número en el diagrama, prioridad, nombre)
    INICIO = (0, 0, "Inicio")
    CRISIS_INICIO = (16, 1, "Inicio de la crisis forzada")
    CRISIS_FIN = (17, 1, "Fin de la crisis forzada")
    CAMBIO_ESTADO = (2, 2, "Cambio de estado económico")
    INFLACION = (3, 3, "Inflación anual")
    JUNTA = (4, 4, "Reunión de la Junta")
    VALORIZACION = (9, 5, "Valorización y gastos anuales")
    REVISION = (14, 6, "Revisión anual de la estrategia")
    VENCE_CDT = (5, 7, "Vencimiento de CDT")
    FIN_CONTRATO = (7, 8, "Fin de contrato de arriendo")
    FIN_VACANCIA = (8, 8, "Fin de vacancia")
    ARRIENDO = (6, 9, "Pago de arriendo")
    REPARACION = (10, 10, "Reparación extraordinaria")
    NECESIDAD = (11, 11, "Necesidad extraordinaria de liquidez")
    VENTA_INMUEBLE = (13, 13, "Venta del inmueble")
    CIERRE_MES = (1, 14, "Cierre de mes")
    FIN = (15, 15, "Fin de la simulación")

    @property
    def numero(self) -> int:
        return self.value[0]

    @property
    def prioridad(self) -> int:
        return self.value[1]

    @property
    def nombre(self) -> str:
        return self.value[2]


@dataclass
class Evento:
    tiempo: float
    tipo: Tipo
    datos: dict[str, Any] = field(default_factory=dict)
    secuencia: int = 0
    cancelado: bool = False


class ListaEventos:
    """Cola de prioridad ordenada por (tiempo, prioridad, secuencia)."""

    def __init__(self) -> None:
        self._heap: list[tuple[float, int, int, Evento]] = []
        self._contador = itertools.count()
        self.programados = 0

    def programar(self, tiempo: float, tipo: Tipo, **datos: Any) -> Evento:
        if tiempo != tiempo:  # NaN
            raise ValueError("Tiempo de evento inválido")
        sec = next(self._contador)
        ev = Evento(tiempo, tipo, datos, sec)
        heapq.heappush(self._heap, (tiempo, tipo.prioridad, sec, ev))
        self.programados += 1
        return ev

    def siguiente(self) -> Evento | None:
        while self._heap:
            _, _, _, ev = heapq.heappop(self._heap)
            if not ev.cancelado:
                return ev
        return None

    @staticmethod
    def cancelar(ev: Evento | None) -> None:
        if ev is not None:
            ev.cancelado = True

    def __len__(self) -> int:
        return sum(1 for *_, e in self._heap if not e.cancelado)

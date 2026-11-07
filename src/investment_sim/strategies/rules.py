"""Las cuatro estrategias del punto 11.

E1 a E3 compran y mantienen: solo venden para cubrir liquidez y renuevan los CDT
por el mismo capital. E4 empieza como E2 y cada año ajusta el portafolio
financiero (cuenta, CDT y ETF) al estado económico vigente; el inmueble no se toca.
"""
from __future__ import annotations

from dataclasses import dataclass

NOMBRES = {
    "E1": "E1 Seguridad",
    "E2": "E2 Moderada",
    "E3": "E3 Arriesgada",
    "E4": "E4 Dinámica",
}
_L = {"Expansión": "E", "Normal": "N", "Crisis": "C"}


@dataclass(frozen=True)
class Estrategia:
    codigo: str

    @property
    def nombre(self) -> str:
        return NOMBRES[self.codigo]

    @property
    def dinamica(self) -> bool:
        return self.codigo == "E4"

    def pesos_iniciales(self, p) -> dict[str, float]:
        c, d, e, i = p.lista(f"{self.codigo}.pesos")
        return {"cuenta": c, "cdt": d, "etf": e, "inmueble": i}

    def objetivo(self, estado_eco: str, p) -> dict[str, float]:
        """Pesos objetivo del portafolio financiero en la revisión (solo E4)."""
        c, d, e = p.lista(f"E4.obj_{_L[estado_eco]}")
        return {"cuenta": c, "cdt": d, "etf": e}


def todas() -> list[Estrategia]:
    return [Estrategia(c) for c in NOMBRES]

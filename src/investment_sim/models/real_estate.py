"""Valorización del inmueble y reajuste del canon (7.6 y 5.4.3)."""
from __future__ import annotations

_L = {"Expansión": "E", "Normal": "N", "Crisis": "C"}


def valorizacion_log(estado: str, z: float, r_anterior: float | None, p) -> float:
    """r = ln(V nuevo / V anterior) del año.

    estado: r ~ Normal(μ(S), σ(S)) (punto 8).
    ar1:    r = μ + φ·(r anterior − μ) + σ·Z (IPVU, alternativa de 7.6).
    """
    if p.s("val.modelo") == "ar1":
        mu, phi, sd = p.lista("val.ar1")
        prev = mu if r_anterior is None else r_anterior
        return mu + phi * (prev - mu) + sd * z
    letra = _L[estado]
    return p.f(f"val.mu_{letra}") + p.f(f"val.sd_{letra}") * z


def reajuste_canon(canon: float, valor: float, ipc_previo: float, z: float, p) -> float:
    """Canon renovado; ipc_previo en % anual.

    ipc:  canon·(1 + IPC previo), como en el diagrama.
    dane: canon·(1 + a + b·IPC previo + σ·Z), sin superar el IPC previo (Ley 820, art. 20).
    En los dos casos el canon no pasa del tope mensual sobre el valor (art. 18).
    """
    if p.s("inm.reajuste_modelo") == "dane":
        aumento = p.f("inm.reajuste_a") + p.f("inm.reajuste_b") * ipc_previo + p.f("inm.reajuste_sigma") * z
        aumento = min(aumento, ipc_previo)
    else:
        aumento = ipc_previo
    return min(canon * (1 + aumento / 100), p.f("inm.tope_canon") * valor)

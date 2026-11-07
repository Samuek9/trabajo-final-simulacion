"""TPM (decisión de la Junta), cuenta remunerada y tasa del CDT (7.2 a 7.4)."""
from __future__ import annotations

import numpy as np


def tendencia(cambio_3m: float, umbral: float) -> str:
    if cambio_3m > umbral:
        return "Subiendo"
    if cambio_3m < -umbral:
        return "Bajando"
    return "Estable"


def decidir(tend: str, u1: float, u2: float, p, bajadas_estres: bool = False) -> float:
    """Cambio de la TPM en puntos (negativo si baja).

    u1 elige bajar, mantener o subir con la fila de la tendencia; u2 elige el tamaño.
    Con bajadas_estres se usan los tamaños de bajada del escenario de estrés.
    """
    clave = {"Bajando": "junta.bajando", "Estable": "junta.estable", "Subiendo": "junta.subiendo"}[tend]
    k = int(np.searchsorted(np.cumsum(p.lista(clave)), u1, side="right"))  # 0 baja, 1 mantiene, 2 sube
    if k == 1 or k > 2:
        return 0.0
    if k == 2:
        tam, pr = p.lista("junta.subidas"), p.lista("junta.p_subidas")
    elif bajadas_estres:
        tam, pr = p.lista("junta.bajadas_estres"), p.lista("junta.p_bajadas_estres")
    else:
        tam, pr = p.lista("junta.bajadas"), p.lista("junta.p_bajadas")
    pr = np.array(pr) / np.sum(pr)
    j = int(np.searchsorted(np.cumsum(pr), u2, side="right"))
    paso = tam[min(j, len(tam) - 1)]
    return paso if k == 2 else -paso


def decidir_logit(cambio_3m: float, tpm: float, pi: float, u1: float, u2: float, p,
                  bajadas_estres: bool = False) -> float:
    """Logit ordenado (2001-2026): tendencia de la inflación y brecha de Taylor.

    latente = b_tend·Δπ(3 meses) + b_taylor·[TPM − (r* + π + 0,5·(π − meta))]
    P(baja) = logística(k1 − latente); P(sube) = 1 − logística(k2 − latente).
    """
    b_tend, b_taylor, k1, k2 = p.lista("junta.logit")
    brecha = tpm - (p.f("junta.r_neutral") + pi + 0.5 * (pi - p.f("junta.meta_inflacion")))
    lat = b_tend * cambio_3m + b_taylor * brecha
    p_baja = 1 / (1 + np.exp(-(k1 - lat)))
    p_sube = 1 - 1 / (1 + np.exp(-(k2 - lat)))
    if u1 < p_baja:
        if bajadas_estres:
            tam, pr = p.lista("junta.bajadas_estres"), p.lista("junta.p_bajadas_estres")
        else:
            tam, pr = p.lista("junta.bajadas_logit"), p.lista("junta.p_bajadas_logit")
        signo = -1.0
    elif u1 > 1 - p_sube:
        tam, pr = p.lista("junta.subidas_logit"), p.lista("junta.p_subidas_logit")
        signo = 1.0
    else:
        return 0.0
    pr = np.array(pr) / np.sum(pr)
    j = int(np.searchsorted(np.cumsum(pr), u2, side="right"))
    return signo * tam[min(j, len(tam) - 1)]


def limitar_tpm(tpm: float, p) -> float:
    return min(max(tpm, p.f("tpm.min")), p.f("tpm.max"))


def fase(cambio_tpm_3m: float) -> str:
    if cambio_tpm_3m > 1e-9:
        return "subida"
    if cambio_tpm_3m < -1e-9:
        return "bajada"
    return "estable"


def margen_cdt(cambio_tpm_3m: float, z: float, u: float, p, residuos: np.ndarray | None) -> float:
    """Margen del CDT sobre la TPM en puntos, acotado al rango observado.

    m5:   a + b·ΔTPM(3 meses) + ε; ε es un residuo observado elegido con u
          (o una t con u si no hay residuos).
    fase: Normal(μ, σ) de la fase de la TPM, con z.
    """
    if p.s("cdt.modelo") == "m5":
        if p.s("cdt.residuos") == "empirico" and residuos is not None and len(residuos):
            eps = float(residuos[min(int(u * len(residuos)), len(residuos) - 1)])
        else:
            from scipy.stats import t as t_dist
            eps = float(t_dist.ppf(min(max(u, 1e-9), 1 - 1e-9), p.f("cdt.t_nu")) * p.f("cdt.t_escala"))
        m = p.f("cdt.a") + p.f("cdt.b") * cambio_tpm_3m + eps
    else:
        mu, sd = p.lista(f"cdt.fase_{fase(cambio_tpm_3m)}")
        m = mu + sd * z
    return min(max(m, p.f("cdt.margen_min")), p.f("cdt.margen_max"))


def margen_esperado(cambio_tpm_3m: float, p) -> float:
    """Margen sin sorpresa: sirve como tasa de mercado en una venta anticipada."""
    if p.s("cdt.modelo") == "m5":
        return p.f("cdt.a") + p.f("cdt.b") * cambio_tpm_3m
    return p.lista(f"cdt.fase_{fase(cambio_tpm_3m)}")[0]


def valor_venta_anticipada(K: float, r0: float, r_mercado: float, castigo: float, plazo_anios: float,
                           transcurrido_anios: float) -> float:
    """K·(1 + r0)^T / (1 + r_mercado + c)^(T − τ). Tasas como fracción anual."""
    resto = max(plazo_anios - transcurrido_anios, 0.0)
    return K * (1 + r0) ** plazo_anios / (1 + r_mercado + castigo) ** resto

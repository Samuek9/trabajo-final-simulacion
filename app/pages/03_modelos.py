import pandas as pd
import streamlit as st

from components import estado, ui

p = estado.params("base")

st.title("Modelos y economía")
ui.subtitulo("Modelo seleccionado para cada proceso aleatorio. Los parámetros estimados son de consulta; el inversionista y las estrategias se configuran en sus páginas.")


def pc(x, d=2):
    return f"{x * 100:.{d}f} %".replace(".", ",")


def n(x, d=3):
    return f"{x:.{d}f}".replace(".", ",")


phi, sg, c = p.f("inf.phi"), p.f("inf.sigma"), p.f("inf.c")
lp = (2.718281828 ** (c / (1 - phi)))
logit = p.lista("junta.logit")
ar = p.lista("val.ar1")

MODELOS = [
    ("Estado de la economía", "PIB trimestral del DANE, 2006 a 2026",
     "Cadena de Markov de tres estados que cambia cada trimestre: P(estado siguiente = j | estado actual = i) = pᵢⱼ.",
     f"Expansión si el PIB crece más de {pc(p.f('eco.umbral_expansion'), 0)} y crisis si crece menos de {pc(p.f('eco.umbral_crisis'), 0)}. "
     f"Matriz (expansión, normal, crisis): expansión {n(p.f('markov.E_E'))} / {n(p.f('markov.E_N'))} / {n(p.f('markov.E_C'))}; "
     f"normal {n(p.f('markov.N_E'))} / {n(p.f('markov.N_N'))} / {n(p.f('markov.N_C'))}; "
     f"crisis {n(p.f('markov.C_E'))} / {n(p.f('markov.C_N'))} / {n(p.f('markov.C_C'))}."),
    ("Inflación", "IPC anual del DANE, 2001 a 2025",
     "AR(1) en logaritmos: ln πₜ = a + φ · ln πₜ₋₁ + σ · Zₜ, con Z normal estándar. Entre diciembres, línea recta.",
     f"a = {n(c)}; φ = {n(phi)}; σ = {n(sg)}; nivel de largo plazo = {n(lp, 2)} %."),
    ("Decisión de la Junta del Banco de la República", "202 decisiones de 2001 a 2026",
     "Logit ordenado: P(baja) = F(k₁ − x); P(sube) = 1 − F(k₂ − x), con x = b_tend · tendencia + b_taylor · brecha de Taylor y F logística. "
     "El tamaño del cambio se sortea de las frecuencias observadas.",
     f"b_tend = {n(logit[0])}; b_taylor = {n(logit[1])}; k₁ = {n(logit[2])}; k₂ = {n(logit[3])}. "
     f"TPM entre {n(p.f('tpm.min'), 2)} % y {n(p.f('tpm.max'), 0)} %."),
    ("Tasa de la cuenta remunerada", "19 tasas de cuatro cuentas digitales, 2023 a 2026",
     "Tasa de la cuenta = α × TPM.",
     f"α = {n(p.f('cta.alfa'))}; retención en la fuente de {pc(p.f('imp.retencion_cuenta'), 0)}."),
    ("Margen del CDT sobre la TPM", "CDT a 360 días y TPM semanal, 2008 a 2026",
     "Tasa del CDT = TPM + margen; margen = a + b · ΔTPM(3 meses) + ε, con ε sorteado de los residuos observados.",
     f"a = {n(p.f('cdt.a'))}; b = {n(p.f('cdt.b'))}; retención en la fuente de {pc(p.f('imp.retencion'), 0)}; "
     f"castigo por venta anticipada de {pc(p.f('cdt.castigo'), 0)}."),
    ("Rendimiento del ETF", "Valor neto del iShares COLCAP, 2011 a 2026",
     "g = ln(1 + R) ~ t de Student con ν grados de libertad, media μ y volatilidad σ(estado de la economía).",
     f"ν = {p.i('etf.nu')}; μ = {pc(p.f('etf.mu'))} mensual; σ = {pc(p.f('etf.vol_E'))} (expansión y normal) y {pc(p.f('etf.vol_C'))} (crisis)."),
    ("Valorización del inmueble", "Índice de precios de la vivienda usada, 2001 a 2025",
     "AR(1) sobre la variación logarítmica anual: rₜ = μ + φ · (rₜ₋₁ − μ) + σ · Zₜ.",
     f"μ = {pc(ar[0])} anual; φ = {n(ar[1])}; σ = {pc(ar[2])}; parte de la variación de 2025 ({pc(p.f('val.r0'))})."),
    ("Reajuste del canon", "IPC de arriendos del DANE, 2010 a 2025",
     "Reajuste = a + b · IPC(t−1) + ε, sin superar el IPC del año anterior (Ley 820 de 2003).",
     f"a = {n(p.f('inm.reajuste_a'))} %; b = {n(p.f('inm.reajuste_b'))}; σ de ε = {n(p.f('inm.reajuste_sigma'))} puntos; "
     f"canon inicial de {pc(p.f('inm.canon_inicial'))} mensual del valor."),
    ("Vacancia", "Tiempo en oferta de apartamentos en Medellín, 2025 a 2026",
     "Al vencer cada contrato, el inquilino no renueva con probabilidad q(estado); la vacancia dura una gamma de forma k y media m(estado).",
     f"k = {n(p.f('vac.forma'), 1)}; q = {pc(p.f('vac.p_E'), 0)} / {pc(p.f('vac.p_N'), 0)} / {pc(p.f('vac.p_C'), 0)}; "
     f"m = {n(p.f('vac.media_E'), 1)} / {n(p.f('vac.media_N'), 1)} / {n(p.f('vac.media_C'), 1)} meses (expansión, normal, crisis)."),
    ("Necesidades extraordinarias de liquidez", "Sin datos de monto; referencia de frecuencia en la encuesta IEFIC del DANE",
     "Proceso de Poisson de frecuencia λ(estado); monto lognormal con mediana igual a un porcentaje del capital inicial.",
     f"λ = {n(p.f('liq.tasa_E'), 4)} / {n(p.f('liq.tasa_N'), 4)} / {n(p.f('liq.tasa_C'), 4)} por mes; mediana = {pc(p.f('liq.mediana_pct'), 1)} del capital; "
     f"σ del logaritmo = {n(p.f('liq.sigma'), 1)}. Son supuestos del inversionista y se ajustan en la página Inversionista."),
]
CORTO = ["Cadena de Markov, 3 estados", "AR(1) en logaritmos", "Logit ordenado", "Proporcional a la TPM", "TPM + margen lineal",
         "t de Student", "AR(1) anual", "Lineal con tope legal", "Bernoulli + gamma", "Poisson + lognormal"]
filas = [(f[0], f[1], CORTO[i], f[3]) for i, f in enumerate(MODELOS)]
st.dataframe(pd.DataFrame(filas, columns=["Proceso", "Datos", "Modelo", "Parámetros estimados"]), hide_index=True, width="stretch",
             column_config={"Proceso": st.column_config.TextColumn(width="medium"), "Datos": st.column_config.TextColumn(width="medium"),
                            "Modelo": st.column_config.TextColumn(width="small"),
                            "Parámetros estimados": st.column_config.TextColumn(width="large")})
st.markdown("##### Especificación de cada modelo")
for f in MODELOS:
    with st.expander(f[0]):
        st.markdown(f[2])
st.caption("Las pruebas estadísticas que respaldan cada elección están en el informe y en el libro de datos del proyecto.")

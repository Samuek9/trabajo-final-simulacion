"""Figuras del informe a partir de outputs/resultados (no vuelve a simular, salvo la réplica de ejemplo).

  python tools/figuras_informe.py
"""
from __future__ import annotations

import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / "src"))
from investment_sim.config.parameters import cargar  # noqa: E402
from investment_sim.experiments.runner import una_replica  # noqa: E402

RES = RAIZ / "outputs" / "resultados"
FIG = RAIZ / "docs" / "figuras"
COL = {"E1": "#2F6B9A", "E2": "#3C8D5A", "E3": "#C0612B", "E4": "#7A4E9C"}
NOM = {"E1": "E1 Seguridad", "E2": "E2 Moderada", "E3": "E3 Arriesgada", "E4": "E4 Dinámica"}
ESC = {"base": "Base", "adverso": "Adverso", "estres": "Estrés"}
EST = ["E1", "E2", "E3", "E4"]
plt.rcParams.update({"font.family": "Arial", "font.size": 9, "axes.spines.top": False, "axes.spines.right": False,
                     "axes.grid": True, "grid.alpha": 0.25, "figure.dpi": 100})


def guardar(fig, nombre):
    fig.tight_layout()
    fig.savefig(FIG / nombre, dpi=200)
    plt.close(fig)


def coma(x, dec=0):
    return f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def abanico(b):
    fig, ejes = plt.subplots(1, 3, figsize=(10, 3.4), sharey=True)
    for ax, esc in zip(ejes, ESC):
        for e in EST:
            d = b[(b.escenario == esc) & (b.estrategia == e) & (b.serie == "W_real")].sort_values("mes")
            x = d.mes / 12
            ax.fill_between(x, d.p25 / 1e6, d.p75 / 1e6, color=COL[e], alpha=0.12, lw=0)
            ax.plot(x, d.p50 / 1e6, color=COL[e], lw=1.8, label=NOM[e])
        ax.axhline(1000, color="gray", ls="--", lw=1)
        ax.set_title(ESC[esc])
        ax.set_xlabel("Años")
    ejes[0].set_ylabel("Patrimonio real (millones de hoy)")
    ejes[0].legend(frameon=False, fontsize=8, loc="upper left")
    guardar(fig, "f01_abanico.png")


def cajas(tab):
    fig, ejes = plt.subplots(1, 3, figsize=(10, 3.4), sharey=True)
    for ax, esc in zip(ejes, ESC):
        datos = [tab[(tab.escenario == esc) & (tab.estrategia == e)].W_real / 1e6 for e in EST]
        bp = ax.boxplot(datos, whis=(5, 95), showfliers=False, patch_artist=True, widths=0.6)
        for caja, e in zip(bp["boxes"], EST):
            caja.set(facecolor=COL[e], alpha=0.35, edgecolor=COL[e])
        for m in bp["medians"]:
            m.set(color="black")
        ax.axhline(1000, color="gray", ls="--", lw=1)
        ax.axhline(700, color="gray", ls=":", lw=1)
        ax.set_xticks(range(1, 5), EST)
        ax.set_title(ESC[esc])
    ejes[0].set_ylabel("Patrimonio real final (millones de hoy)")
    guardar(fig, "f02_cajas.png")


def prob_meta(r):
    fig, ax = plt.subplots(figsize=(7.5, 3.2))
    ancho = 0.2
    for k, e in enumerate(EST):
        d = r[r.estrategia == e].set_index("escenario").loc[list(ESC)]
        x = np.arange(3) + (k - 1.5) * ancho
        ax.bar(x, d.p_meta * 100, ancho, color=COL[e], label=NOM[e],
               yerr=[(d.p_meta - d.p_meta_lo) * 100, (d.p_meta_hi - d.p_meta) * 100], capsize=2, error_kw={"lw": 0.8})
    ax.set_xticks(range(3), ESC.values())
    ax.set_ylabel("Probabilidad de meta (%)")
    ax.legend(frameon=False, ncol=4, fontsize=8)
    guardar(fig, "f03_prob_meta.png")


def iliquidez(b):
    fig, ejes = plt.subplots(1, 2, figsize=(9, 3.1), sharey=True)
    for ax, esc in zip(ejes, ("adverso", "estres")):
        for e in EST:
            d = b[(b.escenario == esc) & (b.estrategia == e) & (b.serie == "frac_iliquido")].sort_values("mes")
            ax.plot(d.mes / 12, d.p50 * 100, color=COL[e], lw=1.5, label=NOM[e])
        ax.set_title(ESC[esc])
        ax.set_xlabel("Años")
    ejes[0].set_ylabel("Réplicas en iliquidez (%)")
    ejes[0].legend(frameon=False, fontsize=8)
    guardar(fig, "f04_iliquidez.png")


def capital_meta(cm):
    fig, ejes = plt.subplots(1, 2, figsize=(10, 3.5))
    d = cm[(cm.escenario == "base") & (cm.estrategia == "E2")].pivot(index="meta", columns="capital", values="p_meta")
    im = ejes[0].imshow(d.values * 100, cmap="Blues", vmin=0, vmax=100, aspect="auto", origin="lower")
    ejes[0].set_xticks(range(d.shape[1]), [coma(c / 1e6) for c in d.columns])
    ejes[0].set_yticks(range(d.shape[0]), [coma(m / 1e6) for m in d.index])
    for i in range(d.shape[0]):
        for j in range(d.shape[1]):
            v = d.values[i, j] * 100
            ejes[0].text(j, i, f"{v:.0f}", ha="center", va="center", fontsize=7.5, color="white" if v > 55 else "black")
    ejes[0].set_xlabel("Capital inicial (millones)")
    ejes[0].set_ylabel("Meta real (millones)")
    ejes[0].set_title("E2, escenario base: probabilidad de meta (%)")
    ejes[0].grid(False)
    fig.colorbar(im, ax=ejes[0], fraction=0.04)
    for esc, ls in (("base", "-"), ("estres", "--")):
        for e in EST:
            q = cm[(cm.escenario == esc) & (cm.estrategia == e) & (cm.meta == 1e9)].sort_values("capital")
            ejes[1].plot(q.capital / 1e6, q.p_meta * 100, ls, color=COL[e], lw=1.5, label=f"{e} {ESC[esc].lower()}")
    ejes[1].axhline(80, color="gray", ls=":", lw=1)
    ejes[1].set_xlabel("Capital inicial (millones)")
    ejes[1].set_ylabel("Probabilidad de llegar a 1.000 M (%)")
    ejes[1].legend(frameon=False, fontsize=7, ncol=2)
    guardar(fig, "f05_capital_meta.png")


def tornado(t):
    fig, ejes = plt.subplots(1, 2, figsize=(10, 4.2), sharey=True)
    orden = t[t.estrategia == "E2"].assign(r=lambda d: d.rango_p_meta).sort_values("r").factor.tolist()
    for ax, e in zip(ejes, ("E2", "E3")):
        d = t[t.estrategia == e].set_index("factor").loc[orden]
        ax.barh(range(len(d)), (d.bajo_p_meta - d.ref_p_meta) * 100, color="#9DB9D5", label="Valor bajo")
        ax.barh(range(len(d)), (d.alto_p_meta - d.ref_p_meta) * 100, color="#1F4E78", alpha=0.85, label="Valor alto")
        ax.axvline(0, color="black", lw=0.8)
        ax.set_title(f"{NOM[e]}: referencia {d.ref_p_meta.iloc[0] * 100:.1f} %")
        ax.set_xlabel("Cambio en la probabilidad de meta (puntos)")
    ejes[0].set_yticks(range(len(orden)), [o.split(" (")[0] for o in orden], fontsize=8)
    ejes[0].legend(frameon=False, fontsize=8, loc="lower left")
    guardar(fig, "f06_tornado.png")


def nominal_real(nr):
    fig, ax = plt.subplots(figsize=(7.5, 3.0))
    d = nr[nr.escenario == "base"].set_index("estrategia").loc[EST]
    x = np.arange(4)
    ax.bar(x - 0.18, d.p_meta_nominal * 100, 0.36, color="#9DB9D5", label="Meta medida en pesos corrientes")
    ax.bar(x + 0.18, d.p_meta_real * 100, 0.36, color="#1F4E78", label="Meta medida en pesos de hoy")
    for i, (a, b) in enumerate(zip(d.p_meta_nominal, d.p_meta_real)):
        ax.text(i - 0.18, a * 100 + 1.5, f"{a * 100:.0f}", ha="center", fontsize=8)
        ax.text(i + 0.18, b * 100 + 1.5, f"{b * 100:.0f}", ha="center", fontsize=8)
    ax.set_xticks(x, [NOM[e] for e in EST])
    ax.set_ylabel("Probabilidad de llegar a 1.000 M (%)")
    ax.set_ylim(0, 110)
    ax.legend(frameon=False, fontsize=8, loc="upper right")
    guardar(fig, "f07_nominal_real.png")


def precision(pr):
    fig, ax = plt.subplots(figsize=(6, 2.8))
    ax.errorbar(pr.replicas, pr.p_meta * 100, yerr=[(pr.p_meta - pr.lo) * 100, (pr.hi - pr.p_meta) * 100], fmt="o",
                color="#1F4E78", capsize=3)
    ax.set_xscale("log")
    ax.set_xticks(pr.replicas, [coma(x) for x in pr.replicas])
    ax.set_xlabel("Réplicas")
    ax.set_ylabel("P(meta) de E2 (%)")
    guardar(fig, "f08_precision.png")


def replica():
    rep = una_replica(cargar("base"), "E4", cargar("base").i("exp.semilla"), traza=True)
    tr = rep.trayectoria
    meses = np.arange(len(tr["W"]))
    fig, ejes = plt.subplots(3, 1, figsize=(9, 5.6), sharex=True, gridspec_kw={"height_ratios": [2, 1, 0.6]})
    ejes[0].stackplot(meses, *[tr[k] / 1e6 for k in ("cuenta", "cdt", "etf", "inmueble")],
                      labels=["Cuenta", "CDT", "ETF", "Inmueble"], colors=["#7FA7C9", "#2F6B9A", "#C0612B", "#3C8D5A"], alpha=0.85)
    ejes[0].plot(meses, tr["W_real"] / 1e6, color="black", ls=":", label="Patrimonio real")
    ejes[0].set_ylabel("Millones")
    ejes[0].legend(frameon=False, fontsize=7.5, ncol=5, loc="upper left")
    ejes[1].plot(meses, tr["tpm"], color="#1F4E78", label="TPM")
    ejes[1].plot(meses, tr["inflacion"], color="#B03A2E", label="Inflación")
    ejes[1].set_ylabel("%")
    ejes[1].legend(frameon=False, fontsize=7.5)
    ejes[2].step(meses, tr["estado"], where="post", color="gray")
    ejes[2].set_yticks([0, 1, 2], ["Expansión", "Normal", "Crisis"])
    ejes[2].set_xlabel("Mes")
    guardar(fig, "f09_replica_E4.png")


def main():
    FIG.mkdir(parents=True, exist_ok=True)
    b = pd.read_csv(RES / "bandas.csv")
    abanico(b)
    cajas(pd.read_csv(RES / "replicas.csv.gz"))
    prob_meta(pd.read_csv(RES / "resumen.csv"))
    iliquidez(b)
    capital_meta(pd.read_csv(RES / "capital_meta.csv"))
    tornado(pd.read_csv(RES / "tornado.csv"))
    nominal_real(pd.read_csv(RES / "nominal_vs_real.csv"))
    precision(pd.read_csv(RES / "precision_replicas.csv"))
    replica()
    print("ok", sorted(p.name for p in FIG.glob("*.png")))


if __name__ == "__main__":
    main()

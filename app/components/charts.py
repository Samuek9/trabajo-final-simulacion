"""Gráficos Plotly del simulador. Cada uno trae el botón de cámara para bajarlo en PNG."""
from __future__ import annotations

import numpy as np
import pandas as pd
import plotly.graph_objects as go

from .ui import COLORES, NOMBRES

CONFIG = {"displaylogo": False, "toImageButtonOptions": {"format": "png", "scale": 2}}
LAYOUT = dict(margin=dict(l=10, r=10, t=40, b=10), legend=dict(orientation="h", y=-0.15),
              hovermode="x unified", font=dict(size=13))


def _rgba(hex_color: str, a: float) -> str:
    h = hex_color.lstrip("#")
    return f"rgba({int(h[0:2], 16)},{int(h[2:4], 16)},{int(h[4:6], 16)},{a})"


def abanico(tray: dict[str, np.ndarray], meta: float | None = None, titulo: str = "Patrimonio real (pesos de hoy)",
            millones: bool = True) -> go.Figure:
    """Mediana y bandas 5-95 y 25-75 por estrategia."""
    fig = go.Figure()
    esc = 1e6 if millones else 1
    for e, m in tray.items():
        meses = np.arange(m.shape[1])
        anios = meses / 12
        p5, p25, p50, p75, p95 = np.percentile(m / esc, [5, 25, 50, 75, 95], axis=0)
        c = COLORES[e]
        fig.add_trace(go.Scatter(x=np.r_[anios, anios[::-1]], y=np.r_[p95, p5[::-1]], fill="toself",
                                 fillcolor=_rgba(c, 0.10), line=dict(width=0), hoverinfo="skip",
                                 name=f"{e} 5-95", showlegend=False))
        fig.add_trace(go.Scatter(x=np.r_[anios, anios[::-1]], y=np.r_[p75, p25[::-1]], fill="toself",
                                 fillcolor=_rgba(c, 0.20), line=dict(width=0), hoverinfo="skip",
                                 name=f"{e} 25-75", showlegend=False))
        fig.add_trace(go.Scatter(x=anios, y=p50, line=dict(color=c, width=2.5), name=NOMBRES[e],
                                 hovertemplate="%{y:,.0f} M"))
    if meta:
        fig.add_hline(y=meta / esc, line_dash="dash", line_color="gray", annotation_text="Meta",
                      annotation_position="top left")
    fig.update_layout(title=titulo, xaxis_title="Años", yaxis_title="Millones de COP", **LAYOUT)
    return fig


def abanico_bandas(b: dict[str, pd.DataFrame], meta: float | None = None,
                   titulo: str = "Patrimonio real (pesos de hoy)") -> go.Figure:
    """Igual que abanico, pero desde percentiles por mes ya calculados (columnas p5 ... p95)."""
    fig = go.Figure()
    for e, d in b.items():
        x = np.arange(len(d)) / 12
        c = COLORES[e]
        for lo, hi, a in (("p5", "p95", 0.10), ("p25", "p75", 0.20)):
            fig.add_trace(go.Scatter(x=np.r_[x, x[::-1]], y=np.r_[d[hi], d[lo][::-1]] / 1e6, fill="toself",
                                     fillcolor=_rgba(c, a), line=dict(width=0), hoverinfo="skip", showlegend=False))
        fig.add_trace(go.Scatter(x=x, y=d["p50"] / 1e6, line=dict(color=c, width=2.5), name=NOMBRES[e],
                                 hovertemplate="%{y:,.0f} M"))
    if meta:
        fig.add_hline(y=meta / 1e6, line_dash="dash", line_color="gray", annotation_text="Meta",
                      annotation_position="top left")
    fig.update_layout(title=titulo, xaxis_title="Años", yaxis_title="Millones de COP", **LAYOUT)
    return fig


def serie_estrategias(s: dict[str, np.ndarray], titulo: str, eje: str, escala: float = 100) -> go.Figure:
    fig = go.Figure()
    for e, y in s.items():
        fig.add_trace(go.Scatter(x=np.arange(len(y)) / 12, y=np.asarray(y) * escala, name=NOMBRES[e],
                                 line=dict(color=COLORES[e], width=2)))
    fig.update_layout(title=titulo, xaxis_title="Años", yaxis_title=eje, **LAYOUT)
    return fig


def distribucion_final(tabla: pd.DataFrame, col: str = "W_real", meta: float | None = None) -> go.Figure:
    fig = go.Figure()
    for e, g in tabla.groupby("estrategia"):
        fig.add_trace(go.Violin(x=[NOMBRES[e]] * len(g), y=g[col] / 1e6, name=NOMBRES[e], line_color=COLORES[e],
                                box_visible=True, meanline_visible=True, points=False, spanmode="hard"))
    if meta:
        fig.add_hline(y=meta / 1e6, line_dash="dash", line_color="gray", annotation_text="Meta")
    fig.update_layout(title="Distribución del patrimonio real final", yaxis_title="Millones de COP (pesos de hoy)",
                      showlegend=False, **{k: v for k, v in LAYOUT.items() if k != "hovermode"})
    return fig


def riesgo_rentabilidad(res: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    for _, r in res.iterrows():
        e = r["estrategia"]
        fig.add_trace(go.Scatter(x=[r["p_iliquidez"] * 100], y=[r["p_meta"] * 100], mode="markers+text",
                                 marker=dict(size=18 + 40 * r["caida_max_p50"], color=COLORES[e], opacity=.85),
                                 text=[e], textposition="top center", name=NOMBRES[e],
                                 error_y=dict(type="data", symmetric=False, array=[(r["p_meta_hi"] - r["p_meta"]) * 100],
                                              arrayminus=[(r["p_meta"] - r["p_meta_lo"]) * 100]),
                                 hovertemplate=f"{NOMBRES[e]}<br>Meta: %{{y:.1f}} %<br>Iliquidez: %{{x:.1f}} %<extra></extra>"))
    fig.update_layout(title="Meta frente a iliquidez (tamaño = caída máxima mediana)",
                      xaxis_title="Probabilidad de algún episodio de iliquidez (%)",
                      yaxis_title="Probabilidad de alcanzar la meta (%)", **{k: v for k, v in LAYOUT.items() if k != "hovermode"})
    return fig


def iliquidez_tiempo(tray: dict[str, np.ndarray]) -> go.Figure:
    fig = go.Figure()
    for e, m in tray.items():
        fig.add_trace(go.Scatter(x=np.arange(m.shape[1]) / 12, y=m.mean(axis=0) * 100, name=NOMBRES[e],
                                 line=dict(color=COLORES[e], width=2)))
    fig.update_layout(title="Réplicas en iliquidez cada mes", xaxis_title="Años", yaxis_title="% de réplicas", **LAYOUT)
    return fig


def mapa_capital_meta(t: pd.DataFrame, estrategia: str) -> go.Figure:
    d = t[t.estrategia == estrategia].pivot(index="meta", columns="capital", values="p_meta").sort_index(ascending=False)
    fig = go.Figure(go.Heatmap(z=d.values * 100, x=[f"{c / 1e6:,.0f}" for c in d.columns], y=[f"{m / 1e6:,.0f}" for m in d.index],
                               colorscale="Blues", zmin=0, zmax=100, text=np.round(d.values * 100).astype(int),
                               texttemplate="%{text} %", colorbar=dict(title="%")))
    fig.update_layout(title=f"{NOMBRES[estrategia]}: probabilidad de alcanzar la meta",
                      xaxis_title="Capital inicial (millones)", yaxis_title="Meta real (millones)",
                      **{k: v for k, v in LAYOUT.items() if k != "hovermode"})
    return fig


def tornado(t: pd.DataFrame, estrategia: str, medida: str = "p_meta") -> go.Figure:
    d = t[t.estrategia == estrategia].copy()
    if medida == "p_meta":
        ref, bajo, alto, esc, unidad = "ref_p_meta", "bajo_p_meta", "alto_p_meta", 100, "puntos"
    else:
        ref, bajo, alto, esc, unidad = "ref_W", "bajo_W", "alto_W", 1e-6, "millones"
    d["b"] = (d[bajo] - d[ref]) * esc
    d["a"] = (d[alto] - d[ref]) * esc
    d["rango"] = (d["a"] - d["b"]).abs()
    d = d.sort_values("rango")
    fig = go.Figure()
    fig.add_trace(go.Bar(y=d.factor, x=d.b, orientation="h", name="Valor bajo", marker_color="#9DB9D5"))
    fig.add_trace(go.Bar(y=d.factor, x=d.a, orientation="h", name="Valor alto", marker_color="#1F4E78"))
    fig.update_layout(barmode="overlay", title=f"{NOMBRES[estrategia]}: cambio frente a los valores ingresados ({unidad})",
                      xaxis_title=f"Cambio ({unidad})", height=120 + 34 * len(d),
                      **{k: v for k, v in LAYOUT.items() if k != "hovermode"})
    return fig


def trayectoria_replica(tr: dict[str, np.ndarray]) -> go.Figure:
    from plotly.subplots import make_subplots
    meses = np.arange(len(tr["W"]))
    fig = make_subplots(rows=3, cols=1, shared_xaxes=True, row_heights=[0.5, 0.25, 0.25], vertical_spacing=0.06,
                        subplot_titles=("Composición del patrimonio (millones)", "TPM e inflación (%)", "Estado económico"))
    for k, nombre, c in (("cuenta", "Cuenta", "#7FA7C9"), ("cdt", "CDT", "#2F6B9A"), ("etf", "ETF", "#C0612B"),
                         ("inmueble", "Inmueble", "#3C8D5A")):
        fig.add_trace(go.Scatter(x=meses, y=tr[k] / 1e6, stackgroup="w", name=nombre, line=dict(width=0.5, color=c)), 1, 1)
    fig.add_trace(go.Scatter(x=meses, y=tr["W_real"] / 1e6, name="Patrimonio real", line=dict(color="black", dash="dot")), 1, 1)
    fig.add_trace(go.Scatter(x=meses, y=tr["tpm"], name="TPM", line=dict(color="#1F4E78")), 2, 1)
    fig.add_trace(go.Scatter(x=meses, y=tr["inflacion"], name="Inflación", line=dict(color="#B03A2E")), 2, 1)
    fig.add_trace(go.Scatter(x=meses, y=tr["estado"], name="Estado", mode="lines", line=dict(shape="hv", color="gray")), 3, 1)
    fig.update_yaxes(tickvals=[0, 1, 2], ticktext=["Expansión", "Normal", "Crisis"], row=3, col=1)
    fig.update_xaxes(title_text="Mes", row=3, col=1)
    fig.update_layout(height=720, **LAYOUT)
    return fig

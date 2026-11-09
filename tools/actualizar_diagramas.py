"""Pone al día los textos de los diagramas del equipo (docs/diagramas) con el modelo que usa el simulador.

Los diagramas se hicieron en Graphviz y solo existe el PDF. Este script reemplaza línea por línea
el texto de las cajas que cambiaron (sin mover cajas ni flechas), cambia la etiqueta de la flecha del estado
económico a la valorización (ahora solo pesa en el estrés) y vuelve a generar las imágenes PNG que usa el informe.
Los PDF originales del equipo quedan en insumos_equipo/06_estrategias_escenarios_diseno.

  python tools/actualizar_diagramas.py
"""
from __future__ import annotations

from pathlib import Path

import fitz

RAIZ = Path(__file__).resolve().parents[1]
ORIGEN = RAIZ.parent / "insumos_equipo" / "06_estrategias_escenarios_diseno"
DESTINO = RAIZ / "docs" / "diagramas"
FUENTES = {False: r"C:\Windows\Fonts\arial.ttf", True: r"C:\Windows\Fonts\ariali.ttf"}

EVENTOS = [
    ("• Estado económico S = Normal; TPM = 12,25 %; tasa cuenta = 0,913 × TPM;",
     "• Estado económico S = Normal; TPM = 12,25 %; tasa cuenta = 0,898 × TPM;"),
    ("• Abre los CDT: tasa = TPM + margen de la fase; inmueble ocupado, canon =",
     "• Abre los CDT: tasa = TPM + margen (modelo M5); inmueble ocupado, canon ="),
    ("• Tendencia = inflación de hoy − inflación de hace 3 meses → Subiendo /",
     "• Tendencia = inflación de hoy − inflación de hace 3 meses; brecha ="),
    ("Estable / Bajando", "   TPM − (1 + π + 0,5·(π − 3))  (regla de Taylor)"),
    ("• U₁: decisión (discreta empírica):  Subiendo (0 | 0,18 | 0,82)  Estable",
     "• U1: logit ordenado con x = 1,647·tendencia − 0,202·brecha:"),
    ("(0,13 | 0,62 | 0,25)  Bajando (0,62 | 0,38 | 0)  [Baja | Mantiene |",
     "   P(baja) = F(−1,599 − x);  P(sube) = 1 − F(1,526 − x)  [F logística;"),
    ("Sube]", "   si no, mantiene]"),
    ("• U₂: tamaño si sube 0,25 / 0,50 / 0,75 / 1,00 / 1,50 %;  si baja 0,25 /",
     "• U2: tamaño si sube 0,25 a 1,50 puntos;  si baja 0,25 a 2 puntos"),
    ("0,50 %  (en crisis o estrés, bajadas de hasta 2 %)",
     "   (frecuencias de 2001 a 2026; en estrés, bajadas de 0,5 a 2)"),
    ("• Tasa de la cuenta = 0,913 × TPM; fase del CDT (TPM hoy vs. hace 3 meses)",
     "• Tasa de la cuenta = 0,898 × TPM; ΔTPM de 3 meses para el margen del CDT"),
    ("TPM y fase", "TPM y ΔTPM"),
    ("• ETF: g = 0,43 % + s(S) × T,  T ~ t de Student (ν = 5);  s(S) según el",
     "• ETF: g = 0,36 % + s(S) × T,  T ~ t de Student (ν = 5);  s(S) según el"),
    ("• Cuenta: interés = saldo × [(1 + 0,913·TPM)^(1/12) − 1] × (1 − 0,04 de",
     "• Cuenta: interés = saldo × [(1 + 0,898·TPM)^(1/12) − 1] × (1 − 0,07 de"),
    ("valorización del inmueble,", "   pesos objetivo de E4,"),
    ("μ, σ valorización", "−5 %/año en crisis forzada (estrés)"),
    ("• r ~ Normal(μ(S), σ(S)):  expansión 11,9 % (3,5 %);  normal y crisis 8,3",
     "• r = 0,0755 + 0,676·(r anterior − 0,0755) + 0,0236·Z  (AR(1) del IPVU,"),
    ("% (2,2 %)  [en logaritmos]", "   igual en los tres estados; en logaritmos)"),
    ("• Si renueva: canon × (1 + inflación de dic. anterior), tope 1 % mensual",
     "• Si renueva: canon × (1 + 1,68 % + 0,48·IPC previo + ε), sin pasar el IPC;"),
    ("del valor (Ley 820)", "   tope 1 % mensual del valor (Ley 820)"),
    ("• Margen ~ Normal(μ, σ) de la fase:  Subida (2,68 %; 1,98 %)  Estable",
     "• Margen M5 = 1,06 + 0,67·ΔTPM (3 meses) + ε, con ε sorteado entre"),
    ("(0,86 %; 0,51 %)  Bajada (0,16 %; 1,08 %), limitado al rango observado",
     "   los residuos observados (bootstrap, 982 semanas)"),
    ("r_mercado + 1 %)^(T−τ); r_mercado = TPM + margen 180 días de la fase",
     "   r_mercado + 1 %)^(T−τ); r_mercado = TPM + margen M5 sin sorpresa"),
    ("10 %, el dinero llega en 3 a 6 meses", "   10 %, el dinero llega en 4 a 7 meses"),
    ("3 a 6 meses después de decidir la venta urgente", "4 a 7 meses después de decidir la venta urgente"),
    ("• Tiempo de venta ~ Uniforme(3, 6) meses (generado en el evento 12)",
     "• Tiempo de venta ~ Uniforme(4, 7) meses (generado en el evento 12)"),
    ("+Uniforme(3, 6)", "+Uniforme(4, 7)"),
    ("• Monto ~ Lognormal  (λ_rep y monto: supuesto del equipo, por calibrar)",
     "• Monto ~ Lognormal. Apagado: λ_rep = 0 porque no hay datos"),
    ("TPM, fase", "TPM, ΔTPM"),
    ("Tasa de política y fase del ciclo (subida / estable / bajada)", "Tasa de política y su cambio en los últimos 3 meses"),
]


# Diagrama de flujo: (página, texto viejo, texto nuevo). El texto va centrado en su caja.
FLUJO = [
    (1, "volatilidad del ETF, valorización", "volatilidad del ETF, pesos de E4,"),
    (1, "del inmueble, vacancia, frecuencia", "vacancia, frecuencia"),
    (1, "Tendencia = inflación de hoy −", "x = 1,647·tendencia − 0,202·brecha"),
    (1, "inflación de hace 3 meses", "(tendencia de 3 meses; brecha de Taylor)"),
    (1, "¿Tendencia?", "Probabilidades"),
    (1, "Subiendo", "baja"),
    (1, "Estable", "mantiene"),
    (1, "Bajando", "sube"),
    (1, "Subiendo: P(Baja, Mantiene, Sube)", "P(baja) = F(−1,599 − x)"),
    (1, "= 0 | 0,18 | 0,82", "(F logística)"),
    (1, "Estable: 0,13 | 0,62 | 0,25", "P(mantiene) = el resto"),
    (1, "Bajando: 0,62 | 0,38 | 0", "P(sube) = 1 − F(1,526 − x)"),
    (1, "Generar U₂: tamaño 0,25 / 0,50 /", "Generar U2: tamaño 0,25 / 0,50 /"),
    (1, "0,75 / 1,00 / 1,50 % (prob. 22 /", "0,75 / 1 / 1,50 (prob. 48 / 25 /"),
    (1, "11 / 11 / 44 / 11 %)", "4 / 19 / 4 %)"),
    (1, "Generar U₂: tamaño 0,25 / 0,50 %", "Generar U2: tamaño 0,25 a 2"),
    (1, "(prob. 40 / 60 %); en crisis o", "puntos (frecuencias 2001-2026);"),
    (1, "estrés hasta 2 %", "en estrés, de 0,5 a 2"),
    (1, "Tasa de la cuenta = 0,913 × TPM;", "Tasa de la cuenta = 0,898 × TPM;"),
    (1, "fase del CDT (TPM de hoy vs. hace", "ΔTPM de 3 meses para el margen"),
    (1, "3 meses)", "del CDT"),
    (2, "g = 0,43 % + s(S) × T  (s según el", "g = 0,36 % + s(S) × T  (s según el"),
    (2, "+ 0,913·TPM)^(1/12) − 1] × (1 −", "+ 0,898·TPM)^(1/12) − 1] × (1 −"),
    (2, "0,04)", "0,07)"),
    (2, "Generar margen ~ Normal(μ, σ) de", "Generar margen M5 = 1,06 +"),
    (2, "la fase; limitar al rango", "0,67·ΔTPM (3 meses) + ε,"),
    (2, "observado. Subida 2,68 % (1,98);", "ε sorteado entre los residuos"),
    (2, "Estable 0,86 % (0,51); Bajada 0,16", "observados (bootstrap, 982"),
    (2, "% (1,08)", "semanas)"),
    (3, "Renueva: canon × (1 + inflación de", "Renueva: canon × (1 + 1,68 % +"),
    (3, "dic. anterior), tope 1 % mensual", "0,48·IPC previo + ε), sin pasar"),
    (3, "del valor", "el IPC; tope 1 % del valor"),
    (3, "Generar r ~ Normal(μ(S), σ(S)):", "Generar r = 0,0755 + 0,676·(r"),
    (3, "expansión 11,9 % (3,5 %); normal y", "anterior − 0,0755) + 0,0236·Z"),
    (3, "crisis 8,3 % (2,2 %)", "(AR(1) del IPVU, sin estados)"),
    (3, "(parámetros por calibrar)", "(apagado: λ_rep = 0, sin datos)"),
    (4, "Uniforme(3, 6) meses; programar", "Uniforme(4, 7) meses; programar"),
]


def _lineas(page):
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            texto = "".join(s["text"] for s in l["spans"])
            yield texto.strip(), l


def reemplazar(page, cambios, quitar_flechas=(), centrado=False) -> list[str]:
    pendientes, hechos = dict(cambios), []
    nuevos = []
    for texto, l in list(_lineas(page)):
        if texto in pendientes:
            s0 = l["spans"][0]
            nuevos.append((fitz.Rect(l["bbox"]), s0["origin"], s0["size"], "Italic" in s0["font"], s0["color"], pendientes.pop(texto)))
            hechos.append(texto)
    # Solo la franja central de la línea: las líneas vecinas están a una altura de letra y se tocan.
    for r, origen, tam, *_ in nuevos:
        page.add_redact_annot(fitz.Rect(r.x0 - 0.5, origen[1] - 0.55 * tam, r.x1 + 0.5, origen[1] - 0.15 * tam), fill=False)
    for punto in quitar_flechas:
        page.add_redact_annot(fitz.Rect(punto, punto) + (-1.5, -1.5, 1.5, 1.5), fill=False)
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                          graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_TOUCHED if quitar_flechas else fitz.PDF_REDACT_LINE_ART_NONE)
    cajas = [(d["rect"], not any(it[0] in ("re", "qu", "c") for it in d["items"]))
             for d in page.get_drawings() if d["rect"].width < 450 and d["rect"].height > 15]
    for r, origen, tam, italica, color, nuevo in nuevos:
        if not nuevo:
            continue
        nombre = "ari" if italica else "ar"
        page.insert_font(fontname=nombre, fontfile=FUENTES[italica])
        rgb = tuple(((color >> k) & 255) / 255 for k in (16, 8, 0))
        fuente = fitz.Font(fontfile=FUENTES[italica])
        if centrado:
            limite = _ancho_caja(cajas, r) if tam >= 12 else r.width + 30
            while fuente.text_length(nuevo, fontsize=tam) > limite:
                tam -= 0.25
            x = (r.x0 + r.x1) / 2 - fuente.text_length(nuevo, fontsize=tam) / 2
        else:
            x = origen[0] if not nuevo.startswith("   ") else r.x0
        page.insert_text((x, origen[1]), nuevo, fontname=nombre, fontsize=tam, color=rgb)
    return list(pendientes)


def _ancho_caja(cajas, r) -> float:
    """Ancho útil de la caja más chica que contiene la línea (en un rombo, algo menos de la mitad)."""
    c = fitz.Point((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2)
    dentro = [(b, rombo) for b, rombo in cajas if b.contains(c)]
    if not dentro:
        return r.width + 20
    b, rombo = min(dentro, key=lambda x: x[0].width * x[0].height)
    return b.width * (0.55 if rombo else 1) - 10


def png(pdf: Path, salida: Path, pagina: int = 0, dpi: int = 150) -> None:
    with fitz.open(pdf) as d:
        d[pagina].get_pixmap(dpi=dpi).save(salida)


def main() -> None:
    doc = fitz.open(ORIGEN / "diagrama_eventos.pdf")
    falta = reemplazar(doc[0], EVENTOS)
    doc.subset_fonts()
    doc.save(DESTINO / "diagrama_eventos.pdf", garbage=3, deflate=True)
    print("eventos, sin encontrar:", falta)
    png(DESTINO / "diagrama_eventos.pdf", DESTINO / "diagrama_eventos.png", dpi=110)

    doc = fitz.open(ORIGEN / "diagrama_de_flujo_completo.pdf")
    for k, page in enumerate(doc):
        falta = reemplazar(page, [(a, b) for p, a, b in FLUJO if p == k], centrado=True)
        if falta:
            print("flujo, página", k + 1, "sin encontrar:", falta)
    doc.subset_fonts()
    doc.save(DESTINO / "diagrama_de_flujo_completo.pdf", garbage=3, deflate=True)
    for k in range(len(doc)):
        png(DESTINO / "diagrama_de_flujo_completo.pdf", DESTINO / f"flujo_pagina_{k + 1}.png", k, dpi=110)


if __name__ == "__main__":
    main()

"""Corta los diagramas (docs/diagramas/*.pdf) en partes legibles para el Anexo C del informe.

  python tools/diagramas_anexo.py

Cada hoja del PDF se divide en el menor número de partes que, puestas en una página A3
horizontal, dejan el texto de las cajas en 7 puntos o más. Las partes se solapan un poco para
no cortar una caja sin repetirla. Salida: docs/diagramas/anexo/*.png y partes.json.
"""
from __future__ import annotations

import json
import math
from pathlib import Path

import fitz

RAIZ = Path(__file__).resolve().parents[1]
DIAG = RAIZ / "docs" / "diagramas"
SALIDA = DIAG / "anexo"
ANCHO, ALTO = 1040.0, 580.0   # espacio útil de un A3 horizontal (420 × 297 mm, márgenes de 2,5 cm) sin la leyenda, en pt
ESCALA_MIN = 0.62              # 11-12 pt del original quedan en unos 7 pt
SOLAPE = 30.0                  # pt del original que se repiten entre partes vecinas
DPI = 200

HOJAS = [("diagrama_eventos.pdf", "Diagrama de eventos"),
         ("diagrama_de_flujo_completo.pdf", "Diagrama de flujo")]


def partes(w: float, h: float) -> tuple[int, int]:
    """Menor número de partes (columnas, filas) con escala suficiente."""
    mejor = None
    for cols in range(1, 6):
        for filas in range(1, 8):
            tw, th = w / cols + SOLAPE, h / filas + SOLAPE
            if min(ANCHO / tw, ALTO / th) >= ESCALA_MIN:
                if mejor is None or cols * filas < mejor[0] * mejor[1]:
                    mejor = (cols, filas)
                break
    return mejor


def main() -> None:
    SALIDA.mkdir(parents=True, exist_ok=True)
    for viejo in SALIDA.glob("*.png"):
        viejo.unlink()
    lista = []
    for archivo, nombre in HOJAS:
        doc = fitz.open(DIAG / archivo)
        for k, pag in enumerate(doc):
            r = pag.rect
            cols, filas = partes(r.width, r.height)
            n = cols * filas
            for i in range(filas):
                for j in range(cols):
                    x0 = max(r.x0, r.x0 + j * r.width / cols - SOLAPE / 2)
                    x1 = min(r.x1, r.x0 + (j + 1) * r.width / cols + SOLAPE / 2)
                    y0 = max(r.y0, r.y0 + i * r.height / filas - SOLAPE / 2)
                    y1 = min(r.y1, r.y0 + (i + 1) * r.height / filas + SOLAPE / 2)
                    clip = fitz.Rect(x0, y0, x1, y1)
                    escala = min(ANCHO / clip.width, ALTO / clip.height, 1.0)
                    pix = pag.get_pixmap(matrix=fitz.Matrix(escala * DPI / 72, escala * DPI / 72), clip=clip, alpha=False)
                    destino = SALIDA / f"{Path(archivo).stem}_p{k + 1}_{i + 1}_{j + 1}.png"
                    pix.save(destino)
                    hoja = f", página {k + 1} de {len(doc)}" if len(doc) > 1 else ""
                    parte = f", parte {i * cols + j + 1} de {n} (fila {i + 1}, columna {j + 1})" if n > 1 else ""
                    lista.append({"archivo": destino.name, "leyenda": f"{nombre}{hoja}{parte}.",
                                  "ancho_pulgadas": round(clip.width * escala / 72, 2),
                                  "escala": round(escala, 3)})
    (SALIDA / "partes.json").write_text(json.dumps(lista, indent=2, ensure_ascii=False), encoding="utf-8")
    print(len(lista), "partes; escala mínima", min(x["escala"] for x in lista))


if __name__ == "__main__":
    main()

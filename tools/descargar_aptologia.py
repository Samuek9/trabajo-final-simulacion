"""Descarga completa de los avisos de arriendo de apartamentos en Medellín (Aptologia).

API abierta, licencia CC BY 4.0: https://www.aptologia.com/datos-abiertos
La API entrega 24 avisos por página. Se recorren todas las páginas sin tope de precio
(la descarga del equipo usó max=3.500.000 y se cortó en 4.917 avisos) y se quitan los
avisos repetidos por id, porque el listado cambia mientras se pagina.

  python tools/descargar_aptologia.py

Escribe data/fuentes/inmueble/raw/aptologia_avisos_arriendo_medellin_<fecha>.csv.
Son precios pedidos de avisos activos, no cánones de contratos firmados.
"""
from __future__ import annotations

import json
import time
import urllib.request
from datetime import date
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parents[1]
URL = "https://www.aptologia.com/api/abierta/avisos?ciudad=medellin&tipo=apartamento&pagina={}"
PAUSA = 0.4


def pagina(n: int) -> dict:
    req = urllib.request.Request(URL.format(n), headers={"User-Agent": "trabajo-final-simulacion (uso academico)"})
    for intento in range(4):
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                return json.loads(r.read().decode("utf-8"))
        except Exception:
            time.sleep(2 * (intento + 1))
    raise RuntimeError(f"La página {n} falló cuatro veces")


def main() -> None:
    avisos, n, total = {}, 0, None
    while True:
        d = pagina(n)
        total = total or d["total"]
        for a in d["avisos"]:
            avisos[a["id"]] = a
        if n % 100 == 0:
            print(f"página {n}: {len(avisos)} avisos de {total}", flush=True)
        if not d["hay_mas"] or not d["avisos"]:
            break
        n += 1
        time.sleep(PAUSA)
    df = pd.DataFrame(avisos.values())
    df["descargado"] = date.today().isoformat()
    salida = RAIZ / "data" / "fuentes" / "inmueble" / "raw" / f"aptologia_avisos_arriendo_medellin_{date.today():%Y-%m}.csv"
    df.to_csv(salida, index=False, encoding="utf-8")
    print(f"{len(df)} avisos únicos de {total} reportados, {n + 1} páginas -> {salida}")


if __name__ == "__main__":
    main()

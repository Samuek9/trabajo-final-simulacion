"""Frecuencia de meses con déficit cubierto con ahorros o venta de activos (IEFIC 2017 y 2018).

Fuente: Departamento Administrativo Nacional de Estadística: www.dane.gov.co
Encuesta de Carga Financiera y Educación Financiera de los Hogares (IEFIC), catálogo 626.
Los microdatos van en data/fuentes/liquidez/iefic_microdatos/ y no se publican (Ley 79 de
1993); este script solo escribe el resumen.

Preguntas (Cuestionario_IEFIC_2018.pdf, capítulo B):
  P2479  En el último mes los gastos del hogar frente a sus ingresos han sido:
         1 mayores, 2 menores, 3 iguales.
  P2482  En los meses en que los gastos han superado los ingresos, ¿qué han hecho?
         _2 vender un activo, _3 disponer de los ahorros.

Un hogar por directorio (fila del jefe), Medellín (departamento 05), ponderado con FEX_C.
El ingreso del hogar es la suma de INGTOTOB de sus miembros.

  python tools/iefic_liquidez.py
"""
from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

RAIZ = Path(__file__).resolve().parents[1]
CARPETA = RAIZ / "data" / "fuentes" / "liquidez"


def num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s.astype(str).str.strip().str.replace(",", "."), errors="coerce")


def hogares(anio: int) -> pd.DataFrame:
    d = pd.read_csv(CARPETA / "iefic_microdatos" / f"IEFIC_{anio}" / f"IEFIC_{anio}.csv", sep=None,
                    engine="python", encoding="latin-1")
    d.columns = ["SECUENCIA_P" if "SECUENCIA_P" in c else c for c in d.columns]
    d["ing"] = num(d.INGTOTOB).fillna(0)
    ingreso = d.groupby("DIRECTORIO").ing.sum()
    h = d[d.P6050 == 1].drop_duplicates("DIRECTORIO").set_index("DIRECTORIO")
    h["ing_hogar"] = ingreso
    h["fex"] = num(h.FEX_C)
    h = h[(h.DEPARTAMENTO == 5) & (h.P2479.astype(str).str.strip() != "")].copy()
    h["deficit"] = h.P2479.astype(str).str.strip() == "1"
    h["con_patrimonio"] = h.deficit & ((h.P2482_3.astype(str).str.strip() == "1") | (h.P2482_2.astype(str).str.strip() == "1"))
    return h


def resumen() -> pd.DataFrame:
    filas = []
    for anio in (2017, 2018):
        h = hogares(anio)
        for grupo, cuantil in (("Todos", 0.0), ("Ingreso del hogar en el 20 % más alto", 0.8),
                               ("Ingreso del hogar en el 10 % más alto", 0.9)):
            corte = float(np.nanquantile(h.ing_hogar, cuantil)) if cuantil else 0.0
            g = h[h.ing_hogar >= corte]
            p = float(np.average(g.con_patrimonio, weights=g.fex))
            filas.append({"anio": anio, "grupo": grupo, "hogares": len(g), "ingreso_mensual_desde": corte,
                          "p_mes_deficit": float(np.average(g.deficit, weights=g.fex)),
                          "p_mes_deficit_con_ahorros_o_activos": p, "meses_entre_eventos": 1 / p})
    return pd.DataFrame(filas)


if __name__ == "__main__":
    r = resumen()
    r.to_csv(CARPETA / "iefic_resumen_medellin.csv", index=False, encoding="utf-8")
    print(r.round(3).to_string(index=False))

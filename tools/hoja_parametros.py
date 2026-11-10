"""Escribe la hoja PARAMETROS con formato a partir de tools/parametros_def.py."""
from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

sys.path.insert(0, str(Path(__file__).parent))
from parametros_def import FILAS, VISIBLES_PREFIJOS  # noqa: E402

COLUMNAS = ["ID", "Clave", "Grupo", "Parámetro", "Base", "Adverso", "Estrés", "Unidad", "Estado", "Fuente", "Soporte"]
ANCHOS = [7, 24, 13, 52, 16, 14, 14, 20, 12, 52, 18]
COLUMNAS_USUARIO = ["ID", "Clave", "Grupo", "Parámetro", "Valor", "Unidad"]
ANCHOS_USUARIO = [7, 24, 16, 68, 16, 24]
COLOR_ESTADO = {
    "Configurable": "FFF2CC", "Verificado": "D9EAD3", "Estimado": "D9EAD3", "Calibrado": "FCE5CD",
    "Supuesto": "FCE5CD", "Apagado": "F4CCCC", "Calculado": "DDEBF7",
}
AZUL = "404040"  # gris oscuro: encabezados sobrios


def es_visible(clave: str) -> bool:
    return clave.startswith(VISIBLES_PREFIJOS)


def _formato(cel, v):
    if isinstance(v, float):
        cel.number_format = "0.0000" if abs(v) < 10 else "#,##0.00"
    if isinstance(v, int) and abs(v) >= 1000:
        cel.number_format = "#,##0"


def escribir(ws) -> None:
    """PARAMETROS: lo que el usuario puede cambiar (inversionista, experimento, estrategias y necesidades de liquidez)."""
    ws["A1"] = "Parámetros del inversionista y del experimento"
    ws["A1"].font = Font(bold=True, size=14, color=AZUL)
    ws["A2"] = ("Los valores de esta hoja se pueden cambiar; el simulador los lee al iniciar. Los procesos aleatorios "
                "(inflación, tasas, ETF, inmueble, vacancia, estados de la economía) están estimados con datos y quedan fijos.")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:F2")
    ws.row_dimensions[2].height = 34
    fila0 = 4
    borde = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
    for j, c in enumerate(COLUMNAS_USUARIO, 1):
        cel = ws.cell(fila0, j, c)
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor=AZUL)
        cel.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(j)].width = ANCHOS_USUARIO[j - 1]
    filas = [f for f in FILAS if es_visible(f[1])]
    for i, f in enumerate(filas, fila0 + 1):
        for j, v in enumerate((f[0], f[1], f[2], f[3], f[4], f[7]), 1):
            cel = ws.cell(i, j, v)
            cel.border = borde
            cel.alignment = Alignment(vertical="top", wrap_text=j == 4)
            if j == 5:
                _formato(cel, v)
                cel.fill = PatternFill("solid", fgColor=COLOR_ESTADO["Configurable"])
    ws.freeze_panes = ws.cell(fila0 + 1, 3)


def escribir_modelo(ws) -> None:
    """MODELO (oculta): valores estimados con datos, fijos. El simulador los lee junto con PARAMETROS."""
    ws["A1"] = "Valores del modelo (fijos)"
    ws["A1"].font = Font(bold=True, size=14, color=AZUL)
    ws["A2"] = ("Adverso y Estrés vacíos = igual que la base; en Estrés el valor rige solo mientras dura la crisis del escenario.")
    ws.merge_cells("A2:K2")
    fila0 = 4
    borde = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
    for j, c in enumerate(COLUMNAS, 1):
        cel = ws.cell(fila0, j, c)
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor=AZUL)
        cel.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(j)].width = ANCHOS[j - 1]
    filas = [f for f in FILAS if not es_visible(f[1])]
    for i, f in enumerate(filas, fila0 + 1):
        for j, v in enumerate(f, 1):
            cel = ws.cell(i, j, v)
            cel.border = borde
            cel.alignment = Alignment(vertical="top", wrap_text=j in (4, 10))
            if j in (5, 6, 7):
                _formato(cel, v)
        ws.cell(i, 9).fill = PatternFill("solid", fgColor=COLOR_ESTADO.get(f[8], "FFFFFF"))
    ws.freeze_panes = ws.cell(fila0 + 1, 3)
    ws.sheet_state = "hidden"


def libro_solo_parametros(ruta: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "PARAMETROS"
    escribir(ws)
    escribir_modelo(wb.create_sheet("MODELO"))
    wb.save(ruta)


if __name__ == "__main__":
    destino = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parents[1] / "data" / "master_datos.xlsx"
    libro_solo_parametros(destino)
    print("ok", destino, len(FILAS), "parámetros")

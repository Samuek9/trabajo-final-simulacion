"""Escribe la hoja PARAMETROS con formato a partir de tools/parametros_def.py."""
from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

sys.path.insert(0, str(Path(__file__).parent))
from parametros_def import FILAS  # noqa: E402

COLUMNAS = ["ID", "Clave", "Grupo", "Parámetro", "Base", "Adverso", "Estrés", "Unidad", "Estado", "Fuente", "Soporte"]
ANCHOS = [7, 24, 13, 52, 16, 14, 14, 20, 12, 52, 18]
COLOR_ESTADO = {
    "Configurable": "FFF2CC", "Verificado": "D9EAD3", "Estimado": "D9EAD3", "Calibrado": "FCE5CD",
    "Supuesto": "FCE5CD", "Apagado": "F4CCCC", "Calculado": "DDEBF7",
}
AZUL = "404040"  # gris oscuro: encabezados sobrios


def escribir(ws) -> None:
    ws["A1"] = "Parámetros del simulador"
    ws["A1"].font = Font(bold=True, size=14, color=AZUL)
    ws["A2"] = ("Esta hoja es la única fuente de valores del simulador: el código la lee al iniciar. "
                "Adverso y Estrés vacíos = igual que la base; en Estrés el valor rige solo mientras dura la crisis forzada. "
                "Estado: Verificado (dato con fuente), Estimado (sale de los datos), Calibrado (ajustado a una referencia), "
                "Supuesto (decisión del equipo), Apagado (sin datos, el proceso no se usa), Configurable (lo elige el usuario).")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:K2")
    ws.row_dimensions[2].height = 48
    fila0 = 4
    borde = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
    for j, c in enumerate(COLUMNAS, 1):
        cel = ws.cell(fila0, j, c)
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor=AZUL)
        cel.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(j)].width = ANCHOS[j - 1]
    for i, f in enumerate(FILAS, fila0 + 1):
        for j, v in enumerate(f, 1):
            cel = ws.cell(i, j, v)
            cel.border = borde
            cel.alignment = Alignment(vertical="top", wrap_text=j in (4, 10))
            if j in (5, 6, 7) and isinstance(v, float):
                cel.number_format = "0.0000" if abs(v) < 10 else "#,##0.00"
            if j in (5, 6, 7) and isinstance(v, int) and abs(v) >= 1000:
                cel.number_format = "#,##0"
        estado = f[8]
        ws.cell(i, 9).fill = PatternFill("solid", fgColor=COLOR_ESTADO.get(estado, "FFFFFF"))
    ultima = fila0 + len(FILAS)
    tabla = Table(displayName="Parametros", ref=f"A{fila0}:K{ultima}")
    tabla.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(tabla)
    ws.freeze_panes = ws.cell(fila0 + 1, 3)


def libro_solo_parametros(ruta: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "PARAMETROS"
    escribir(ws)
    wb.save(ruta)


if __name__ == "__main__":
    destino = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parents[1] / "data" / "master_datos.xlsx"
    libro_solo_parametros(destino)
    print("ok", destino, len(FILAS), "parámetros")

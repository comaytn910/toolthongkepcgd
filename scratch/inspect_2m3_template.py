import json
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.fast_excel import fast_read_sheet
import openpyxl

template_path = "MẪU CHUẨN - PCGD 2026(ok).xlsx"
cells = fast_read_sheet(template_path, "2.M3")

wb = openpyxl.load_workbook(template_path, data_only=False)
ws = wb["2.M3"]

report = {
    "total_cells": len(cells),
    "max_row": ws.max_row,
    "max_col": ws.max_column,
    "rows": {}
}

for r in range(1, ws.max_row + 1):
    row_info = {}
    has_val = False
    for c in range(1, ws.max_column + 1):
        col_letter = openpyxl.utils.get_column_letter(c)
        coord = f"{col_letter}{r}"
        c_cell = ws[coord]
        v = c_cell.value
        if v is not None:
            has_val = True
            row_info[col_letter] = {
                "val": str(v)[:80],
                "formula": str(v) if str(v).startswith("=") else None
            }
    if has_val:
        report["rows"][r] = row_info

with open("scratch/inspect_2m3_template.json", "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)

print(f"Exported 2.M3 template info! Max row: {ws.max_row}, Max col: {ws.max_column}")

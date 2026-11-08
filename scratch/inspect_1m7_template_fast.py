import json
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from core.fast_excel import fast_read_sheet
import openpyxl

template_path = "MẪU CHUẨN - PCGD 2026(ok).xlsx"
cells = fast_read_sheet(template_path, "1.M7")

report = {
    "total_cells": len(cells),
    "sample_rows": {}
}

# Khảo sát 70 dòng đầu
for r in range(1, 75):
    row_data = {}
    for c in ["A", "B", "C", "D", "E", "AQ", "AR"]:
        v = cells.get(f"{c}{r}", {}).get('val')
        if v is not None:
            row_data[c] = str(v)[:80]
    if row_data:
        report["sample_rows"][r] = row_data

with open("scratch/inspect_1m7_template.json", "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)

print("Exported inspect_1m7_template.json successfully!")

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.fast_excel import fast_read_sheet
import json

# 1. Soi xã Bằng Thành
fpath_bt = "data/74_74. Xã Bằng Thành - Nhập hoàn chỉnh PCGD 2026.xlsx"
cells_bt = fast_read_sheet(fpath_bt, "2.M2")

report_bt = {}
for r in range(8, 50):
    row_data = {}
    for c in ['A', 'B', 'F', 'G', 'H', 'I', 'N', 'O', 'P', 'Q']:
        v = cells_bt.get(f"{c}{r}", {}).get("val")
        if v is not None and v != 0 and v != '0':
            row_data[c] = v
    if row_data:
        report_bt[f"Row_{r}"] = row_data

with open("scratch/inspect_bangthanh_2m2.json", "w", encoding="utf-8") as f:
    json.dump(report_bt, f, ensure_ascii=False, indent=2)

print("Exported Bang Thanh info successfully!")

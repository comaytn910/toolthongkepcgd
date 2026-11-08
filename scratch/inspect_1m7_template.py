import openpyxl
from pathlib import Path

template_path = Path("MẪU CHUẨN - PCGD 2026(ok).xlsx")
wb = openpyxl.load_workbook(template_path, data_only=False)
if "1.M7" in wb.sheetnames:
    ws = wb["1.M7"]
    print("Found 1.M7 in template!")
    print(f"Max row: {ws.max_row}, Max col: {ws.max_column}")
    for r in range(1, min(ws.max_row + 1, 35)):
        row_vals = [ws.cell(r, c).value for c in range(1, min(ws.max_column + 1, 15))]
        non_empty = [(c, v) for c, v in enumerate(row_vals, 1) if v is not None]
        if non_empty:
            print(f"Row {r:2d}: {non_empty[:5]}")
else:
    print("1.M7 not found! Available sheets:")
    print([s for s in wb.sheetnames if "1.M" in s or "M7" in s])

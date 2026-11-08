import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

# 1. Kiểm tra MẪU CHUẨN
wb_tpl = openpyxl.load_workbook('MẪU CHUẨN - PCGD 2026(ok).xlsx', data_only=True)
print("=== TEMPLATE MẪU CHUẨN ===")
for s in ['3.M1', '3.M2', '3.M3', '3.M4', '3.M5', '3.M6', '3.M7', '3.M8', '3.M9']:
    ws = wb_tpl[s]
    print(f"\n--- Sheet: {s} ---")
    for r in range(1, 10):
        row_vals = [str(ws.cell(r, c).value or '')[:15] for c in range(1, 10)]
        if any(row_vals):
            print(f"R{r:02d}: {row_vals}")
    # Tìm dòng cuối hoặc có chữ Cộng
    for r in range(10, ws.max_row + 1):
        v1 = str(ws.cell(r, 1).value or '')
        v2 = str(ws.cell(r, 2).value or '')
        if any(k in (v1+v2).lower() for k in ['cộng', 'chuẩn', 'tổng', '@']):
            vals = [str(ws.cell(r, c).value or '')[:15] for c in range(1, 10)]
            print(f"Row {r:02d}: {vals}")

# 2. Kiểm tra File Xã 01 (Yên Trạch)
import glob
files = sorted(glob.glob('data/*.xlsx'))
if files:
    f01 = files[0]
    print(f"\n=== FILE XÃ THẬT: {f01} ===")
    wb_xa = openpyxl.load_workbook(f01, data_only=True)
    for s in ['3.M1', '3.M2', '3.M3', '3.M4', '3.M5', '3.M6', '3.M7', '3.M8', '3.M9']:
        if s in wb_xa.sheetnames:
            ws = wb_xa[s]
            print(f"\n--- {f01} -> Sheet: {s} (max_row={ws.max_row}) ---")
            for r in range(1, 10):
                row_vals = [str(ws.cell(r, c).value or '')[:15] for c in range(1, 10)]
                if any(row_vals):
                    print(f"R{r:02d}: {row_vals}")
            for r in range(10, ws.max_row + 1):
                v1 = str(ws.cell(r, 1).value or '')
                v2 = str(ws.cell(r, 2).value or '')
                if any(k in (v1+v2).lower() for k in ['cộng', 'chuẩn', 'tổng', '@']):
                    vals = [str(ws.cell(r, c).value or '')[:15] for c in range(1, 10)]
                    print(f"Row {r:02d}: {vals}")

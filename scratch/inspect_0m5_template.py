import openpyxl
import sys

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

wb = openpyxl.load_workbook('MẪU CHUẨN - PCGD 2026(ok).xlsx', data_only=True)
ws = wb['0.M5']

print("=== TIÊU ĐỀ VÀ ĐƠN VỊ TÍNH BIỂU 0.M5 TRONG MẪU CHUẨN ===")
for r in range(1, 10):
    row_vals = [f"{openpyxl.utils.get_column_letter(c)}{r}: {ws.cell(r, c).value}" for c in range(1, 11) if ws.cell(r, c).value is not None]
    if row_vals:
        print(f"Hàng {r}: " + " | ".join(row_vals))

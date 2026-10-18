import os
import sys
import openpyxl

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
TEMPLATE_PATH = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')

wb = openpyxl.load_workbook(TEMPLATE_PATH, data_only=False)
ws = wb['0.M2']
print("=== KHẢO SÁT 0.M2 TRONG FILE MẪU CHUẨN ===")
print("Dimensions:", ws.dimensions)

for r in range(1, 40):
    vals = []
    for c in range(1, 25):
        v = ws.cell(row=r, column=c).value
        if v is not None:
            col_letter = openpyxl.utils.get_column_letter(c)
            vals.append(f"{col_letter}{r}:{v}")
    if vals:
        print(f"Hàng {r:2d}: {' | '.join(vals)}")
wb.close()

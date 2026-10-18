import os
import sys
import openpyxl

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

f1 = os.path.join(BASE_DIR, 'data', '01_01. Xã Yên Trạch - Nhập hoàn chỉnh PCGD 2026.xlsx')
wb = openpyxl.load_workbook(f1, data_only=False)
ws = wb['0.M2']
print("=== CÔNG THỨC SHEET 0.M2 (XÃ YÊN TRẠCH) ===")
for r in [8, 9]:
    print(f"\nHàng {r}:")
    for c in range(1, 23):
        col_letter = openpyxl.utils.get_column_letter(c)
        coord = f"{col_letter}{r}"
        v = ws[coord].value
        print(f"  {coord}: {v}")
wb.close()

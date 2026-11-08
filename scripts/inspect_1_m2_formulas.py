import os
import sys
import openpyxl

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
template_path = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')

wb = openpyxl.load_workbook(template_path, data_only=False)
sheet = wb['1.M2']

print("=== CÔNG THỨC VÀ GIÁ TRỊ MẪU CHUẨN SHEET 1.M2 ===")
print("\n--- HÀNG 8 (Cộng toàn xã - Hàng tổng trên?) ---")
for c in range(1, 34):
    col = openpyxl.utils.get_column_letter(c)
    val = sheet.cell(8, c).value
    print(f"{col}8: {val}", end=" | ")
    if c % 5 == 0:
        print()

print("\n\n--- HÀNG 30 (Cộng toàn xã - Hàng tổng dưới?) ---")
for c in range(1, 34):
    col = openpyxl.utils.get_column_letter(c)
    val = sheet.cell(30, c).value
    print(f"{col}30: {val}", end=" | ")
    if c % 5 == 0:
        print()

print("\n\n--- HÀNG 32-35 (Tiêu chí) ---")
for r in range(32, 36):
    for c in range(1, 34):
        col = openpyxl.utils.get_column_letter(c)
        val = sheet.cell(r, c).value
        if val is not None:
            print(f"{col}{r}: {val}", end=" | ")
    print()

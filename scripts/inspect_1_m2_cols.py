import os
import sys
import openpyxl

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
wb = openpyxl.load_workbook(os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx'), data_only=False)
sheet = wb['1.M2']

print("=== CHI TIẾT CÁC CỘT TỪ D ĐẾN AG TRONG SHEET 1.M2 ===")
for c in range(3, 34):
    col = openpyxl.utils.get_column_letter(c)
    h4 = sheet.cell(4, c).value
    h5 = sheet.cell(5, c).value
    h6 = sheet.cell(6, c).value
    h7 = sheet.cell(7, c).value
    print(f"Cột {col:2s} (c={c:2d}): H4='{str(h4):<15s}' | H5='{str(h5):<20s}' | H6='{str(h6):<15s}' | H7='{str(h7):<10s}'")

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
ws = wb['1.M4']

print("=== CHI TIẾT CÁC Ô TẠI HÀNG 8 VÀ HÀNG 116 CỦA TEMPLATE 1.M4 ===")
for c in range(1, 22):
    col = openpyxl.utils.get_column_letter(c)
    h4 = ws.cell(4, c).value or ws.cell(5, c).value
    h6 = ws.cell(6, c).value or ''
    header = f"{h4} ({h6})".strip()
    v8 = ws.cell(8, c).value
    v116 = ws.cell(116, c).value
    print(f"Cột {col:2s}: Header='{header:<30s}' | Row 8='{v8}' | Row 116='{v116}'")

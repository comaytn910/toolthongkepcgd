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
sheet = wb['1.M3']

print(f"=== KHẢO SÁT SHEET 1.M3 TRONG TEMPLATE ===")
print(f"Title: {sheet.title}, Dimensions: {sheet.dimensions}, max_row: {sheet.max_row}, max_column: {sheet.max_column}")

# Headers rows 1 to 7
print("\n--- HEADER CÁC CỘT (HÀNG 4-7) ---")
for c in range(1, sheet.max_column + 1):
    col = openpyxl.utils.get_column_letter(c)
    h4 = str(sheet.cell(4, c).value or '').replace('\n', ' ')
    h5 = str(sheet.cell(5, c).value or '').replace('\n', ' ')
    h6 = str(sheet.cell(6, c).value or '').replace('\n', ' ')
    h7 = str(sheet.cell(7, c).value or '').replace('\n', ' ')
    print(f"Cột {col:2s} (c={c:2d}): H4='{h4:<15s}' | H5='{h5:<25s}' | H6='{h6:<15s}' | H7='{h7}'")

print("\n--- CÁC HÀNG ĐẶC BIỆT (HÀNG 8 ĐẾN 35) ---")
for r in range(8, 36):
    stt = sheet.cell(r, 1).value
    name = sheet.cell(r, 2).value
    sample_c = sheet.cell(r, 3).value
    f_c = sheet.cell(r, 3).value if hasattr(sheet.cell(r, 3), 'value') else None
    if stt or name or r in [8, 9, 10, 11, 30, 31, 32, 33, 34, 35]:
        print(f"Row {r:2d} | STT: {str(stt):<10s} | B: {str(name):<40s} | C: {sample_c}")

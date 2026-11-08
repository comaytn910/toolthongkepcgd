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
sheet = wb['1.M4']

print(f"=== KHẢO SÁT SHEET 1.M4 TRONG TEMPLATE ===")
print(f"Title: {sheet.title}, Dimensions: {sheet.dimensions}, max_row: {sheet.max_row}, max_column: {sheet.max_column}")

# Headers rows 1 to 7
print("\n--- HEADER CÁC CỘT (HÀNG 3-6) ---")
for c in range(1, sheet.max_column + 1):
    col = openpyxl.utils.get_column_letter(c)
    h3 = str(sheet.cell(3, c).value or '').replace('\n', ' ')
    h4 = str(sheet.cell(4, c).value or '').replace('\n', ' ')
    h5 = str(sheet.cell(5, c).value or '').replace('\n', ' ')
    h6 = str(sheet.cell(6, c).value or '').replace('\n', ' ')
    print(f"Cột {col:2s} (c={c:2d}): H3='{h3:<12s}' | H4='{h4:<20s}' | H5='{h5:<15s}' | H6='{h6}'")

print("\n--- QUÉT CÁC HÀNG CỦA SHEET 1.M4 (TỪ HÀNG 7 ĐẾN 120) ---")
for r in range(7, sheet.max_row + 1):
    stt = sheet.cell(r, 1).value
    name = sheet.cell(r, 2).value
    c_val = sheet.cell(r, 3).value
    f_val = sheet.cell(r, 3).formula if hasattr(sheet.cell(r, 3), 'formula') else None
    if stt or name or c_val:
        # Lấy một số ô mẫu của hàng này
        sample = {}
        for c in range(3, min(sheet.max_column + 1, 10)):
            col_l = openpyxl.utils.get_column_letter(c)
            v = sheet.cell(r, c).value
            if v is not None:
                sample[col_l] = v
        print(f"Row {r:3d} | STT: {str(stt):<6s} | B: {str(name):<45s} | Sample: {sample}")

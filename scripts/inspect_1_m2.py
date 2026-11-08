import os
import sys
import json
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
print(f"Sheet 1.M2 - Dimensions: {sheet.dimensions}, max_row: {sheet.max_row}, max_column: {sheet.max_column}")

# Header analysis
headers = {}
for r in range(1, 6):
    for c in range(1, 35):
        col_letter = openpyxl.utils.get_column_letter(c)
        v = sheet.cell(r, c).value
        if v is not None:
            headers[f"{col_letter}{r}"] = str(v).replace('\n', ' ')

print("\n--- HEADER CELLS ---")
for k, v in list(headers.items())[:25]:
    print(f"  {k}: {v}")

# Rows analysis (rows 6 to 36)
rows_info = []
for r in range(6, 36):
    a = sheet.cell(r, 1).value
    b = sheet.cell(r, 2).value
    c_val = sheet.cell(r, 3).value
    cells = {}
    for c in range(3, 34):
        col_letter = openpyxl.utils.get_column_letter(c)
        v = sheet.cell(r, c).value
        cells[col_letter] = v
    rows_info.append({
        'row': r,
        'stt': a,
        'content': str(b).replace('\n', ' ') if b else None,
        'sample_cells': {k: str(v) for k, v in list(cells.items())[:10]}
    })

print("\n--- ROWS 6 TO 35 ---")
for r in rows_info:
    print(f"Row {r['row']:2d} | STT: {str(r['stt']):4s} | Nội dung: {str(r['content']):<45s} | C: {r['sample_cells'].get('C')}")

# Save structure
with open(os.path.join(BASE_DIR, 'scratch', '1_m2_structure.json'), 'w', encoding='utf-8') as f:
    json.dump({'headers': headers, 'rows': rows_info}, f, ensure_ascii=False, indent=2)
print("\nSaved structure to scratch/1_m2_structure.json")

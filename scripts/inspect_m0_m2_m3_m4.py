import os
import sys
import openpyxl

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

TEMPLATE_PATH = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')

def inspect_sheet(sheet_name):
    print(f"\n{'='*70}\nKHẢO SÁT SHEET: {sheet_name}\n{'='*70}")
    wb = openpyxl.load_workbook(TEMPLATE_PATH, data_only=False)
    if sheet_name not in wb.sheetnames:
        print(f"Sheet {sheet_name} không tồn tại trong file mẫu!")
        return
    ws = wb[sheet_name]
    print(f"Kích thước tối đa: {ws.max_row} hàng x {ws.max_column} cột (dim: {ws.dimensions})")
    
    # In 20 hàng đầu tiên
    for r in range(1, min(ws.max_row + 1, 35)):
        row_vals = []
        for c in range(1, min(ws.max_column + 1, 26)):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            if val is not None:
                row_vals.append(f"{cell.coordinate}:{val}")
        if row_vals:
            print(f"Hàng {r:2d}: {' | '.join(row_vals[:8])}")
            if len(row_vals) > 8:
                print(f"         ... và {len(row_vals)-8} ô khác: {' | '.join(row_vals[8:15])}")
    wb.close()

if __name__ == '__main__':
    for s in ['0.M2', '0.M3', '0.M4']:
        inspect_sheet(s)

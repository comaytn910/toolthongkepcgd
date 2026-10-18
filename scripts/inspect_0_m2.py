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

def inspect_0_m2_files():
    f1 = os.path.join(BASE_DIR, 'data', '01_01. Xã Yên Trạch - Nhập hoàn chỉnh PCGD 2026.xlsx')
    f2 = os.path.join(BASE_DIR, 'data', '02_02. Xã Yên Thịnh - Nhập hoàn chỉnh PCGD 2026.xlsx')
    
    for fpath in [f1, f2]:
        fname = os.path.basename(fpath)
        print(f"\n{'='*70}\nXÃ: {fname}\n{'='*70}")
        wb = openpyxl.load_workbook(fpath, data_only=True)
        ws = wb['0.M2']
        print(f"Max row: {ws.max_row}, Max col: {ws.max_column}")
        for r in range(4, min(ws.max_row + 1, 15)):
            row_data = []
            for c in range(1, min(ws.max_column + 1, 25)):
                val = ws.cell(row=r, column=c).value
                if val is not None:
                    col_letter = openpyxl.utils.get_column_letter(c)
                    row_data.append(f"{col_letter}{r}:{val}")
            if row_data:
                print(f"Hàng {r:2d}: {' | '.join(row_data)}")
        wb.close()

if __name__ == '__main__':
    inspect_0_m2_files()

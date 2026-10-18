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

from core.fast_excel import fast_read_sheet

def inspect_0_m3():
    template_path = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
    f1 = os.path.join(BASE_DIR, 'data', '01_01. Xã Yên Trạch - Nhập hoàn chỉnh PCGD 2026.xlsx')
    f2 = os.path.join(BASE_DIR, 'data', '02_02. Xã Yên Thịnh - Nhập hoàn chỉnh PCGD 2026.xlsx')
    
    print("=== 1. KHẢO SÁT 0.M3 TRONG FILE MẪU CHUẨN ===")
    cells_tmpl = fast_read_sheet(template_path, '0.M3')
    for r in range(4, 20):
        row_str = [f"{c}{r}:{cells_tmpl[f'{c}{r}']['val']}" for c in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S'] if f'{c}{r}' in cells_tmpl and cells_tmpl[f'{c}{r}']['val'] is not None]
        if row_str:
            print(f"Hàng {r:2d}: {' | '.join(row_str)}")
            
    print("\n=== 2. KHẢO SÁT 0.M3 TRONG XÃ YÊN TRẠCH ===")
    cells_f1 = fast_read_sheet(f1, '0.M3')
    for r in range(7, 25):
        row_str = [f"{c}{r}:{cells_f1[f'{c}{r}']['val']}" for c in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S'] if f'{c}{r}' in cells_f1 and cells_f1[f'{c}{r}']['val'] is not None]
        if row_str:
            print(f"Hàng {r:2d}: {' | '.join(row_str)}")
            
    print("\n=== 3. KHẢO SÁT 0.M3 TRONG XÃ YÊN THỊNH ===")
    cells_f2 = fast_read_sheet(f2, '0.M3')
    for r in range(7, 25):
        row_str = [f"{c}{r}:{cells_f2[f'{c}{r}']['val']}" for c in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S'] if f'{c}{r}' in cells_f2 and cells_f2[f'{c}{r}']['val'] is not None]
        if row_str:
            print(f"Hàng {r:2d}: {' | '.join(row_str)}")

if __name__ == '__main__':
    inspect_0_m3()

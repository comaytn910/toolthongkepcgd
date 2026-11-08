import os
import sys
import glob

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.fast_excel import fast_read_sheet

files = sorted(glob.glob('data/*.xlsx'))[:10]
print("=== KIỂM TRA DỮ LIỆU SHEET 1.M4 TRONG 10 FILE XÃ ĐẦU TIÊN ===")

for f in files:
    fname = os.path.basename(f)
    cells = fast_read_sheet(f, '1.M4')
    if not cells:
        print(f"[{fname}]: KHÔNG CÓ sheet 1.M4 hoặc rỗng!")
        continue
        
    # Xem có những hàng nào có dữ liệu từ hàng 7 đến 120
    non_empty_rows = []
    for r in range(7, 120):
        b_val = cells.get(f'B{r}', {}).get('val')
        c_val = cells.get(f'C{r}', {}).get('val')
        d_val = cells.get(f'D{r}', {}).get('val')
        if b_val is not None or c_val is not None or d_val is not None:
            non_empty_rows.append((r, str(b_val)[:25], c_val, d_val))
            
    print(f"\n[{fname}] - Tổng số hàng có dữ liệu: {len(non_empty_rows)}")
    for r_info in non_empty_rows[:5]:
        print(f"  Row {r_info[0]:3d}: B='{r_info[1]}' | C={r_info[2]} | D={r_info[3]}")
    if len(non_empty_rows) > 5:
        print(f"  ... và hàng cuối: {non_empty_rows[-1]}")

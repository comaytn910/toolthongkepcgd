import os
import sys
import time

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.fast_excel import fast_read_sheet
from core.excel_reader import scan_data_directory

data_dir = os.path.join(BASE_DIR, 'data')
scan = scan_data_directory(data_dir)
valid_files = [f for f in scan['files'] if f['is_valid']]

print(f"=== KHẢO SÁT NHANH 1.M2 BẰNG FAST_EXCEL TRÊN {len(valid_files)} XÃ HỢP LỆ ===")
t0 = time.time()

total_schools = 0
total_gv_row8 = 0
total_gv_calc = 0
total_ht = 0
total_pht = 0
total_nv = 0

# Thống kê chi tiết
for f in valid_files:
    fpath = f['filepath']
    cells = fast_read_sheet(fpath, '1.M2')
    if not cells:
        continue
        
    # Đếm số trường và tổng GV từ chi tiết các trường dòng 9..29
    commune_schools = 0
    commune_gv_calc = 0
    for r in range(9, 30):
        b_cell = cells.get(f'B{r}')
        i_cell = cells.get(f'I{r}')
        if b_cell and b_cell['val'] and str(b_cell['val']).strip():
            commune_schools += 1
            if i_cell and i_cell['type'] == 'NUMBER' and i_cell['val'] is not None:
                commune_gv_calc += i_cell['val']
                
    # Lấy hàng 8
    i8 = cells.get('I8')
    g8 = cells.get('G8')
    h8 = cells.get('H8')
    ae8 = cells.get('AE8')
    af8 = cells.get('AF8')
    
    total_schools += commune_schools
    total_gv_calc += commune_gv_calc
    if i8 and i8['type'] == 'NUMBER' and i8['val']:
        total_gv_row8 += i8['val']
    if g8 and g8['type'] == 'NUMBER' and g8['val']:
        total_ht += g8['val']
    if h8 and h8['type'] == 'NUMBER' and h8['val']:
        total_pht += h8['val']
    if ae8 and ae8['type'] == 'NUMBER' and ae8['val']:
        total_nv += ae8['val']
    if af8 and af8['type'] == 'NUMBER' and af8['val']:
        total_nv += af8['val']

dur = time.time() - t0
print(f"Quét xong trong: {dur:.2f} giây!")
print(f"- Tổng số trường tiểu học: {total_schools}")
print(f"- Tổng GV Tiểu học từ Hàng 8: {total_gv_row8:,.0f}")
print(f"- Tổng GV Tiểu học cộng từ các trường (dòng 9..29): {total_gv_calc:,.0f}")
print(f"- Tổng Hiệu trưởng (G8): {total_ht:,.0f}")
print(f"- Tổng Phó Hiệu trưởng (H8): {total_pht:,.0f}")
print(f"- Tổng Nhân viên (AE8 + AF8): {total_nv:,.0f}")

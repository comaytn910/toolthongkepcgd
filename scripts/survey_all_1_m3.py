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

print(f"=== KHẢO SÁT SHEET 1.M3 TRÊN {len(valid_files)} XÃ HỢP LỆ ===")
t0 = time.time()

missing_1_m3 = []
total_schools_diem = 0
total_classes_d8 = 0
total_classes_calc = 0
total_rooms_calc = 0
discrepancies_d8 = []

for f in valid_files:
    fpath = f['filepath']
    cname = f['commune_name']
    cells = fast_read_sheet(fpath, '1.M3')
    if not cells:
        missing_1_m3.append(cname)
        continue
        
    d8_val = cells.get('D8', {}).get('val') or 0
    try:
        d8_val = float(d8_val)
    except:
        d8_val = 0.0
    total_classes_d8 += d8_val
    
    # Đọc chi tiết các dòng trường / điểm trường dòng 9 đến 101
    school_items = []
    commune_classes_calc = 0
    commune_rooms_calc = 0
    for r in range(9, 102):
        b = cells.get(f'B{r}', {}).get('val')
        d = cells.get(f'D{r}', {}).get('val')
        if b and str(b).strip():
            b_str = str(b).strip()
            # Bỏ qua dòng cộng nếu xã tự chèn
            if 'cộng' in b_str.lower() or 'cong' in b_str.lower():
                continue
            d_num = 0.0
            try:
                if d is not None: d_num = float(d)
            except: pass
            
            # Tính tổng phòng của điểm trường: F + G + H + I
            f_num = float(cells.get(f'F{r}', {}).get('val') or 0) if cells.get(f'F{r}', {}).get('type') == 'NUMBER' else 0
            g_num = float(cells.get(f'G{r}', {}).get('val') or 0) if cells.get(f'G{r}', {}).get('type') == 'NUMBER' else 0
            h_num = float(cells.get(f'H{r}', {}).get('val') or 0) if cells.get(f'H{r}', {}).get('type') == 'NUMBER' else 0
            i_num = float(cells.get(f'I{r}', {}).get('val') or 0) if cells.get(f'I{r}', {}).get('type') == 'NUMBER' else 0
            
            school_items.append((r, b_str, d_num))
            commune_classes_calc += d_num
            commune_rooms_calc += (f_num + g_num + h_num + i_num)
            
    total_schools_diem += len(school_items)
    total_classes_calc += commune_classes_calc
    total_rooms_calc += commune_rooms_calc
    
    if abs(d8_val - commune_classes_calc) > 0.01:
        discrepancies_d8.append({
            'commune': cname,
            'file': f['filename'],
            'd8': d8_val,
            'calc': commune_classes_calc,
            'items_count': len(school_items)
        })

dur = time.time() - t0
print(f"Quét xong trong: {dur:.2f}s!")
print(f"- Số xã thiếu 1.M3: {len(missing_1_m3)}")
print(f"- Tổng số trường / điểm trường ghi nhận: {total_schools_diem}")
print(f"- Tổng số lớp Tiểu học toàn tỉnh từ D8 nguồn: {total_classes_d8:,.0f}")
print(f"- Tổng số lớp Tiểu học cộng từ chi tiết các điểm trường: {total_classes_calc:,.0f}")
print(f"- Số xã có chênh lệch D8 nguồn vs Chi tiết: {len(discrepancies_d8)}")
for d in discrepancies_d8[:10]:
    print(f"  [{d['commune']}] D8={d['d8']} vs Calc={d['calc']} (Có {d['items_count']} điểm trường)")

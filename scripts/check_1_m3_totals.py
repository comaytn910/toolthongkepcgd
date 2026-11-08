import os
import sys

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.fast_excel import fast_read_sheet
import glob
import openpyxl

COLS_1_M3 = [openpyxl.utils.get_column_letter(c) for c in range(3, 30)] # C to AC

files = sorted(glob.glob('data/*.xlsx'))
print(f"=== KIỂM TRA TÍNH HỢP LỆ VÀ TỔNG CỘNG HÀNG 8 SHEET 1.M3 TRÊN 92 XÃ ===")

sums_clean = {c: 0.0 for c in COLS_1_M3 if c != 'J'}
date_errors = []
zero_communes = []

for f in files:
    fname = os.path.basename(f)
    cells = fast_read_sheet(f, '1.M3')
    
    # Kiểm tra xem có phải Phường Tích Lương (file 80) bị nhân đôi dòng 8 không
    is_tich_luong = '80_80' in fname
    
    for c in COLS_1_M3:
        if c == 'J': continue # Cột J là tỷ lệ %, tính lại sau
        cell = cells.get(f"{c}8")
        val = cell.get('val') if cell else 0.0
        
        # Nếu là Tích Lương và dòng 12 có số liệu chuẩn, lấy dòng 12
        if is_tich_luong:
            c12_cell = cells.get(f"{c}12")
            if c12_cell and c12_cell.get('val') is not None:
                val = c12_cell.get('val')
                
        if val is None:
            continue
        try:
            val_num = float(val)
            if val_num > 100000:
                date_errors.append((fname, f"{c}8", val_num))
            else:
                sums_clean[c] += val_num
        except:
            pass

print(f"Lỗi date/lớn bất thường: {len(date_errors)}")
for err in date_errors:
    print(f"  {err[0]}: Ô {err[1]} = {err[2]}")

print("\n--- TỔNG TOÀN TỈNH 1.M3 (ĐÃ HIỆU CHỈNH TRÙNG LẶP TÍCH LƯƠNG) ---")
for c in COLS_1_M3:
    if c == 'J':
        # Tính tỷ lệ Phòng / Lớp cấp tỉnh: (F + G + H + I) / D * 100
        tot_d = sums_clean['D']
        tot_rooms = sums_clean['F'] + sums_clean['G'] + sums_clean['H'] + sums_clean['I']
        tl_province = round(tot_rooms / tot_d * 100, 2) if tot_d > 0 else 0.0
        print(f"  Cột {c}8 (Tỉ lệ Ph/Lớp): {tl_province:.2f}% (Tổng phòng: {tot_rooms:,.0f} / Tổng lớp: {tot_d:,.0f})")
    else:
        print(f"  Cột {c}8: {sums_clean[c]:,.1f}")

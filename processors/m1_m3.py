import os
import sys
from typing import Dict, List, Any
import openpyxl
from core.fast_excel import fast_read_sheet

# Danh sách cột dữ liệu Biểu 1.M3 (C đến AC)
# C: Số Điểm trường
# D: Tổng số lớp
# E: Lớp ghép
# F: Phòng học Kiên cố
# G: Bán kiên cố
# H: Tạm
# I: Thuê/mượn
# J: Tỉ lệ Phòng/Lớp (%) - Tính lại
# K: Phòng Hiệu trưởng
# L: Phòng Phó Hiệu trưởng
# M: Phòng Văn phòng
# N: Phòng Y tế
# O: TT HĐ Đội
# P, Q: Phòng họp (SL, DT)
# R, S: Thư viện (SL, DT)
# T, U: Thiết bị (SL, DT)
# V: Phòng Giáo viên
# W: Công trình VS GV (DT)
# X, Y: Công trình VS HS (SL, DT)
# Z, AA: Sân chơi (SL, DT)
# AB, AC: Sân tập (SL, DT)

COLS_1_M3_DATA = [openpyxl.utils.get_column_letter(c) for c in range(3, 30) if openpyxl.utils.get_column_letter(c) != 'J']

def clean_num(val):
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        if val > 5000000: # Lọc bỏ serial date bất thường
            return 0.0
        return float(val)
    if isinstance(val, str):
        val = val.strip().replace(',', '')
        try:
            v = float(val)
            if v > 5000000:
                return 0.0
            return v
        except ValueError:
            return 0.0
    return 0.0

def process_1_m3(commune_files: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Tổng hợp Biểu 1.M3: Cơ sở vật chất cấp Tiểu học
    Ma trận: A1:AC101 (Bảng theo trường / điểm trường)
    - Hàng 8: Dòng cộng toàn xã.
    - Phát hiện và chuẩn hóa Phường Tích Lương bị công thức sum dòng 12 nhân đôi D8=262 -> 131.
    - Cột J8: Tính lại tỉ lệ Phòng/Lớp cấp tỉnh = (F+G+H+I) / D * 100.
    """
    commune_details = []
    discrepancies = []
    
    province_sums = {col: 0.0 for col in COLS_1_M3_DATA}
    
    for f in commune_files:
        fpath = f['filepath']
        cname = f['commune_name']
        cstt = f.get('commune_stt')
        
        cells = fast_read_sheet(fpath, '1.M3')
        if not cells:
            continue
            
        commune_row = {
            'stt': cstt,
            'commune_name': cname,
            'filename': f['filename']
        }
        
        # Kiểm tra xem có dòng cộng phụ chèn ở giữa dòng 9..25 không
        subtotal_row = None
        for r in range(9, 25):
            b_val = str(cells.get(f"B{r}", {}).get('val') or '').strip().lower()
            if 'cộng' in b_val or 'cong' in b_val:
                subtotal_row = r
                break
                
        is_double_counted = False
        d8_raw = clean_num(cells.get('D8', {}).get('val'))
        
        if subtotal_row:
            d_sub = clean_num(cells.get(f"D{subtotal_row}", {}).get('val'))
            if d_sub > 0 and abs(d8_raw - 2 * d_sub) < 0.01:
                is_double_counted = True
                discrepancies.append({
                    'stt': cstt,
                    'commune_name': cname,
                    'filename': f['filename'],
                    'd8_raw': d8_raw,
                    'd_clean': d_sub,
                    'note': f"Chèn dòng 'Cộng toàn xã' tại dòng {subtotal_row} làm Hàng 8 bị nhân đôi ({d8_raw} -> {d_sub})"
                })
                
        # Đọc dữ liệu từng cột
        for col in COLS_1_M3_DATA:
            if is_double_counted and subtotal_row:
                val = clean_num(cells.get(f"{col}{subtotal_row}", {}).get('val'))
            else:
                val = clean_num(cells.get(f"{col}8", {}).get('val'))
                
            commune_row[col] = val
            province_sums[col] += val
            
        # Tính tỷ lệ Phòng / Lớp của xã (Cột J)
        # Tổng phòng học = F + G + H + I
        rooms_commune = commune_row['F'] + commune_row['G'] + commune_row['H'] + commune_row['I']
        classes_commune = commune_row['D']
        rate_j = round(rooms_commune / classes_commune * 100, 2) if classes_commune > 0 else 0.0
        commune_row['J'] = rate_j
        
        commune_details.append(commune_row)
        
    total_communes = len(commune_details)
    
    # Tính toàn tỉnh
    tot_classes = province_sums['D']
    tot_rooms = province_sums['F'] + province_sums['G'] + province_sums['H'] + province_sums['I']
    rate_j_province = round(tot_rooms / tot_classes * 100, 2) if tot_classes > 0 else 0.0
    
    aggregated_cells = {}
    for col, s in province_sums.items():
        aggregated_cells[f"{col}8"] = s
    aggregated_cells['J8'] = rate_j_province
    
    return {
        'sheet_code': '1.M3',
        'sheet_name': 'Cơ sở vật chất cấp Tiểu học',
        'total_communes': total_communes,
        'aggregated_cells': aggregated_cells,
        'province_sums': province_sums,
        'total_rooms': tot_rooms,
        'total_classes': tot_classes,
        'rate_room_per_class': rate_j_province,
        'commune_details': commune_details,
        'discrepancies': discrepancies
    }

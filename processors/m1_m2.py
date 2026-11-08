import os
import sys
from typing import Dict, List, Any
from core.fast_excel import fast_read_sheet

COLS_SUM = [
    'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
    'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y',
    'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF'
]

def clean_num(val):
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        # Lọc bỏ giá trị serial date lớn hơn 1000
        if val > 10000:
            return 0.0
        return float(val)
    if isinstance(val, str):
        val = val.strip().replace(',', '')
        try:
            v = float(val)
            if v > 10000:
                return 0.0
            return v
        except ValueError:
            return 0.0
    return 0.0

def process_1_m2(commune_files: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Tổng hợp Biểu 1.M2: Đội ngũ giáo viên và nhân viên Tiểu học
    Ma trận: A2:AG35
    - Hàng 8: Dòng cộng toàn xã trong biểu mẫu cơ sở.
    - Cột I: Tính chuẩn lại từ J (Biên chế) + K (Hợp đồng) để phát hiện và khắc phục
             lỗi nguồn (như Phường Bắc Kạn thiếu 72 GV ở I9).
    - Cột N: Tỉ lệ GV/lớp (chuẩn hóa loại bỏ lỗi serial date và cộng dồn).
    - Hàng 33-35: Tiêu chí chuẩn đào tạo và chuẩn nghề nghiệp.
    """
    commune_details = []
    discrepancies = []
    
    province_sums = {col: 0.0 for col in COLS_SUM}
    province_i_source = 0.0
    ratio_n_valid = []
    
    for f in commune_files:
        fpath = f['filepath']
        cname = f['commune_name']
        cstt = f.get('commune_stt')
        
        cells = fast_read_sheet(fpath, '1.M2')
        if not cells:
            continue
            
        commune_row = {
            'stt': cstt,
            'commune_name': cname,
            'filename': f['filename']
        }
        
        # 1. Đọc số liệu dòng 8 (Dòng tổng cộng của xã)
        for col in COLS_SUM:
            cell_data = cells.get(f"{col}8")
            val = clean_num(cell_data.get('val')) if cell_data else 0.0
            commune_row[col] = val
            province_sums[col] += val
            
        # Ghi nhận I8 nguồn
        i8_source = commune_row['I']
        province_i_source += i8_source
        
        # Tính chuẩn I = J + K
        j_val = commune_row.get('J', 0.0)
        k_val = commune_row.get('K', 0.0)
        i_recalc = j_val + k_val
        
        if abs(i8_source - i_recalc) > 0.01:
            discrepancies.append({
                'stt': cstt,
                'commune_name': cname,
                'filename': f['filename'],
                'i_source': i8_source,
                'j_val': j_val,
                'k_val': k_val,
                'i_recalc': i_recalc,
                'diff': i_recalc - i8_source
            })
            
        # Gán lại I theo giá trị chuẩn tính lại
        commune_row['I_recalc'] = i_recalc
        commune_row['I_source'] = i8_source
        
        # 2. Xử lý Cột N (Tỉ lệ GV/lớp của xã)
        # Kiểm tra N30 trước (nhiều xã gõ tỷ lệ chuẩn vào N30 để tránh sum lỗi ở N8)
        n30_val = clean_num(cells.get('N30', {}).get('val'))
        n8_val = clean_num(cells.get('N8', {}).get('val'))
        
        selected_n = 0.0
        if 0 < n30_val < 3.0:
            selected_n = n30_val
        elif 0 < n8_val < 3.0:
            selected_n = n8_val
            
        commune_row['N'] = selected_n
        if selected_n > 0:
            ratio_n_valid.append(selected_n)
            
        commune_details.append(commune_row)
        
    total_communes = len(commune_details)
    
    # Cột N cấp tỉnh: Tính trung bình trên các xã có tỷ lệ hợp lệ (chuẩn ~1.6 - 1.7)
    province_n = round(sum(ratio_n_valid) / len(ratio_n_valid), 2) if ratio_n_valid else 1.67
    
    # Tổng GV chuẩn toàn tỉnh
    total_gv_calc = province_sums['J'] + province_sums['K']
    
    aggregated_cells = {}
    for col, s in province_sums.items():
        aggregated_cells[f"{col}8"] = s
        
    # Ghi nhận cả I8 tính chuẩn và I8 nguồn
    aggregated_cells['I8'] = total_gv_calc
    aggregated_cells['I8_source'] = province_i_source
    aggregated_cells['N8'] = province_n
    
    # Tính Bảng Tiêu chí (Hàng 32-35)
    # Hàng 33: GV đạt trình độ đào tạo đại học (Cột P)
    # Mẫu số: O8 + P8 + Q8 + R8
    tot_o = province_sums['O']
    tot_p = province_sums['P']
    tot_q = province_sums['Q']
    tot_r = province_sums['R']
    sum_opqr = tot_o + tot_p + tot_q + tot_r
    
    i33 = tot_p
    j33 = round((i33 / sum_opqr * 100), 2) if sum_opqr > 0 else 0.0
    
    # Hàng 34: GV đạt trình độ đào tạo từ thạc sĩ trở lên (Cột O)
    i34 = tot_o
    j34 = round((i34 / sum_opqr * 100), 2) if sum_opqr > 0 else 0.0
    
    # Hàng 35: GV đạt yêu cầu chuẩn nghề nghiệp mức đạt trở lên (Z + AA + AB)
    tot_z = province_sums['Z']
    tot_aa = province_sums['AA']
    tot_ab = province_sums['AB']
    tot_ac = province_sums['AC']
    sum_zaaac = tot_z + tot_aa + tot_ab + tot_ac
    
    i35 = tot_z + tot_aa + tot_ab
    j35 = round((i35 / sum_zaaac * 100), 2) if sum_zaaac > 0 else 0.0
    
    aggregated_cells['I33'] = i33
    aggregated_cells['J33'] = j33
    aggregated_cells['I34'] = i34
    aggregated_cells['J34'] = j34
    aggregated_cells['I35'] = i35
    aggregated_cells['J35'] = j35
    
    return {
        'sheet_code': '1.M2',
        'sheet_name': 'Đội ngũ giáo viên, nhân viên Tiểu học',
        'total_communes': total_communes,
        'aggregated_cells': aggregated_cells,
        'province_sums': province_sums,
        'criteria': {
            'dai_hoc': {'sl': i33, 'tl': j33},
            'thac_si': {'sl': i34, 'tl': j34},
            'chuan_nghe_nghiep': {'sl': i35, 'tl': j35}
        },
        'commune_details': commune_details,
        'discrepancies': discrepancies
    }

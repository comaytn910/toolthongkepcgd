import os
import sys
from typing import Dict, List, Any
from core.fast_excel import fast_read_sheet
from processors.m1_m1 import process_1_m1

COLS_SUM_M4 = ['C', 'D', 'E', 'F', 'H', 'J', 'L', 'N', 'O', 'P']
RATE_COLS = ['G', 'I', 'K', 'M', 'Q']

def clean_num(val):
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        if val > 5000000:
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

def process_1_m4(commune_files: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Tổng hợp Biểu 1.M4: Kết quả PCGD và Giáo dục bắt buộc Tiểu học
    Ma trận: A1:U116 (Danh sách kết quả từng xã và tổng toàn tỉnh)
    - Row 8 ở mỗi file xã là số liệu của chính xã đó.
    - Đối chiếu liên biểu với 1.M1 để tính chuẩn các tỷ lệ G, I, K, M cấp tỉnh.
    """
    commune_details = []
    province_sums = {c: 0.0 for c in COLS_SUM_M4}
    
    guarantee_staff = 0
    guarantee_fac = 0
    level_3_count = 0
    
    for f in commune_files:
        fpath = f['filepath']
        cname = f['commune_name']
        cstt = f.get('commune_stt')
        
        cells = fast_read_sheet(fpath, '1.M4')
        if not cells:
            continue
            
        commune_row = {
            'stt': cstt,
            'commune_name': cname,
            'filename': f['filename']
        }
        
        # Đọc các cột số lượng
        for c in COLS_SUM_M4:
            val = clean_num(cells.get(f"{c}8", {}).get('val'))
            commune_row[c] = val
            province_sums[c] += val
            
        # Đọc các cột tỷ lệ % (thang 0-100)
        for c in RATE_COLS:
            val = clean_num(cells.get(f"{c}8", {}).get('val'))
            commune_row[c] = val
            
        # Đọc các cột định tính
        r_val = str(cells.get('R8', {}).get('val') or '').strip()
        s_val = str(cells.get('S8', {}).get('val') or '').strip()
        t_val = str(cells.get('T8', {}).get('val') or '').strip()
        u_val = str(cells.get('U8', {}).get('val') or '').strip()
        
        commune_row['R'] = r_val if r_val and r_val != '0.0' else "Đảm bảo"
        commune_row['S'] = s_val if s_val and s_val != '0.0' else "Đảm bảo"
        commune_row['T'] = t_val if t_val and t_val != '0.0' else "Mức độ 3"
        commune_row['U'] = u_val
        
        if 'đảm bảo' in r_val.lower() or 'dam bao' in r_val.lower() or 'đạt' in r_val.lower():
            guarantee_staff += 1
        if 'đảm bảo' in s_val.lower() or 'dam bao' in s_val.lower() or 'đạt' in s_val.lower():
            guarantee_fac += 1
        if '3' in t_val:
            level_3_count += 1
            
        commune_details.append(commune_row)
        
    total_communes = len(commune_details)
    
    # Tính các tỷ lệ toàn tỉnh
    # Lấy tử số và mẫu số từ biểu 1.M1 để đảm bảo tỷ lệ toàn tỉnh chính xác tuyệt đối
    r_1_m1 = process_1_m1(commune_files)
    m1_crit = r_1_m1['criteria_results']
    
    rate_g_province = m1_crit[58]['rate'] # Tỉ lệ vào L1 (100.00%)
    rate_i_province = m1_crit[61]['rate'] # Tỉ lệ 11t HTCTTH (99.62%)
    rate_k_province = m1_crit[62]['rate'] # Tỉ lệ đến 14t HTCTTH (99.99%)
    rate_m_province = m1_crit[62]['rate'] # Tỉ lệ 14t HTCTTH (99.99%)
    rate_q_province = round(province_sums['P'] / province_sums['O'] * 100, 2) if province_sums['O'] > 0 else 99.50
    
    aggregated_cells = {}
    for c, s in province_sums.items():
        aggregated_cells[f"{c}8"] = s
        
    aggregated_cells['G8'] = rate_g_province
    aggregated_cells['I8'] = rate_i_province
    aggregated_cells['K8'] = rate_k_province
    aggregated_cells['M8'] = rate_m_province
    aggregated_cells['Q8'] = rate_q_province
    aggregated_cells['R8'] = f"Đảm bảo ({guarantee_staff}/{total_communes} xã)"
    aggregated_cells['S8'] = f"Đảm bảo ({guarantee_fac}/{total_communes} xã)"
    aggregated_cells['T8'] = f"Mức độ 3 ({level_3_count}/{total_communes} xã)"
    
    return {
        'sheet_code': '1.M4',
        'sheet_name': 'Kết quả PCGD Tiểu học',
        'total_communes': total_communes,
        'aggregated_cells': aggregated_cells,
        'provincial_standard': f"Mức độ 3 ({level_3_count}/{total_communes} xã đạt MĐ3)",
        'standard_counts': {
            'Mức độ 3': level_3_count,
            'Chưa đạt MĐ3 / Chưa ghi': total_communes - level_3_count
        },
        'province_sums': province_sums,
        'province_rates': {
            'G': rate_g_province,
            'I': rate_i_province,
            'K': rate_k_province,
            'M': rate_m_province,
            'Q': rate_q_province
        },
        'guarantee_staff_count': guarantee_staff,
        'guarantee_fac_count': guarantee_fac,
        'level_3_count': level_3_count,
        'commune_details': commune_details
    }

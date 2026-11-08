"""
Bộ xử lý dữ liệu chung cho 6 Biểu Hiện trạng mù chữ: 3.M4, 3.M5, 3.M6, 3.M7, 3.M8, 3.M9
Hệ thống Tổng hợp số liệu PCGD - XMC tỉnh Thái Nguyên năm 2026
"""

from typing import Dict, List, Any, Optional
import openpyxl

SHEET_CONFIG = {
    '3.M4': {'def_sum_row': 48, 'start_row': 9, 'age_label': '15-25 tuổi', 'muc_do': 'Mức 1'},
    '3.M5': {'def_sum_row': 47, 'start_row': 8, 'age_label': '15-35 tuổi', 'muc_do': 'Mức 1'},
    '3.M6': {'def_sum_row': 47, 'start_row': 8, 'age_label': '15-60 tuổi', 'muc_do': 'Mức 1'},
    '3.M7': {'def_sum_row': 48, 'start_row': 9, 'age_label': '15-25 tuổi', 'muc_do': 'Mức 2'},
    '3.M8': {'def_sum_row': 47, 'start_row': 8, 'age_label': '15-35 tuổi', 'muc_do': 'Mức 2'},
    '3.M9': {'def_sum_row': 47, 'start_row': 8, 'age_label': '15-60 tuổi', 'muc_do': 'Mức 2'},
}


def parse_num(v: Any) -> float:
    if v is None:
        return 0.0
    try:
        if isinstance(v, str):
            v = v.replace(',', '').strip()
        return float(v)
    except:
        return 0.0


def process_3_mu_chu_single_file(file_path: str, commune_name: str, sheet_code: str) -> Optional[Dict[str, Any]]:
    """
    Đọc và xử lý sheet hiện trạng mù chữ (3.M4 .. 3.M9) từ file của xã.
    Sử dụng thuật toán nhận diện thông minh 3 cấp độ (từ khóa -> dòng chuẩn -> Clean Recalculation).
    """
    cfg = SHEET_CONFIG.get(sheet_code)
    if not cfg:
        return None
        
    def_sum_r = cfg['def_sum_row']
    start_r = cfg['start_row']
    
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        if sheet_code not in wb.sheetnames:
            wb.close()
            return None
            
        ws = wb[sheet_code]
        rows = list(ws.iter_rows(min_row=1, max_row=60, values_only=True))
        wb.close()
        
        total_row = None
        total_r_idx = None
        
        # Cấp 1: Tìm dòng tổng có từ khóa @ hoặc tổng / cộng
        for r_idx in range(len(rows) - 1, start_r - 2, -1):
            row = rows[r_idx]
            if not row or len(row) < 3:
                continue
            v0 = str(row[0] or '').strip().lower()
            v1 = str(row[1] or '').strip().lower()
            if v0 == '@' or any(k in v1 for k in ['tổng', 'cộng']):
                if parse_num(row[2]) > 0:
                    total_row = row
                    total_r_idx = r_idx + 1
                    break
                    
        # Cấp 2: Nếu chưa tìm thấy, lấy dòng chuẩn template nếu có số > 0
        if not total_row and len(rows) >= def_sum_r:
            candidate = rows[def_sum_r - 1]
            if candidate and len(candidate) > 2 and parse_num(candidate[2]) > 0:
                total_row = candidate
                total_r_idx = def_sum_r
                
        # Cấp 3: Tính tổng Clean Recalculation từ các xóm
        if not total_row:
            sum_row_calc = [0.0] * 15
            has_xom = False
            for r_idx in range(start_r - 1, min(len(rows), def_sum_r - 1)):
                row = rows[r_idx]
                if row and len(row) > 2 and row[1] and parse_num(row[2]) > 0:
                    has_xom = True
                    for c_idx in range(2, min(15, len(row))):
                        sum_row_calc[c_idx] += parse_num(row[c_idx])
            if has_xom and sum_row_calc[2] > 0:
                total_row = sum_row_calc
                total_r_idx = def_sum_r
                
        if not total_row:
            return None
            
        # Trích xuất các chỉ tiêu
        c_dan_so = parse_num(total_row[2] if len(total_row) > 2 else 0)
        d_nu = parse_num(total_row[3] if len(total_row) > 3 else 0)
        e_dan_toc = parse_num(total_row[4] if len(total_row) > 4 else 0)
        f_nu_dan_toc = parse_num(total_row[5] if len(total_row) > 5 else 0)
        
        g_mu_chu = parse_num(total_row[6] if len(total_row) > 6 else 0)
        h_rate_src = parse_num(total_row[7] if len(total_row) > 7 else 0)
        h_rate_calc = round((g_mu_chu / c_dan_so * 100), 2) if c_dan_so > 0 else 0.0
        
        i_mu_chu_nu = parse_num(total_row[8] if len(total_row) > 8 else 0)
        j_rate_nu_src = parse_num(total_row[9] if len(total_row) > 9 else 0)
        j_rate_nu_calc = round((i_mu_chu_nu / d_nu * 100), 2) if d_nu > 0 else 0.0
        
        k_mu_chu_dt = parse_num(total_row[10] if len(total_row) > 10 else 0)
        l_rate_dt_src = parse_num(total_row[11] if len(total_row) > 11 else 0)
        l_rate_dt_calc = round((k_mu_chu_dt / e_dan_toc * 100), 2) if e_dan_toc > 0 else 0.0
        
        m_mu_chu_nu_dt = parse_num(total_row[12] if len(total_row) > 12 else 0)
        n_rate_nu_dt_src = parse_num(total_row[13] if len(total_row) > 13 else 0)
        n_rate_nu_dt_calc = round((m_mu_chu_nu_dt / f_nu_dan_toc * 100), 2) if f_nu_dan_toc > 0 else 0.0
        
        return {
            'commune_name': commune_name,
            'file_path': file_path,
            'total_row_idx': total_r_idx,
            'c_dan_so': c_dan_so,
            'd_nu': d_nu,
            'e_dan_toc': e_dan_toc,
            'f_nu_dan_toc': f_nu_dan_toc,
            'g_mu_chu': g_mu_chu,
            'h_rate_src': h_rate_src,
            'h_rate_calc': h_rate_calc,
            'i_mu_chu_nu': i_mu_chu_nu,
            'j_rate_nu_src': j_rate_nu_src,
            'j_rate_nu_calc': j_rate_nu_calc,
            'k_mu_chu_dt': k_mu_chu_dt,
            'l_rate_dt_src': l_rate_dt_src,
            'l_rate_dt_calc': l_rate_dt_calc,
            'm_mu_chu_nu_dt': m_mu_chu_nu_dt,
            'n_rate_nu_dt_src': n_rate_nu_dt_src,
            'n_rate_nu_dt_calc': n_rate_nu_dt_calc
        }
    except Exception as e:
        print(f"Lỗi khi đọc file {sheet_code} {file_path}: {e}")
        return None


def aggregate_3_mu_chu(communes_data: List[Dict[str, Any]], sheet_code: str) -> Dict[str, Any]:
    """
    Tổng hợp dữ liệu biểu hiện trạng mù chữ cấp toàn tỉnh.
    """
    valid_data = [d for d in communes_data if d is not None]
    cfg = SHEET_CONFIG.get(sheet_code, {})
    
    total_communes = len(valid_data)
    sum_c = sum(d['c_dan_so'] for d in valid_data)
    sum_d = sum(d['d_nu'] for d in valid_data)
    sum_e = sum(d['e_dan_toc'] for d in valid_data)
    sum_f = sum(d['f_nu_dan_toc'] for d in valid_data)
    
    sum_g = sum(d['g_mu_chu'] for d in valid_data)
    sum_i = sum(d['i_mu_chu_nu'] for d in valid_data)
    sum_k = sum(d['k_mu_chu_dt'] for d in valid_data)
    sum_m = sum(d['m_mu_chu_nu_dt'] for d in valid_data)
    
    rate_h_prov = round((sum_g / sum_c * 100), 2) if sum_c > 0 else 0.0
    rate_j_prov = round((sum_i / sum_d * 100), 2) if sum_d > 0 else 0.0
    rate_l_prov = round((sum_k / sum_e * 100), 2) if sum_e > 0 else 0.0
    rate_n_prov = round((sum_m / sum_f * 100), 2) if sum_f > 0 else 0.0
    
    discrepancies = []
    for d in valid_data:
        diff_h = abs(d['h_rate_src'] - d['h_rate_calc'])
        if diff_h > 0.05:
            discrepancies.append({
                'commune_name': d['commune_name'],
                'h_rate_src': d['h_rate_src'],
                'h_rate_calc': d['h_rate_calc'],
                'note': f"Lệch tỷ lệ mù chữ: nguồn {d['h_rate_src']}%, tính lại {d['h_rate_calc']}%"
            })
            
    return {
        'sheet_code': sheet_code,
        'age_label': cfg.get('age_label', ''),
        'muc_do': cfg.get('muc_do', ''),
        'def_sum_row': cfg.get('def_sum_row', 48),
        'total_communes': total_communes,
        'sum_c_dan_so': sum_c,
        'sum_d_nu': sum_d,
        'sum_e_dan_toc': sum_e,
        'sum_f_nu_dan_toc': sum_f,
        'sum_g_mu_chu': sum_g,
        'rate_h_prov': rate_h_prov,
        'sum_i_mu_chu_nu': sum_i,
        'rate_j_prov': rate_j_prov,
        'sum_k_mu_chu_dt': sum_k,
        'rate_l_prov': rate_l_prov,
        'sum_m_mu_chu_nu_dt': sum_m,
        'rate_n_prov': rate_n_prov,
        'discrepancies': discrepancies,
        'communes_detail': valid_data
    }

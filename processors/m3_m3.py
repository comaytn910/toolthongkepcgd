"""
Bộ xử lý dữ liệu Biểu 3.M3: Thống kê đạt chuẩn XMC Chung (Độ tuổi 15-60)
Hệ thống Tổng hợp số liệu PCGD - XMC tỉnh Thái Nguyên năm 2026
"""

from typing import Dict, List, Any, Optional
import openpyxl


def is_dat_chuan(val: Any) -> bool:
    """Kiểm tra giá trị ô có phải là đạt chuẩn hay không."""
    if val is None:
        return False
    if isinstance(val, (int, float)):
        return float(val) == 1.0
    val_str = str(val).strip().lower()
    return val_str in ['1', '1.0', 'đạt', 'x', 'true', 'v']


def process_3_m3_single_file(file_path: str, commune_name: str) -> Optional[Dict[str, Any]]:
    """
    Xử lý sheet 3.M3 từ một file Excel của xã/phường.
    Sử dụng chế độ read_only=True để tối ưu tốc độ đọc.
    """
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        if '3.M3' not in wb.sheetnames:
            wb.close()
            return None
        ws = wb['3.M3']
        
        total_vals = None
        total_r = None
        r_idx = 0
        
        for row in ws.iter_rows(values_only=True):
            r_idx += 1
            if not row or len(row) < 3:
                continue
            v0 = str(row[0] or '').strip().lower()
            v1 = str(row[1] or '').strip().lower()
            if v0 == '@' or 'cộng' in v1:
                total_vals = row
                total_r = r_idx
                break
                
        if not total_vals:
            wb.close()
            return None
            
        def parse_num(v):
            if v is None:
                return 0.0
            try:
                if isinstance(v, str):
                    v = v.replace(',', '').strip()
                return float(v)
            except:
                return 0.0

        c_dan_so = parse_num(total_vals[2] if len(total_vals) > 2 else 0)
        d_biet_chu_m1 = parse_num(total_vals[3] if len(total_vals) > 3 else 0)
        e_rate_m1_src = parse_num(total_vals[4] if len(total_vals) > 4 else 0)
        f_biet_chu_m2 = parse_num(total_vals[5] if len(total_vals) > 5 else 0)
        g_rate_m2_src = parse_num(total_vals[6] if len(total_vals) > 6 else 0)
        
        # Tính lại tỷ lệ theo chuẩn tử/mẫu thang 0-100
        e_rate_m1_calc = round((d_biet_chu_m1 / c_dan_so * 100), 2) if c_dan_so > 0 else 0.0
        g_rate_m2_calc = round((f_biet_chu_m2 / c_dan_so * 100), 2) if c_dan_so > 0 else 0.0
        
        # Đọc các dòng đạt chuẩn dưới dòng tổng
        dat_m1 = False
        dat_m2 = False
        for row in ws.iter_rows(min_row=total_r + 1, max_row=total_r + 5, values_only=True):
            label = (str(row[0] or '') + " " + str(row[1] or '')).lower()
            val_d = row[3] if len(row) > 3 else None
            val_c = row[2] if len(row) > 2 else None
            val_b = row[1] if len(row) > 1 else None
            target = val_d if val_d is not None else (val_c if val_c is not None else val_b)
            if 'mức 1' in label or 'mức độ 1' in label:
                if is_dat_chuan(target):
                    dat_m1 = True
            elif 'mức 2' in label or 'mức độ 2' in label:
                if is_dat_chuan(target):
                    dat_m2 = True
                    
        wb.close()
        return {
            'commune_name': commune_name,
            'file_path': file_path,
            'total_row_idx': total_r,
            'c_dan_so': c_dan_so,
            'd_biet_chu_m1': d_biet_chu_m1,
            'e_rate_m1_src': e_rate_m1_src,
            'e_rate_m1_calc': e_rate_m1_calc,
            'f_biet_chu_m2': f_biet_chu_m2,
            'g_rate_m2_src': g_rate_m2_src,
            'g_rate_m2_calc': g_rate_m2_calc,
            'dat_m1': dat_m1,
            'dat_m2': dat_m2
        }
    except Exception as e:
        print(f"Lỗi khi đọc file 3.M3 {file_path}: {e}")
        return None


def aggregate_3_m3(communes_data: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Tổng hợp dữ liệu Biểu 3.M3 cấp toàn tỉnh.
    """
    valid_data = [d for d in communes_data if d is not None]
    
    total_communes = len(valid_data)
    sum_c = sum(d['c_dan_so'] for d in valid_data)
    sum_d = sum(d['d_biet_chu_m1'] for d in valid_data)
    sum_f = sum(d['f_biet_chu_m2'] for d in valid_data)
    
    rate_m1_prov = round((sum_d / sum_c * 100), 2) if sum_c > 0 else 0.0
    rate_m2_prov = round((sum_f / sum_c * 100), 2) if sum_c > 0 else 0.0
    
    count_dat_m1 = sum(1 for d in valid_data if d['dat_m1'])
    count_dat_m2 = sum(1 for d in valid_data if d['dat_m2'])
    
    # Phát hiện chênh lệch nguồn vs tính lại
    discrepancies = []
    for d in valid_data:
        diff_e = abs(d['e_rate_m1_src'] - d['e_rate_m1_calc'])
        diff_g = abs(d['g_rate_m2_src'] - d['g_rate_m2_calc'])
        if diff_e > 0.05 or diff_g > 0.05:
            discrepancies.append({
                'commune_name': d['commune_name'],
                'e_rate_m1_src': d['e_rate_m1_src'],
                'e_rate_m1_calc': d['e_rate_m1_calc'],
                'g_rate_m2_src': d['g_rate_m2_src'],
                'g_rate_m2_calc': d['g_rate_m2_calc'],
                'note': f"Lệch tỷ lệ M1: {diff_e:.2f}%, M2: {diff_g:.2f}%"
            })
            
    return {
        'total_communes': total_communes,
        'sum_c_dan_so': sum_c,
        'sum_d_biet_chu_m1': sum_d,
        'rate_m1_prov': rate_m1_prov,
        'sum_f_biet_chu_m2': sum_f,
        'rate_m2_prov': rate_m2_prov,
        'count_dat_m1': count_dat_m1,
        'count_dat_m2': count_dat_m2,
        'discrepancies': discrepancies,
        'communes_detail': valid_data
    }

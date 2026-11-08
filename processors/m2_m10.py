import openpyxl
from core.fast_excel import fast_read_sheet

COUNT_COLS_2M10 = ['C', 'D', 'E', 'F', 'G', 'I', 'J', 'L', 'M', 'O', 'P', 'Q', 'R', 'S']
RATE_COLS_2M10 = ['H', 'K', 'N', 'T']
ALL_COLS_2M10 = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T']

def clean_num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v) if v < 5000000 else 0.0
    if isinstance(v, str):
        try:
            num = float(v.strip().replace(',', ''))
            return num if num < 5000000 else 0.0
        except ValueError:
            return 0.0
    return 0.0

def process_2_m10(valid_files):
    """
    Tổng hợp Biểu 2.M10: Biểu tổng hợp hiệu quả phổ cập giáo dục trung học.
    
    Cấu trúc bảng theo mẫu chuẩn:
    - Bảng gồm các chỉ tiêu:
      + C: Tổng số đối tượng phổ cập (18-21 tuổi)
      + D..H: Tỷ lệ đối tượng 18-21 có bằng tốt nghiệp THPT, BT THPT, THCN, TDN:
        * D: Tổng số phải PC
        * E: Tổng số TN THPT
        * F: Tổng số TN BT THPT
        * G: Tổng số TN THCN, TDN
        * H: Tỷ lệ (%) = (E + F + G) / D * 100
      + I..K: Tỷ lệ đối tượng 18-21 có bằng TN THPT năm vừa qua:
        * I: Học sinh lớp 12 THPT
        * J: Tổng số TN THPT
        * K: Tỷ lệ (%) = J / I * 100
      + L..N: Tỷ lệ đối tượng 18-21 có bằng TN THPT hệ GDTX, TCCN năm vừa qua:
        * L: Học sinh học năm cuối GDTX/TCCN
        * M: Tổng số TN GDTX/TCCN
        * N: Tỷ lệ (%) = M / L * 100
      + O..Q: Đi học các trình độ sau THPT:
        * O: Đại học
        * P: Cao đẳng
        * Q: Trung cấp
      + R..T: Bằng đào tạo nghề và tỷ lệ:
        * R: Tổng số học nghề từ 3 năm
        * S: Tổng số có bằng
        * T: Tỷ lệ (%) = S / R * 100
    - Dòng tổng cộng mẫu chuẩn: Hàng 47
    
    Quy tắc nghiệp vụ:
    1. Định vị động dòng Tổng cộng của 92 xã:
       - Quét nhãn "tổng cộng" / "tổng số" / "cộng" / "tổng" ở cột B hoặc A.
       - Thích ứng hoàn hảo với các biến thể hàng (dòng 9, 12, 17..47).
    2. Clean Recalculation cấp tỉnh:
       - Các cột số lượng cộng độc lập từ 92 xã.
       - Các cột tỷ lệ (%) tính lại chuẩn từ tổng tử / tổng mẫu cấp tỉnh:
         + H47 = sum(E47:G47) / D47 * 100
         + K47 = J47 / I47 * 100
         + N47 = M47 / L47 * 100
         + T47 = S47 / R47 * 100
       - Định dạng hiển thị tỷ lệ: 0.00"%"
    """
    province_counts = {c: 0.0 for c in COUNT_COLS_2M10}
    commune_summaries = []
    
    total_communes = len(valid_files)
    
    for f in valid_files:
        fpath = f['filepath']
        cname = f['commune_name']
        cstt = f.get('commune_stt')
        
        cells = fast_read_sheet(fpath, '2.M10')
        if not cells:
            continue
            
        row_total = None
        for r in range(7, 70):
            b_val = str(cells.get(f"B{r}", {}).get("val") or "").strip().lower()
            a_val = str(cells.get(f"A{r}", {}).get("val") or "").strip().lower()
            if "tổng cộng" in b_val or "tổng số" in b_val or b_val == "cộng" or a_val == "cộng" or "tổng" in b_val:
                row_total = r
                break
                
        if row_total is None:
            row_total = 47
            
        c_vals = {c: clean_num(cells.get(f"{c}{row_total}", {}).get("val")) for c in COUNT_COLS_2M10}
        
        # Nếu dòng row_total bị trống thì tính tổng từ các hàng chi tiết 8..row_total - 1
        if sum(c_vals.values()) == 0:
            for r in range(8, row_total):
                for c in COUNT_COLS_2M10:
                    c_vals[c] += clean_num(cells.get(f"{c}{r}", {}).get("val"))
                    
        # Tính tỷ lệ cấp xã
        c_rates = {}
        sum_efg_c = c_vals['E'] + c_vals['F'] + c_vals['G']
        c_rates['H'] = round(sum_efg_c / c_vals['D'] * 100, 2) if c_vals['D'] > 0 else 0.0
        c_rates['K'] = round(c_vals['J'] / c_vals['I'] * 100, 2) if c_vals['I'] > 0 else 0.0
        c_rates['N'] = round(c_vals['M'] / c_vals['L'] * 100, 2) if c_vals['L'] > 0 else 0.0
        c_rates['T'] = round(c_vals['S'] / c_vals['R'] * 100, 2) if c_vals['R'] > 0 else 0.0
        
        for c in COUNT_COLS_2M10:
            province_counts[c] += c_vals[c]
            
        commune_summaries.append({
            'stt': cstt,
            'commune_name': cname,
            'filename': f.get('filename', ''),
            'row_total': row_total,
            'counts': c_vals,
            'rates': c_rates
        })

    # Tính tỷ lệ cấp tỉnh
    province_rates = {}
    sum_efg_prov = province_counts['E'] + province_counts['F'] + province_counts['G']
    province_rates['H'] = round(sum_efg_prov / province_counts['D'] * 100, 2) if province_counts['D'] > 0 else 0.0
    province_rates['K'] = round(province_counts['J'] / province_counts['I'] * 100, 2) if province_counts['I'] > 0 else 0.0
    province_rates['N'] = round(province_counts['M'] / province_counts['L'] * 100, 2) if province_counts['L'] > 0 else 0.0
    province_rates['T'] = round(province_counts['S'] / province_counts['R'] * 100, 2) if province_counts['R'] > 0 else 0.0

    # aggregated_cells điền vào template cấp tỉnh tại Hàng 47
    aggregated_cells = {}
    for c in COUNT_COLS_2M10:
        aggregated_cells[f"{c}47"] = province_counts[c]
    for c in RATE_COLS_2M10:
        aggregated_cells[f"{c}47"] = province_rates[c]
        
    return {
        'total_communes_count': total_communes,
        'commune_names': [f['commune_name'] for f in valid_files],
        'commune_summaries': commune_summaries,
        'province_counts': province_counts,
        'province_rates': province_rates,
        'aggregated_cells': aggregated_cells
    }

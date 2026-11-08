import openpyxl
from core.fast_excel import fast_read_sheet

COUNT_COLS_2M9 = ['C', 'D', 'E', 'G', 'H', 'I', 'J', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S']
RATE_COLS_2M9 = ['F', 'K', 'T', 'U', 'V']
ALL_COLS_2M9 = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V']

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

def process_2_m9(valid_files):
    """
    Tổng hợp Biểu 2.M9: Biểu tổng hợp kết quả huy động phổ cập giáo dục trung học.
    
    Cấu trúc bảng theo mẫu chuẩn:
    - Bảng gồm các chỉ tiêu:
      + C: Tổng số đối tượng phổ cập (15 - 21 tuổi)
      + D..H: Tỷ lệ đối tượng 15-18 tốt nghiệp THCS (D: Phải PC, E: TN THCS, F: Tỷ lệ %, G: Còn học, H: Bỏ học)
      + I..K: Tỷ lệ tốt nghiệp THCS năm qua (I: Lớp 9, J: TN THCS, K: Tỷ lệ %)
      + L..N: Phân luồng sau THCS năm qua (L: Vào 10 THPT, M: GDTX/Nghề, N: Bỏ học THPT)
      + O..V: Tỷ lệ thanh niên 18-21 có bằng TN THPT, BTTHPT, TCCN, Nghề (O: Phải PC, P: THPT, Q: GDTX, R: TCCN, S: Nghề, T: TL TCCN, U: TL Nghề, V: Tỷ lệ %)
    - Dòng tổng cộng mẫu chuẩn: Hàng 48
    
    Quy tắc nghiệp vụ:
    1. Định vị động dòng Tổng cộng của 92 xã:
       - Quét nhãn "tổng cộng" / "tổng số" / "cộng" / "tổng" ở cột B hoặc A.
       - Thích ứng hoàn hảo với các xã có số thôn/xóm khác nhau (dòng 10, 13, 18..48).
    2. Clean Recalculation cấp tỉnh:
       - Các cột số lượng cộng độc lập từ 92 xã.
       - Các cột tỷ lệ (%) tính lại từ tổng tử / tổng mẫu cấp tỉnh:
         + F48 = E48 / D48 * 100
         + K48 = J48 / I48 * 100
         + T48 = R48 / O48 * 100
         + U48 = S48 / O48 * 100
         + V48 = (P48 + Q48 + R48 + S48) / O48 * 100
       - Định dạng hiển thị tỷ lệ: 0.00"%"
    """
    province_counts = {c: 0.0 for c in COUNT_COLS_2M9}
    commune_summaries = []
    
    total_communes = len(valid_files)
    
    for f in valid_files:
        fpath = f['filepath']
        cname = f['commune_name']
        cstt = f.get('commune_stt')
        
        cells = fast_read_sheet(fpath, '2.M9')
        if not cells:
            continue
            
        row_total = None
        for r in range(8, 70):
            b_val = str(cells.get(f"B{r}", {}).get("val") or "").strip().lower()
            a_val = str(cells.get(f"A{r}", {}).get("val") or "").strip().lower()
            if "tổng cộng" in b_val or "tổng số" in b_val or b_val == "cộng" or a_val == "cộng" or "tổng" in b_val:
                row_total = r
                break
                
        if row_total is None:
            row_total = 48
            
        c_vals = {c: clean_num(cells.get(f"{c}{row_total}", {}).get("val")) for c in COUNT_COLS_2M9}
        
        # Nếu dòng row_total bị trống thì tính tổng từ các hàng chi tiết 9..row_total - 1
        if sum(c_vals.values()) == 0:
            for r in range(9, row_total):
                for c in COUNT_COLS_2M9:
                    c_vals[c] += clean_num(cells.get(f"{c}{r}", {}).get("val"))
                    
        # Tính tỷ lệ cấp xã
        c_rates = {}
        c_rates['F'] = round(c_vals['E'] / c_vals['D'] * 100, 2) if c_vals['D'] > 0 else 0.0
        c_rates['K'] = round(c_vals['J'] / c_vals['I'] * 100, 2) if c_vals['I'] > 0 else 0.0
        c_rates['T'] = round(c_vals['R'] / c_vals['O'] * 100, 2) if c_vals['O'] > 0 else 0.0
        c_rates['U'] = round(c_vals['S'] / c_vals['O'] * 100, 2) if c_vals['O'] > 0 else 0.0
        sum_pqrs_c = c_vals['P'] + c_vals['Q'] + c_vals['R'] + c_vals['S']
        c_rates['V'] = round(sum_pqrs_c / c_vals['O'] * 100, 2) if c_vals['O'] > 0 else 0.0
        
        for c in COUNT_COLS_2M9:
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
    province_rates['F'] = round(province_counts['E'] / province_counts['D'] * 100, 2) if province_counts['D'] > 0 else 0.0
    province_rates['K'] = round(province_counts['J'] / province_counts['I'] * 100, 2) if province_counts['I'] > 0 else 0.0
    province_rates['T'] = round(province_counts['R'] / province_counts['O'] * 100, 2) if province_counts['O'] > 0 else 0.0
    province_rates['U'] = round(province_counts['S'] / province_counts['O'] * 100, 2) if province_counts['O'] > 0 else 0.0
    sum_pqrs_prov = province_counts['P'] + province_counts['Q'] + province_counts['R'] + province_counts['S']
    province_rates['V'] = round(sum_pqrs_prov / province_counts['O'] * 100, 2) if province_counts['O'] > 0 else 0.0

    # aggregated_cells điền vào template cấp tỉnh tại Hàng 48
    aggregated_cells = {}
    for c in COUNT_COLS_2M9:
        aggregated_cells[f"{c}48"] = province_counts[c]
    for c in RATE_COLS_2M9:
        aggregated_cells[f"{c}48"] = province_rates[c]
        
    return {
        'total_communes_count': total_communes,
        'commune_names': [f['commune_name'] for f in valid_files],
        'commune_summaries': commune_summaries,
        'province_counts': province_counts,
        'province_rates': province_rates,
        'aggregated_cells': aggregated_cells
    }

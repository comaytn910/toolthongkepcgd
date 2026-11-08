import openpyxl
from core.fast_excel import fast_read_sheet

COLS_2M6 = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
ROWS_2M6 = list(range(6, 14))  # Hàng 6 đến 13 (8 độ tuổi: 11-18t)

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

def process_2_m6(valid_files):
    """
    Tổng hợp Biểu 2.M6: Thống kê đối tượng khuyết tật cấp THCS.
    
    Cấu trúc bảng theo mẫu chuẩn:
    - Bảng gồm 8 hàng độ tuổi (Hàng 6..13) tương ứng năm sinh 2015..2008 (11..18 tuổi) và Hàng 14 là HÀNG CỘNG.
    - Cột A: Năm sinh (2015..2008)
    - Cột B: Độ tuổi (11..18)
    - Cột C: Tổng số đối tượng khuyết tật
    - Cột D..K: Dạng tật (Vận động, Nghe nói, Nhìn, Thần kinh tâm thần, Trí tuệ, Tự kỷ, Học tập, Khác)
    - Cột L: Số HS khuyết tật tiếp cận giáo dục
    - Cột M: Tỷ lệ (%) tiếp cận giáo dục = L / C * 100
    
    Quy tắc nghiệp vụ cốt lõi:
    1. Bảng cấu trúc cố định tuyệt đối trên cả 92 xã: Hàng 14 là Hàng Cộng ở 100% các xã.
    2. Cấp tỉnh tính lại tỷ lệ chuẩn:
       - Cột M ở các hàng 6..13: M_r = L_r / C_r * 100 (với C_r > 0)
       - Cột M ở hàng 14 (Tổng cộng): M_14 = L_14 / C_14 * 100
       - Khắc phục lỗi công thức mẫu average(M6:M13) bằng Clean Recalculation từ tổng tử / tổng mẫu.
       - Định dạng tỷ lệ 0.00"%" theo đúng chuẩn quy ước.
    3. Đối soát liên biểu:
       - C14 (3,601 trẻ) khớp với E46/M46 của 2.M3.
       - L14 (2,869 trẻ) khớp với O46 của 2.M3.
    """
    commune_details = {}
    total_communes = len(valid_files)
    
    # Ma trận tổng cộng cấp tỉnh cho từng ô trong bảng (Cột C..L x Hàng 6..14)
    prov_matrix = {f"{c}{r}": 0.0 for r in ROWS_2M6 + [14] for c in COLS_2M6}
    
    for f in valid_files:
        fpath = f['filepath']
        cname = f['commune_name']
        cstt = f.get('commune_stt')
        
        cells = fast_read_sheet(fpath, '2.M6')
        if not cells:
            continue
            
        comm_matrix = {}
        for r in ROWS_2M6 + [14]:
            for c in COLS_2M6:
                val = clean_num(cells.get(f"{c}{r}", {}).get("val"))
                comm_matrix[f"{c}{r}"] = val
                prov_matrix[f"{c}{r}"] += val
                
        commune_details[cname] = {
            'stt': cstt,
            'filename': f.get('filename', ''),
            'matrix': comm_matrix
        }

    # Tính toán lại tỷ lệ (Cột M) cấp tỉnh cho từng hàng 6..14
    prov_rates = {}
    for r in ROWS_2M6:
        c_val = prov_matrix[f"C{r}"]
        l_val = prov_matrix[f"L{r}"]
        r_val = (l_val / c_val * 100) if c_val > 0 else 0.0
        prov_rates[f"M{r}"] = round(r_val, 2)
        
    c_tot = prov_matrix["C14"]
    l_tot = prov_matrix["L14"]
    r_tot = (l_tot / c_tot * 100) if c_tot > 0 else 0.0
    prov_rates["M14"] = round(r_tot, 2)
    
    # Tạo aggregated_cells
    aggregated_cells = {}
    for coord, val in prov_matrix.items():
        aggregated_cells[coord] = val
    for coord, val in prov_rates.items():
        aggregated_cells[coord] = val
        
    return {
        'total_communes_count': total_communes,
        'commune_names': [f['commune_name'] for f in valid_files],
        'commune_details': commune_details,
        'prov_matrix': prov_matrix,
        'prov_rates': prov_rates,
        'aggregated_cells': aggregated_cells,
        'total_kt': c_tot,
        'total_access': l_tot,
        'rate_access': r_tot
    }

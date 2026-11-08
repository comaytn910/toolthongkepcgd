import openpyxl
from core.fast_excel import fast_read_sheet

COLS_2M7 = [openpyxl.utils.get_column_letter(c) for c in range(3, 40)]  # C đến AM (3..39)
DATA_COLS = COLS_2M7[:-1]  # C đến AL

AGE_GROUP_COLS_2M7 = {
    '10 tuổi': ['C', 'D', 'E', 'F'],
    '11 tuổi': ['G', 'H', 'I', 'J'],
    '12 tuổi': ['K', 'L', 'M', 'N'],
    '13 tuổi': ['O', 'P', 'Q', 'R'],
    '14 tuổi': ['S', 'T', 'U', 'V'],
    '15 tuổi': ['W', 'X', 'Y', 'Z'],
    '16 tuổi': ['AA', 'AB', 'AC', 'AD'],
    '17 tuổi': ['AE', 'AF', 'AG', 'AH'],
    '18 tuổi': ['AI', 'AJ', 'AK', 'AL']
}

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

def process_2_m7(valid_files):
    """
    Tổng hợp Biểu 2.M7: Số liệu học sinh của địa phương đang học tại xã/phường ngoài 
    (trong tỉnh) và tỉnh/thành phố ngoài (cấp THCS).
    
    Cấu trúc bảng theo mẫu chuẩn:
    - Bảng gồm 2 phần dạng Expandable:
      + Phần I: Học sinh của địa phương đang học tại xã/phường ngoài (trong tỉnh) (dòng 7..20, dòng 21: Tổng số I)
      + Phần II: Học sinh của địa phương đang học tại tỉnh/thành phố ngoài (ngoại tỉnh) (dòng 23..38, dòng 39: Tổng số II)
      + Dòng 40: Tổng (I+II)
    - Cột C..AL: 9 độ tuổi (10 đến 18 tuổi), mỗi độ tuổi 4 cột lớp (L6, L7, L8, L9)
    - Cột AM: Cột Tổng hàng ngang
    
    Quy tắc nghiệp vụ cốt lõi:
    1. Thuật toán định vị động 2 phần:
       - Tìm p2_header (dòng tiêu đề Phần II) và row_grand (dòng Tổng I+II).
       - Quét lùi từ p2_header để tìm chính xác dòng Tổng Phần I (row_p1).
       - Quét lùi từ row_grand để tìm chính xác dòng Tổng Phần II (row_p2).
       - Thích ứng hoàn hảo với các xã chèn thêm hàng danh sách trường.
    2. Clean Recalculation cấp tỉnh:
       - Hàng 21 (Tổng I cấp tỉnh) = Tổng hợp từ dòng Tổng I của 92 xã.
       - Hàng 39 (Tổng II cấp tỉnh) = Tổng hợp từ dòng Tổng II của 92 xã.
       - Hàng 40 (Tổng I+II cấp tỉnh) = Hàng 21 + Hàng 39 cho từng cột C..AL và Cột AM.
    3. Thu thập danh sách chi tiết các trường ngoài địa bàn được khai báo.
    """
    commune_summaries = []
    school_records = []
    
    province_p1 = {c: 0.0 for c in COLS_2M7}
    province_p2 = {c: 0.0 for c in COLS_2M7}
    province_tot = {c: 0.0 for c in COLS_2M7}
    
    total_communes = len(valid_files)
    
    for f in valid_files:
        fpath = f['filepath']
        cname = f['commune_name']
        cstt = f.get('commune_stt')
        
        cells = fast_read_sheet(fpath, '2.M7')
        if not cells:
            continue
            
        p2_header = None
        row_grand = None
        
        for r in range(6, 150):
            b_val = str(cells.get(f"B{r}", {}).get("val") or "").strip().lower()
            a_val = str(cells.get(f"A{r}", {}).get("val") or "").strip().lower()
            
            if ("học sinh" in b_val or "địa phương" in b_val) and ("ngoại tỉnh" in b_val or "tỉnh/thành phố ngoài" in b_val or "tỉnh ngoài" in b_val):
                p2_header = r
            elif "tổng (i+ii)" in b_val or "tổng i+ii" in b_val or "cộng (i+ii)" in b_val or "tổng (i + ii)" in b_val or "cộng (i + ii)" in b_val:
                row_grand = r
            elif (b_val == "cộng" or a_val == "cộng" or b_val == "tổng cộng") and r > 25 and row_grand is None and p2_header is not None and r > p2_header + 2:
                row_grand = r
                
        # Tìm dòng Tổng Phần I: lùi từ p2_header - 1
        row_p1 = None
        max_search_p1 = p2_header if p2_header else 22
        for r in range(max_search_p1 - 1, 6, -1):
            b_val = str(cells.get(f"B{r}", {}).get("val") or "").strip().lower()
            a_val = str(cells.get(f"A{r}", {}).get("val") or "").strip().lower()
            if "tổng số i" in b_val or "tổng i" in b_val or b_val == "cộng" or a_val == "cộng" or "tổng" in b_val:
                row_p1 = r
                break
                
        # Tìm dòng Tổng Phần II: lùi từ row_grand - 1
        row_p2 = None
        max_search_p2 = row_grand if row_grand else 150
        min_search_p2 = p2_header if p2_header else (row_p1 + 1 if row_p1 else 22)
        for r in range(max_search_p2 - 1, min_search_p2, -1):
            b_val = str(cells.get(f"B{r}", {}).get("val") or "").strip().lower()
            a_val = str(cells.get(f"A{r}", {}).get("val") or "").strip().lower()
            if "tổng số ii" in b_val or "tổng ii" in b_val or b_val == "cộng" or a_val == "cộng" or "tổng" in b_val:
                row_p2 = r
                break
                
        # Fallback
        if row_p1 is None: row_p1 = 21
        if p2_header is None: p2_header = row_p1 + 1
        if row_p2 is None: row_p2 = 39
        if row_grand is None: row_grand = 40
        
        # Đọc dữ liệu dòng Tổng Phần I
        vals_p1 = {c: clean_num(cells.get(f"{c}{row_p1}", {}).get("val")) for c in COLS_2M7}
        # Đọc dữ liệu dòng Tổng Phần II
        vals_p2 = {c: clean_num(cells.get(f"{c}{row_p2}", {}).get("val")) for c in COLS_2M7}
        
        # Chuẩn hóa Cột AM
        calc_am1 = sum(vals_p1[c] for c in DATA_COLS)
        if vals_p1['AM'] == 0 and calc_am1 > 0: vals_p1['AM'] = calc_am1
        
        calc_am2 = sum(vals_p2[c] for c in DATA_COLS)
        if vals_p2['AM'] == 0 and calc_am2 > 0: vals_p2['AM'] = calc_am2
        
        vals_grand = {c: vals_p1[c] + vals_p2[c] for c in COLS_2M7}
        
        # Cộng dồn toàn tỉnh
        for c in COLS_2M7:
            province_p1[c] += vals_p1[c]
            province_p2[c] += vals_p2[c]
            province_tot[c] += vals_grand[c]
            
        # Thu thập chi tiết các trường
        for r in range(7, row_p1):
            s_name = str(cells.get(f"B{r}", {}).get("val") or "").strip()
            am_val = clean_num(cells.get(f"AM{r}", {}).get("val"))
            sum_row = sum(clean_num(cells.get(f"{c}{r}", {}).get("val")) for c in DATA_COLS)
            real_tot = am_val if am_val > 0 else sum_row
            if s_name and not s_name.lower().startswith("cộng") and not s_name.lower().startswith("tổng") and real_tot > 0:
                school_records.append({
                    'commune_stt': cstt,
                    'commune_name': cname,
                    'school_name': s_name,
                    'part': 'Phần I (Trong tỉnh)',
                    'total_students': real_tot
                })
                
        for r in range(p2_header + 1, row_p2):
            s_name = str(cells.get(f"B{r}", {}).get("val") or "").strip()
            am_val = clean_num(cells.get(f"AM{r}", {}).get("val"))
            sum_row = sum(clean_num(cells.get(f"{c}{r}", {}).get("val")) for c in DATA_COLS)
            real_tot = am_val if am_val > 0 else sum_row
            if s_name and not s_name.lower().startswith("cộng") and not s_name.lower().startswith("tổng") and real_tot > 0:
                school_records.append({
                    'commune_stt': cstt,
                    'commune_name': cname,
                    'school_name': s_name,
                    'part': 'Phần II (Ngoài tỉnh)',
                    'total_students': real_tot
                })
                
        commune_summaries.append({
            'stt': cstt,
            'commune_name': cname,
            'filename': f.get('filename', ''),
            'row_p1': row_p1,
            'row_p2': row_p2,
            'row_grand': row_grand,
            'part1_sum': vals_p1['AM'],
            'part2_sum': vals_p2['AM'],
            'total_sum': vals_grand['AM'],
            'vals_p1': vals_p1,
            'vals_p2': vals_p2,
            'vals_grand': vals_grand
        })

    # Tạo aggregated_cells ghi vào template mẫu chuẩn (Hàng 21, 39, 40)
    aggregated_cells = {}
    for c in COLS_2M7:
        aggregated_cells[f"{c}21"] = province_p1[c]
        aggregated_cells[f"{c}39"] = province_p2[c]
        aggregated_cells[f"{c}40"] = province_tot[c]
        
    # Thống kê cơ cấu theo độ tuổi cấp THCS
    age_breakdown = {}
    for age_lbl, c_list in AGE_GROUP_COLS_2M7.items():
        cnt_p1 = sum(province_p1[c] for c in c_list)
        cnt_p2 = sum(province_p2[c] for c in c_list)
        cnt_tot = sum(province_tot[c] for c in c_list)
        age_breakdown[age_lbl] = {
            'part1': cnt_p1,
            'part2': cnt_p2,
            'total': cnt_tot
        }
        
    return {
        'total_communes_count': total_communes,
        'commune_names': [f['commune_name'] for f in valid_files],
        'commune_summaries': commune_summaries,
        'school_records': school_records,
        'total_schools_count': len(school_records),
        'province_p1': province_p1,
        'province_p2': province_p2,
        'province_tot': province_tot,
        'total_part1': province_p1['AM'],
        'total_part2': province_p2['AM'],
        'total_grand': province_tot['AM'],
        'age_breakdown': age_breakdown,
        'aggregated_cells': aggregated_cells
    }

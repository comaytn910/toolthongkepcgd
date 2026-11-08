import openpyxl
from core.fast_excel import fast_read_sheet

COLS_2M8 = [openpyxl.utils.get_column_letter(c) for c in range(3, 40)]  # C đến AM (3..39)
DATA_COLS = COLS_2M8[:-1]  # C đến AL

AGE_GROUP_COLS_2M8 = {
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

def process_2_m8(valid_files):
    """
    Tổng hợp Biểu 2.M8: Số liệu học sinh xã/phường ngoài (trong tỉnh) và tỉnh/thành phố ngoài 
    đang học tại địa phương (cấp THCS).
    
    Cấu trúc bảng theo mẫu chuẩn:
    - Bảng gồm 2 phần dạng Expandable:
      + Phần I: Học sinh xã/phường ngoài (trong tỉnh) đang học tại địa phương (dòng 7..23, dòng 24: Tổng I)
      + Phần II: Học sinh tỉnh/thành phố ngoài đang học tại địa phương (dòng 25..37, dòng 38: Tổng II)
      + Dòng 39: Tổng số (I+II)
    - Cột C..AL: 9 độ tuổi (10 đến 18 tuổi), mỗi độ tuổi 4 cột lớp (L6, L7, L8, L9)
    - Cột AM: Cột Tổng hàng ngang
    
    Quy tắc nghiệp vụ cốt lõi:
    1. Định vị động 2 phần thích ứng mở rộng dòng trường:
       - Tìm p2_header (dòng tiêu đề Phần II) và row_grand (dòng Tổng số I+II).
       - Quét lùi tìm chính xác dòng Tổng Phần I (row_p1) và dòng Tổng Phần II (row_p2).
    2. Clean Recalculation cấp tỉnh:
       - Tính lại Cột AM = sum(C..AL) cho từng xã và cấp tỉnh, khắc phục lỗi 16 xã gõ đè số 39 
         hoặc gõ thiếu ở cột AM nguồn (tiêu biểu Phường Phan Đình Phùng nguồn AM=1284 nhưng tổng C..AL=1702).
       - Hàng 24 (Tổng I cấp tỉnh) = Tổng hợp Phần I từ 92 xã.
       - Hàng 38 (Tổng II cấp tỉnh) = Tổng hợp Phần II từ 92 xã.
       - Hàng 39 (Tổng số I+II cấp tỉnh) = Hàng 24 + Hàng 38.
    3. Thu thập danh sách chi tiết các trường ngoài địa bàn được khai báo.
    """
    commune_summaries = []
    school_records = []
    discrepancies = []
    
    province_p1 = {c: 0.0 for c in COLS_2M8}
    province_p2 = {c: 0.0 for c in COLS_2M8}
    province_tot = {c: 0.0 for c in COLS_2M8}
    
    total_communes = len(valid_files)
    
    for f in valid_files:
        fpath = f['filepath']
        cname = f['commune_name']
        cstt = f.get('commune_stt')
        
        cells = fast_read_sheet(fpath, '2.M8')
        if not cells:
            continue
            
        p2_header = None
        row_grand = None
        
        for r in range(6, 150):
            b_val = str(cells.get(f"B{r}", {}).get("val") or "").strip().lower()
            a_val = str(cells.get(f"A{r}", {}).get("val") or "").strip().lower()
            
            if ("học sinh" in b_val or "tỉnh/thành phố ngoài" in b_val or "tỉnh ngoài" in b_val or "ngoại tỉnh" in b_val) and ("đang học tại địa phương" in b_val or a_val == "ii"):
                p2_header = r
            elif "tổng số" in b_val or "tổng cộng" in b_val or "tổng (i+ii)" in b_val or "tổng i+ii" in b_val or "cộng (i+ii)" in b_val:
                row_grand = r
            elif (b_val == "tổng" or a_val == "tổng" or b_val == "cộng" or a_val == "cộng") and r > 30 and row_grand is None and p2_header is not None and r > p2_header + 2:
                row_grand = r
                
        # Tìm dòng Tổng Phần I: lùi từ p2_header - 1
        row_p1 = None
        max_search_p1 = p2_header if p2_header else 25
        for r in range(max_search_p1 - 1, 6, -1):
            b_val = str(cells.get(f"B{r}", {}).get("val") or "").strip().lower()
            a_val = str(cells.get(f"A{r}", {}).get("val") or "").strip().lower()
            if "tổng" in b_val or "cộng" in b_val or a_val == "tổng" or a_val == "cộng":
                row_p1 = r
                break
                
        # Tìm dòng Tổng Phần II: lùi từ row_grand - 1
        row_p2 = None
        max_search_p2 = row_grand if row_grand else 150
        min_search_p2 = p2_header if p2_header else (row_p1 + 1 if row_p1 else 25)
        for r in range(max_search_p2 - 1, min_search_p2, -1):
            b_val = str(cells.get(f"B{r}", {}).get("val") or "").strip().lower()
            a_val = str(cells.get(f"A{r}", {}).get("val") or "").strip().lower()
            if "tổng" in b_val or "cộng" in b_val or a_val == "tổng" or a_val == "cộng":
                row_p2 = r
                break
                
        # Fallback
        if row_p1 is None: row_p1 = 24
        if p2_header is None: p2_header = 25
        if row_p2 is None: row_p2 = 38
        if row_grand is None: row_grand = 39
        
        # Đọc dữ liệu dòng Tổng Phần I
        vals_p1 = {c: clean_num(cells.get(f"{c}{row_p1}", {}).get("val")) for c in DATA_COLS}
        calc_am1 = sum(vals_p1.values())
        src_am1 = clean_num(cells.get(f"AM{row_p1}", {}).get("val"))
        vals_p1['AM'] = calc_am1 if calc_am1 > 0 else src_am1
        
        # Đọc dữ liệu dòng Tổng Phần II
        vals_p2 = {c: clean_num(cells.get(f"{c}{row_p2}", {}).get("val")) for c in DATA_COLS}
        calc_am2 = sum(vals_p2.values())
        src_am2 = clean_num(cells.get(f"AM{row_p2}", {}).get("val"))
        vals_p2['AM'] = calc_am2 if calc_am2 > 0 else src_am2
        
        # Ghi nhận sai lệch nguồn ở Phần I nếu có
        if abs(calc_am1 - src_am1) > 0.01:
            discrepancies.append({
                'commune_stt': cstt,
                'commune_name': cname,
                'part': 'Phần I',
                'calc_am': calc_am1,
                'src_am': src_am1,
                'diff': calc_am1 - src_am1,
                'reason': 'Cột AM nguồn chứa số tĩnh không khớp tổng các cột tuổi C..AL'
            })
            
        vals_grand = {c: vals_p1[c] + vals_p2[c] for c in COLS_2M8}
        
        # Cộng dồn toàn tỉnh
        for c in COLS_2M8:
            province_p1[c] += vals_p1[c]
            province_p2[c] += vals_p2[c]
            province_tot[c] += vals_grand[c]
            
        # Thu thập chi tiết các trường Phần I
        for r in range(8, row_p1):
            s_name = str(cells.get(f"B{r}", {}).get("val") or "").strip()
            am_val = clean_num(cells.get(f"AM{r}", {}).get("val"))
            sum_row = sum(clean_num(cells.get(f"{c}{r}", {}).get("val")) for c in DATA_COLS)
            real_tot = sum_row if sum_row > 0 else am_val
            if s_name and not s_name.lower().startswith("cộng") and not s_name.lower().startswith("tổng") and real_tot > 0:
                school_records.append({
                    'commune_stt': cstt,
                    'commune_name': cname,
                    'school_name': s_name,
                    'part': 'Phần I (Trong tỉnh)',
                    'total_students': real_tot
                })
                
        # Thu thập chi tiết các trường Phần II
        for r in range(p2_header + 1, row_p2):
            s_name = str(cells.get(f"B{r}", {}).get("val") or "").strip()
            am_val = clean_num(cells.get(f"AM{r}", {}).get("val"))
            sum_row = sum(clean_num(cells.get(f"{c}{r}", {}).get("val")) for c in DATA_COLS)
            real_tot = sum_row if sum_row > 0 else am_val
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

    # aggregated_cells ghi vào template mẫu chuẩn (Hàng 24, 38, 39)
    aggregated_cells = {}
    for c in COLS_2M8:
        aggregated_cells[f"{c}24"] = province_p1[c]
        aggregated_cells[f"{c}38"] = province_p2[c]
        aggregated_cells[f"{c}39"] = province_tot[c]
        
    # Thống kê cơ cấu theo độ tuổi cấp THCS
    age_breakdown = {}
    for age_lbl, c_list in AGE_GROUP_COLS_2M8.items():
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
        'discrepancies': discrepancies,
        'province_p1': province_p1,
        'province_p2': province_p2,
        'province_tot': province_tot,
        'total_part1': province_p1['AM'],
        'total_part2': province_p2['AM'],
        'total_grand': province_tot['AM'],
        'age_breakdown': age_breakdown,
        'aggregated_cells': aggregated_cells
    }

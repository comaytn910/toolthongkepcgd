import openpyxl
from core.fast_excel import fast_read_sheet

# Danh sách cột từ C đến AQ (41 cột tương ứng với các độ tuổi và lớp học)
ALL_COLS_M6 = [openpyxl.utils.get_column_letter(c) for c in range(3, 44)] # C đến AQ

AGE_COL_GROUPS = {
    '6 Tuổi': ['C', 'D', 'E'],
    '7 Tuổi': ['F', 'G', 'H', 'I'],
    '8 Tuổi': ['J', 'K', 'L', 'M', 'N'],
    '9 Tuổi': ['O', 'P', 'Q', 'R', 'S', 'T'],
    '10 Tuổi': ['U', 'V', 'W', 'X', 'Y', 'Z', 'AA'],
    '11 Tuổi': ['AB', 'AC', 'AD', 'AE'],
    '12 Tuổi': ['AF', 'AG', 'AH', 'AI'],
    '13 Tuổi': ['AJ', 'AK', 'AL', 'AM'],
    '14 Tuổi': ['AN', 'AO', 'AP', 'AQ']
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

def process_1_m6(valid_files):
    """
    Tổng hợp Biểu 1.M6: Số liệu học sinh của xã, phường đang học tại xã/phường ngoài 
    (trong tỉnh) và tỉnh/thành phố ngoài (Cấp Tiểu học).
    
    Cấu trúc bảng:
    - Phần I (Hàng 7..27): Học sinh của xã học ngoài xã trong tỉnh (Hàng 27: Cộng Phần I).
    - Phần II (Hàng 28..48): Học sinh của xã học ngoài tỉnh (Hàng 48: Cộng Phần II).
    - Hàng 49: Cộng (I + II) - Tổng toàn bộ học sinh đi học ngoài địa bàn.
    - Cột C đến AQ: 41 cột phân theo độ tuổi (6..14) và lớp học (CRL, L1..L5).
    - Cột AR: Tổng cộng các cột C..AQ.
    
    Quy tắc nghiệp vụ:
    1. Nhận diện động dòng "Cộng" Phần I và Phần II của từng xã theo text tiếng Việt 
       để tránh sai lệch khi xã chèn thêm hàng danh sách trường.
    2. Cấp tỉnh tính lại:
       + Hàng 27 (Tổng Phần I cấp tỉnh) = SUM các dòng Cộng I của 92 xã.
       + Hàng 48 (Tổng Phần II cấp tỉnh) = SUM các dòng Cộng II của 92 xã.
       + Hàng 49 (Tổng I+II cấp tỉnh) = Hàng 27 + Hàng 48 cho từng cột C..AQ và AR.
    3. Thu thập danh sách chi tiết các trường ngoài địa bàn được khai báo.
    """
    commune_summaries = []
    school_records = []
    
    province_p1 = {c: 0.0 for c in ALL_COLS_M6}
    province_p2 = {c: 0.0 for c in ALL_COLS_M6}
    province_tot = {c: 0.0 for c in ALL_COLS_M6}
    
    total_communes = len(valid_files)
    
    for f in valid_files:
        fpath = f['filepath']
        cname = f['commune_name']
        cstt = f.get('commune_stt')
        
        cells = fast_read_sheet(fpath, '1.M6')
        if not cells:
            continue
            
        # Tìm vị trí các dòng Cộng Phần I, Cộng Phần II và Cộng (I + II)
        p2_header = None
        row_tot = None
        for r in range(5, 130):
            b_val = str(cells.get(f"B{r}", {}).get('val') or '').strip().lower()
            a_val = str(cells.get(f"A{r}", {}).get('val') or '').strip().lower()
            if 'học sinh của xã đang học tại' in b_val and 'tỉnh' in b_val and ('ngoài' in b_val or 'thành phố' in b_val):
                p2_header = r
            elif 'cộng (i' in a_val or 'cộng (i' in b_val or 'cộng ( i' in a_val or 'cộng ( i' in b_val:
                row_tot = r
                
        # Tìm dòng Cộng Phần I: dòng có chữ "cộng" trước p2_header
        row_p1 = None
        max_search_p1 = p2_header if p2_header else 28
        for r in range(max_search_p1 - 1, 6, -1):
            b_val = str(cells.get(f"B{r}", {}).get('val') or '').strip().lower()
            a_val = str(cells.get(f"A{r}", {}).get('val') or '').strip().lower()
            if b_val == 'cộng' or a_val == 'cộng':
                row_p1 = r
                break

        # Tìm dòng Cộng Phần II: dòng có chữ "cộng" trước row_tot (hoặc sau p2_header)
        row_p2 = None
        max_search_p2 = row_tot if row_tot else 130
        min_search_p2 = p2_header if p2_header else (row_p1 + 1 if row_p1 else 28)
        for r in range(max_search_p2 - 1, min_search_p2, -1):
            b_val = str(cells.get(f"B{r}", {}).get('val') or '').strip().lower()
            a_val = str(cells.get(f"A{r}", {}).get('val') or '').strip().lower()
            if b_val == 'cộng' or a_val == 'cộng':
                row_p2 = r
                break

        # Fallback vị trí mặc định nếu xã không có text
        if row_p1 is None: row_p1 = 27
        if row_p2 is None: row_p2 = 48
        if row_tot is None: row_tot = 49
        
        # Đọc dữ liệu dòng Cộng I của xã
        p1_row_vals = {c: clean_num(cells.get(f"{c}{row_p1}", {}).get('val')) for c in ALL_COLS_M6}
        sum_p1_commune = sum(p1_row_vals.values())
        
        # Đọc dữ liệu dòng Cộng II của xã
        p2_row_vals = {c: clean_num(cells.get(f"{c}{row_p2}", {}).get('val')) for c in ALL_COLS_M6}
        sum_p2_commune = sum(p2_row_vals.values())
        
        # Dòng tổng (I + II) của xã tính lại chuẩn = p1 + p2
        tot_row_vals = {c: p1_row_vals[c] + p2_row_vals[c] for c in ALL_COLS_M6}
        sum_tot_commune = sum_p1_commune + sum_p2_commune
        
        # Cộng dồn cấp tỉnh
        for c in ALL_COLS_M6:
            province_p1[c] += p1_row_vals[c]
            province_p2[c] += p2_row_vals[c]
            province_tot[c] += tot_row_vals[c]
            
        # Thu thập trường học chi tiết ở Phần I (trước row_p1)
        for r in range(8, row_p1):
            school_name = str(cells.get(f"B{r}", {}).get('val') or '').strip()
            if school_name and school_name.lower() not in ['cộng', 'tổng']:
                row_vals = {c: clean_num(cells.get(f"{c}{r}", {}).get('val')) for c in ALL_COLS_M6}
                r_sum = sum(row_vals.values())
                if r_sum > 0:
                    school_records.append({
                        'commune': cname,
                        'part': 'Ngoài xã (trong tỉnh)',
                        'school_name': school_name,
                        'total_students': r_sum,
                        'by_age': {age: sum(row_vals[c] for c in cols) for age, cols in AGE_COL_GROUPS.items()}
                    })
                    
        # Thu thập trường học chi tiết ở Phần II (giữa row_p1 và row_p2)
        start_p2 = row_p1 + 2 # thường sau dòng tiêu đề II
        for r in range(start_p2, row_p2):
            school_name = str(cells.get(f"B{r}", {}).get('val') or '').strip()
            if school_name and school_name.lower() not in ['cộng', 'tổng']:
                row_vals = {c: clean_num(cells.get(f"{c}{r}", {}).get('val')) for c in ALL_COLS_M6}
                r_sum = sum(row_vals.values())
                if r_sum > 0:
                    school_records.append({
                        'commune': cname,
                        'part': 'Ngoài tỉnh',
                        'school_name': school_name,
                        'total_students': r_sum,
                        'by_age': {age: sum(row_vals[c] for c in cols) for age, cols in AGE_COL_GROUPS.items()}
                    })

        # Lưu bản ghi tóm tắt từng xã
        commune_summaries.append({
            'stt': cstt,
            'commune_name': cname,
            'filename': f['filename'],
            'part1_sum': sum_p1_commune,
            'part2_sum': sum_p2_commune,
            'total_sum': sum_tot_commune,
            'p1_by_age': {age: sum(p1_row_vals[c] for c in cols) for age, cols in AGE_COL_GROUPS.items()},
            'p2_by_age': {age: sum(p2_row_vals[c] for c in cols) for age, cols in AGE_COL_GROUPS.items()},
            'tot_by_age': {age: sum(tot_row_vals[c] for c in cols) for age, cols in AGE_COL_GROUPS.items()}
        })

    # Cột AR cấp tỉnh = SUM(C..AQ)
    total_ar27 = sum(province_p1.values())
    total_ar48 = sum(province_p2.values())
    total_ar49 = sum(province_tot.values())
    
    aggregated_cells = {}
    for c in ALL_COLS_M6:
        aggregated_cells[f"{c}27"] = province_p1[c]
        aggregated_cells[f"{c}48"] = province_p2[c]
        aggregated_cells[f"{c}49"] = province_tot[c]
        
    aggregated_cells['AR27'] = total_ar27
    aggregated_cells['AR48'] = total_ar48
    aggregated_cells['AR49'] = total_ar49
    
    # Thống kê tổng hợp theo nhóm độ tuổi toàn tỉnh
    province_by_age = {
        age: {
            'part1': sum(province_p1[c] for c in cols),
            'part2': sum(province_p2[c] for c in cols),
            'total': sum(province_tot[c] for c in cols),
            'pct': (sum(province_tot[c] for c in cols) / total_ar49 * 100.0) if total_ar49 > 0 else 0.0
        }
        for age, cols in AGE_COL_GROUPS.items()
    }
    
    return {
        'sheet_code': '1.M6',
        'sheet_name': 'Học sinh học ngoài xã, ngoài tỉnh (Tiểu học)',
        'total_communes_count': total_communes,
        'commune_names': [f['commune_name'] for f in valid_files],
        'aggregated_cells': aggregated_cells,
        'province_p1': province_p1,
        'province_p2': province_p2,
        'province_tot': province_tot,
        'province_by_age': province_by_age,
        'total_part1': total_ar27,
        'total_part2': total_ar48,
        'total_grand': total_ar49,
        'commune_summaries': commune_summaries,
        'school_records': school_records
    }

import openpyxl
from core.fast_excel import fast_read_sheet

# Danh sách cột từ C đến AQ (41 cột tương ứng với các độ tuổi và lớp học)
ALL_COLS_M7 = [openpyxl.utils.get_column_letter(c) for c in range(3, 44)] # C đến AQ

AGE_COL_GROUPS_M7 = {
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

def process_1_m7(valid_files):
    """
    Tổng hợp Biểu 1.M7: Số liệu học sinh ngoài địa bàn đang học tại các trường trên địa bàn xã 
    (Cấp Tiểu học).
    
    Cấu trúc bảng theo mẫu chuẩn:
    - Hàng 7: Cộng (I + II) - Tổng toàn bộ học sinh ngoại địa bàn học tại xã.
    - Phần I (Hàng 8..34): Học sinh xã/phường ngoài (trong tỉnh) học tại xã (Hàng 34: Cộng I).
    - Phần II (Hàng 35..68): Học sinh tỉnh/thành phố ngoài học tại xã (Hàng 68: Cộng II).
    - Cột C đến AQ: 41 cột phân theo độ tuổi (6..14) và lớp học (CRL, L1..L5, HTCTTH).
    - Cột AR: Tổng cộng các cột C..AQ.
    
    Quy tắc nghiệp vụ:
    1. Nhận diện động các dòng:
       - row_tot: Dòng có text 'Cộng (I + II)' (ở mẫu chuẩn là hàng 7).
       - row_p1: Dòng 'Cộng I' (kết thúc Phần I).
       - row_p2: Dòng 'Cộng II' (kết thúc Phần II).
    2. Cấp tỉnh tính lại (Clean Recalculation):
       - Hàng 34 (Tổng Phần I cấp tỉnh) = SUM các dòng Cộng I của 92 xã.
       - Hàng 68 (Tổng Phần II cấp tỉnh) = SUM các dòng Cộng II của 92 xã.
       - Hàng 7 (Tổng I+II cấp tỉnh) = Hàng 34 + Hàng 68 cho từng cột C..AQ.
       - Cột AR cấp tỉnh tính lại = SUM(C..AQ).
       - Loại trừ tự động sai lệch gõ đè ở hàng tổng nguồn (ví dụ: xã Nghinh Tường ô H7 bị gõ đè số 3 thay vì 2).
    3. Thu thập danh sách chi tiết các trường ngoài địa bàn học sinh theo học.
    """
    commune_summaries = []
    school_records = []
    discrepancies = []
    
    province_p1 = {c: 0.0 for c in ALL_COLS_M7}
    province_p2 = {c: 0.0 for c in ALL_COLS_M7}
    province_tot = {c: 0.0 for c in ALL_COLS_M7}
    
    total_communes = len(valid_files)
    
    for f in valid_files:
        fpath = f['filepath']
        cname = f['commune_name']
        cstt = f.get('commune_stt')
        
        cells = fast_read_sheet(fpath, '1.M7')
        if not cells:
            continue
            
        row_tot = None
        p1_header = None
        row_p1 = None
        p2_header = None
        row_p2 = None
        
        for r in range(5, 150):
            a_val = str(cells.get(f"A{r}", {}).get('val') or '').strip().lower()
            b_val = str(cells.get(f"B{r}", {}).get('val') or '').strip().lower()
            
            if 'cộng (i' in a_val or 'cộng ( i' in a_val or 'cộng (i' in b_val:
                row_tot = r
            elif 'ngoài (trong tỉnh)' in b_val or ('ngoài' in b_val and 'trong tỉnh' in b_val):
                p1_header = r
            elif b_val == 'cộng i' or a_val == 'cộng i' or (b_val == 'cộng' and p2_header is None and p1_header is not None and r > p1_header):
                row_p1 = r
            elif 'ngoài đang học tại xã' in b_val and ('tỉnh' in b_val or 'thành phố' in b_val):
                p2_header = r
            elif b_val == 'cộng ii' or a_val == 'cộng ii' or (b_val == 'cộng' and p2_header is not None and r > p2_header):
                row_p2 = r
                
        # Giá trị Phần I của xã
        p1_vals = {c: clean_num(cells.get(f"{c}{row_p1}", {}).get('val')) for c in ALL_COLS_M7} if row_p1 else {c: 0.0 for c in ALL_COLS_M7}
        p1_sum = sum(p1_vals.values())
        
        # Giá trị Phần II của xã
        p2_vals = {c: clean_num(cells.get(f"{c}{row_p2}", {}).get('val')) for c in ALL_COLS_M7} if row_p2 else {c: 0.0 for c in ALL_COLS_M7}
        p2_sum = sum(p2_vals.values())
        
        # Tổng chuẩn = p1 + p2
        comm_tot_vals = {c: p1_vals[c] + p2_vals[c] for c in ALL_COLS_M7}
        comm_tot_sum = p1_sum + p2_sum
        
        # Kiểm tra đối soát với Hàng 7 nguồn của xã
        src_tot_ar = clean_num(cells.get(f"AR{row_tot}", {}).get('val')) if row_tot else 0.0
        src_tot_caq_sum = sum(clean_num(cells.get(f"{c}{row_tot}", {}).get('val')) for c in ALL_COLS_M7) if row_tot else 0.0
        
        if abs(src_tot_caq_sum - comm_tot_sum) > 0.01 or abs(src_tot_ar - comm_tot_sum) > 0.01:
            diff_cols = []
            if row_tot:
                for c in ALL_COLS_M7:
                    c_src = clean_num(cells.get(f"{c}{row_tot}", {}).get('val'))
                    c_calc = comm_tot_vals[c]
                    if abs(c_src - c_calc) > 0.01:
                        diff_cols.append(f"Cột {c}: nguồn={c_src}, tính chuẩn={c_calc}")
            discrepancies.append({
                'commune': cname,
                'row_tot': row_tot,
                'src_tot_ar': src_tot_ar,
                'src_caq_sum': src_tot_caq_sum,
                'calc_sum': comm_tot_sum,
                'diff': src_tot_caq_sum - comm_tot_sum,
                'diff_cols': "; ".join(diff_cols)
            })
            
        # Cộng dồn cấp tỉnh
        for c in ALL_COLS_M7:
            province_p1[c] += p1_vals[c]
            province_p2[c] += p2_vals[c]
            province_tot[c] += comm_tot_vals[c]
            
        # Thống kê theo độ tuổi của xã
        comm_age_sum = {}
        for age_name, cols in AGE_COL_GROUPS_M7.items():
            comm_age_sum[age_name] = sum(comm_tot_vals[c] for c in cols)
            
        # Thu thập chi tiết trường học Phần I
        if p1_header and row_p1:
            for r in range(p1_header + 1, row_p1):
                s_name = str(cells.get(f"B{r}", {}).get('val') or '').strip()
                s_vals = {c: clean_num(cells.get(f"{c}{r}", {}).get('val')) for c in ALL_COLS_M7}
                s_sum = sum(s_vals.values())
                if s_name or s_sum > 0:
                    by_age = {age_k: sum(s_vals[c] for c in cols) for age_k, cols in AGE_COL_GROUPS_M7.items()}
                    school_records.append({
                        'commune': cname,
                        'part': 'Phần I (Trong tỉnh)',
                        'school_name': s_name if s_name else f"[Dòng {r} không tên]",
                        'total_students': s_sum,
                        'by_age': by_age
                    })
                    
        # Thu thập chi tiết trường học Phần II
        if p2_header and row_p2:
            for r in range(p2_header + 1, row_p2):
                s_name = str(cells.get(f"B{r}", {}).get('val') or '').strip()
                s_vals = {c: clean_num(cells.get(f"{c}{r}", {}).get('val')) for c in ALL_COLS_M7}
                s_sum = sum(s_vals.values())
                if s_name or s_sum > 0:
                    by_age = {age_k: sum(s_vals[c] for c in cols) for age_k, cols in AGE_COL_GROUPS_M7.items()}
                    school_records.append({
                        'commune': cname,
                        'part': 'Phần II (Ngoài tỉnh)',
                        'school_name': s_name if s_name else f"[Dòng {r} không tên]",
                        'total_students': s_sum,
                        'by_age': by_age
                    })
                    
        commune_summaries.append({
            'stt': cstt,
            'commune_name': cname,
            'filename': f.get('filename', ''),
            'rows': {'tot': row_tot, 'p1': row_p1, 'p2': row_p2},
            'part1_sum': p1_sum,
            'part2_sum': p2_sum,
            'total_sum': comm_tot_sum,
            'tot_by_age': comm_age_sum
        })

    # Tổng kết độ tuổi cấp tỉnh
    province_by_age = {}
    for age_name, cols in AGE_COL_GROUPS_M7.items():
        p1_a = sum(province_p1[c] for c in cols)
        p2_a = sum(province_p2[c] for c in cols)
        tot_a = sum(province_tot[c] for c in cols)
        province_by_age[age_name] = {
            'part1': p1_a,
            'part2': p2_a,
            'total': tot_a
        }
        
    tot_p1 = sum(province_p1.values())
    tot_p2 = sum(province_p2.values())
    tot_grand = sum(province_tot.values())
    
    # Chuẩn bị aggregated_cells để ghi đè vào template 1.M7
    # Template 1.M7:
    # Hàng 7: Cộng (I + II)
    # Hàng 34: Cộng I
    # Hàng 68: Cộng II
    aggregated_cells = {}
    for c in ALL_COLS_M7:
        aggregated_cells[f"{c}7"] = province_tot[c]
        aggregated_cells[f"{c}34"] = province_p1[c]
        aggregated_cells[f"{c}68"] = province_p2[c]
        
    aggregated_cells["AR7"] = tot_grand
    aggregated_cells["AR34"] = tot_p1
    aggregated_cells["AR68"] = tot_p2
    
    return {
        'total_communes_count': total_communes,
        'commune_names': [f['commune_name'] for f in valid_files],
        'commune_summaries': commune_summaries,
        'school_records': school_records,
        'discrepancies': discrepancies,
        'province_p1': province_p1,
        'province_p2': province_p2,
        'province_tot': province_tot,
        'province_by_age': province_by_age,
        'total_part1': tot_p1,
        'total_part2': tot_p2,
        'total_grand': tot_grand,
        'aggregated_cells': aggregated_cells
    }

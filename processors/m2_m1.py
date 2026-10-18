import openpyxl
from core.fast_excel import fast_read_sheet

# 20 cột kênh học tập chi tiết (11-18 tuổi)
# 11t: F (PT), G (TX)
# 12t: H (PT), I (TX)
# 13t: J (PT), K (TX)
# 14t: L (PT), M (TX)
# 15t: N (PT), O (TX), P (GDNN)
# 16t: Q (PT), R (TX), S (GDNN)
# 17t: T (PT), U (TX), V (GDNN)
# 18t: W (PT), X (TX), Y (GDNN)
DETAIL_COLS = ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y']

# 8 cột dân số và khuyết tật (Hàng 7-13)
POP_COLS = ['F', 'H', 'J', 'L', 'N', 'Q', 'T', 'W']

# Cột nhóm Tổng (15-18)
GROUP_COL_Z = 'Z'

# 12 cột tuổi 15-18 dùng để tính cột Z trong các hàng chi tiết 15-49
COLS_15_18 = ['N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y']

# 4 cột tuổi 15-18 dùng để tính cột Z trong hàng dân số 7-12 và hàng 13
POP_COLS_15_18 = ['N', 'Q', 'T', 'W']

# Hàng chi tiết 15 đến 49
DETAIL_ROWS = list(range(15, 50))

# Hàng dân số 7 đến 12
POP_ROWS = list(range(7, 13))

def parse_m2_cell_value(val):
    if val is None:
        return 'EMPTY', None
    if isinstance(val, (int, float)):
        return 'NUMBER', float(val)
    val_str = str(val).strip()
    if val_str == '':
        return 'EMPTY', None
    if val_str.lower() == 'x':
        return 'NOT_APPLICABLE', 'x'
    try:
        clean_num = val_str.replace(',', '')
        return 'NUMBER', float(clean_num)
    except ValueError:
        return 'ERROR_TYPE', val_str

def process_2_m1(valid_files):
    """
    Tổng hợp biểu 2.M1: Thống kê TTN PCGD THCS.
    Quy tắc:
    1. Phân biệt độ tuổi (11-18) và các kênh học tập: PT, TX, GDNN.
    2. Cột Z cấp tỉnh tính lại từ chi tiết các tuổi 15-18 (tránh cộng trùng).
    3. Hàng 13 (Số PPC) tính lại = Hàng 7 (Tổng số) - Hàng 10 (Khuyết tật).
    4. Đối chiếu Z nguồn với Z tính toán của từng xã để phát hiện sai lệch.
    5. Tính bảng tiêu chí (Hàng 52-59) từ số liệu chi tiết đã tổng hợp.
    6. Chuẩn hóa H59 theo thang 0-100 (khắc phục công thức mẫu quên *100).
    """
    commune_data = {}
    source_z_discrepancies = []
    
    aggregated_cells = {}
    valid_commune_counts = {}
    status_cells = {}
    
    total_communes_count = len(valid_files)
    
    # 1. Đọc dữ liệu từ từng xã
    for file_info in valid_files:
        c_name = file_info['commune_name']
        f_path = file_info['filepath']
        
        fast_cells = fast_read_sheet(f_path, '2.M1')
        def get_val(coord):
            return fast_cells.get(coord, {}).get('val')
        
        c_counts = {}
        c_source_z = {}
        c_calc_z = {}
        
        # Đọc hàng dân số 7-12
        for r in POP_ROWS:
            for col in POP_COLS:
                val = get_val(f'{col}{r}')
                t, v = parse_m2_cell_value(val)
                c_counts[(col, r)] = v if t == 'NUMBER' else 0.0
            
            # Đọc Z nguồn và tính Z
            z_src_val = get_val(f'{GROUP_COL_Z}{r}')
            t_z, v_z = parse_m2_cell_value(z_src_val)
            c_source_z[r] = v_z if t_z == 'NUMBER' else 0.0
            calc_z = sum(c_counts[(col, r)] for col in POP_COLS_15_18)
            c_calc_z[r] = calc_z
            
            if abs(c_source_z[r] - calc_z) > 0.01:
                source_z_discrepancies.append({
                    'commune': c_name,
                    'row': r,
                    'row_name': get_val(f'A{r}') or get_val(f'C{r}') or f'Hàng {r}',
                    'source_z': c_source_z[r],
                    'calc_z': calc_z,
                    'diff': calc_z - c_source_z[r]
                })
        
        # Tính hàng 13 (Số PPC = Hàng 7 - Hàng 10)
        for col in POP_COLS:
            c_counts[(col, 13)] = c_counts[(col, 7)] - c_counts[(col, 10)]
        
        z_src_13 = get_val(f'{GROUP_COL_Z}13')
        t_z13, v_z13 = parse_m2_cell_value(z_src_13)
        c_source_z[13] = v_z13 if t_z13 == 'NUMBER' else 0.0
        c_calc_z[13] = c_calc_z[7] - c_calc_z[10]
        if abs(c_source_z[13] - c_calc_z[13]) > 0.01:
            source_z_discrepancies.append({
                'commune': c_name,
                'row': 13,
                'row_name': 'Số phải phổ cập (PPC)',
                'source_z': c_source_z[13],
                'calc_z': c_calc_z[13],
                'diff': c_calc_z[13] - c_source_z[13]
            })
            
        # Đọc hàng chi tiết 15-49
        for r in DETAIL_ROWS:
            for col in DETAIL_COLS:
                val = get_val(f'{col}{r}')
                t, v = parse_m2_cell_value(val)
                c_counts[(col, r)] = v if t == 'NUMBER' else 0.0
                
            z_src_val = get_val(f'{GROUP_COL_Z}{r}')
            t_z, v_z = parse_m2_cell_value(z_src_val)
            c_source_z[r] = v_z if t_z == 'NUMBER' else 0.0
            calc_z = sum(c_counts[(col, r)] for col in COLS_15_18)
            c_calc_z[r] = calc_z
            
            if abs(c_source_z[r] - calc_z) > 0.01:
                source_z_discrepancies.append({
                    'commune': c_name,
                    'row': r,
                    'row_name': get_val(f'B{r}') or get_val(f'C{r}') or get_val(f'D{r}') or f'Hàng {r}',
                    'source_z': c_source_z[r],
                    'calc_z': calc_z,
                    'diff': calc_z - c_source_z[r]
                })
        
        # Đọc tiêu chí của xã (Hàng 52-59)
        c_criteria = {
            'G52': get_val('G52'),
            'G53': get_val('G53'),
            'G55_src': get_val('G55'),
            'H55_src': get_val('H55'),
            'G56_src': get_val('G56'),
            'H56_src': get_val('H56'),
            'G57_src': get_val('G57'),
            'H57_src': get_val('H57'),
            'G58_src': get_val('G58'),
            'H58_src': get_val('H58'),
            'G59_src': get_val('G59'),
            'H59_src': get_val('H59'),
        }
        
        # Tính lại tiêu chí của xã để đối chiếu
        c_calc_g55 = sum(c_counts[(col, r)] for r in range(35, 38) for col in COLS_15_18)
        c_calc_h55 = round(c_calc_g55 / c_calc_z[13] * 100, 2) if c_calc_z[13] > 0 else 0.0
        
        c_calc_g56 = sum(c_counts[(col, r)] for r in range(40, 43) for col in COLS_15_18)
        c_calc_h56 = round(c_calc_g56 / c_calc_z[13] * 100, 2) if c_calc_z[13] > 0 else 0.0
        
        c_calc_g57 = sum(c_counts[(col, 12)] for col in DETAIL_COLS if (col, 12) in c_counts)
        c_calc_denom57 = sum(c_counts[(col, 11)] for col in POP_COLS)
        c_calc_h57 = round(c_calc_g57 / c_calc_denom57 * 100, 2) if c_calc_denom57 > 0 else 0.0
        
        rows_g58 = list(range(15, 18)) + list(range(20, 23)) + list(range(25, 28)) + list(range(30, 33)) + list(range(40, 43))
        c_calc_g58 = sum(c_counts[(col, r)] for r in rows_g58 for col in DETAIL_COLS)
        c_calc_denom58 = sum(c_counts[(col, 13)] for col in POP_COLS)
        c_calc_h58 = round(c_calc_g58 / c_calc_denom58 * 100, 2) if c_calc_denom58 > 0 else 0.0
        
        cols_18 = ['W', 'X', 'Y']
        c_calc_g59 = sum(c_counts[(col, r)] for r in range(35, 38) for col in cols_18)
        c_w13 = c_counts[('W', 13)]
        c_calc_h59 = round(c_calc_g59 / c_w13 * 100, 2) if c_w13 > 0 else 0.0
        
        c_criteria.update({
            'G55_calc': c_calc_g55, 'H55_calc': c_calc_h55,
            'G56_calc': c_calc_g56, 'H56_calc': c_calc_h56,
            'G57_calc': c_calc_g57, 'H57_calc': c_calc_h57,
            'G58_calc': c_calc_g58, 'H58_calc': c_calc_h58,
            'G59_calc': c_calc_g59, 'H59_calc': c_calc_h59,
        })
        commune_data[c_name] = {
            'counts': c_counts,
            'source_z': c_source_z,
            'calc_z': c_calc_z,
            'criteria': c_criteria
        }
        
    # 2. Tổng hợp cấp tỉnh
    # Tổng hợp hàng dân số 7-12
    for r in POP_ROWS:
        for col in POP_COLS:
            key = f'{col}{r}'
            val_sum = sum(commune_data[c]['counts'][(col, r)] for c in commune_data)
            aggregated_cells[key] = val_sum
            valid_commune_counts[key] = total_communes_count
            status_cells[key] = f"Hợp lệ ({total_communes_count}/{total_communes_count} xã)"
            
        # Cột Z tính từ tổng cấp tỉnh của các cột 15-18
        key_z = f'{GROUP_COL_Z}{r}'
        val_z = sum(aggregated_cells[f'{col}{r}'] for col in POP_COLS_15_18)
        aggregated_cells[key_z] = val_z
        valid_commune_counts[key_z] = total_communes_count
        status_cells[key_z] = f"Tính từ N,Q,T,W cấp tỉnh ({total_communes_count}/{total_communes_count} xã)"

    # Tổng hợp hàng 13 (Số PPC = Hàng 7 - Hàng 10)
    for col in POP_COLS:
        key = f'{col}13'
        val = aggregated_cells[f'{col}7'] - aggregated_cells[f'{col}10']
        aggregated_cells[key] = val
        valid_commune_counts[key] = total_communes_count
        status_cells[key] = f"Tính lại = {col}7 - {col}10"
        
    key_z13 = f'{GROUP_COL_Z}13'
    val_z13 = aggregated_cells[f'{GROUP_COL_Z}7'] - aggregated_cells[f'{GROUP_COL_Z}10']
    aggregated_cells[key_z13] = val_z13
    valid_commune_counts[key_z13] = total_communes_count
    status_cells[key_z13] = f"Tính lại = Z7 - Z10"

    # Tổng hợp hàng chi tiết 15-49
    for r in DETAIL_ROWS:
        for col in DETAIL_COLS:
            key = f'{col}{r}'
            val_sum = sum(commune_data[c]['counts'][(col, r)] for c in commune_data)
            aggregated_cells[key] = val_sum
            valid_commune_counts[key] = total_communes_count
            status_cells[key] = f"Hợp lệ ({total_communes_count}/{total_communes_count} xã)"
            
        # Cột Z tính từ tổng cấp tỉnh của 12 cột tuổi 15-18
        key_z = f'{GROUP_COL_Z}{r}'
        val_z = sum(aggregated_cells[f'{col}{r}'] for col in COLS_15_18)
        aggregated_cells[key_z] = val_z
        valid_commune_counts[key_z] = total_communes_count
        status_cells[key_z] = f"Tính từ N:Y cấp tỉnh ({total_communes_count}/{total_communes_count} xã)"

    # 3. Tính bảng tiêu chí cấp tỉnh (Hàng 52-59)
    def parse_standard_level(val, default_level=3.0):
        if val is None:
            return None
        if isinstance(val, (int, float)):
            return float(val)
        val_str = str(val).strip()
        for ch in ['3', '2', '1']:
            if ch in val_str:
                return float(ch)
        return default_level

    # G52: Mức độ PCGDTH (Lấy mức tối thiểu hoặc phổ biến)
    m_52_vals = [parse_standard_level(commune_data[c]['criteria']['G52'], 3.0) for c in commune_data]
    m_52_vals = [v for v in m_52_vals if v is not None]
    prov_g52 = min(m_52_vals) if m_52_vals else 3.0
    aggregated_cells['G52'] = prov_g52
    status_cells['G52'] = "Mức đạt chuẩn PCGDTH"
    
    # G53: Mức độ XMC
    m_53_vals = [parse_standard_level(commune_data[c]['criteria']['G53'], 2.0) for c in commune_data]
    m_53_vals = [v for v in m_53_vals if v is not None]
    prov_g53 = min(m_53_vals) if m_53_vals else 2.0
    aggregated_cells['G53'] = prov_g53
    status_cells['G53'] = "Mức đạt chuẩn XMC"

    # Hàng 55: TTN 15-18 TNTHCS
    # SL: G55 = sum(N35:Y37) cấp tỉnh
    prov_g55 = sum(aggregated_cells[f'{col}{r}'] for r in range(35, 38) for col in COLS_15_18)
    aggregated_cells['G55'] = prov_g55
    status_cells['G55'] = "Tổng hợp từ N35:Y37"
    # TL: H55 = G55 / Z13 * 100
    prov_z13 = aggregated_cells[f'{GROUP_COL_Z}13']
    prov_h55 = round(prov_g55 / prov_z13 * 100, 2) if prov_z13 > 0 else 0.0
    aggregated_cells['H55'] = prov_h55
    status_cells['H55'] = f"G55 / Z13 * 100 ({prov_g55} / {prov_z13})"

    # Hàng 56: TTN 15-18 đã, đang học CT GDPT hoặc GDTX cấp THPT hoặc GDNN
    # SL: G56 = sum(N40:Y42) cấp tỉnh
    prov_g56 = sum(aggregated_cells[f'{col}{r}'] for r in range(40, 43) for col in COLS_15_18)
    aggregated_cells['G56'] = prov_g56
    status_cells['G56'] = "Tổng hợp từ N40:Y42"
    # TL: H56 = G56 / Z13 * 100
    prov_h56 = round(prov_g56 / prov_z13 * 100, 2) if prov_z13 > 0 else 0.0
    aggregated_cells['H56'] = prov_h56
    status_cells['H56'] = f"G56 / Z13 * 100 ({prov_g56} / {prov_z13})"

    # Hàng 57: TTN KT được tiếp cận GD (11-18)
    # SL: G57 = sum(F12:Y12) cấp tỉnh
    prov_g57 = sum(aggregated_cells[f'{col}12'] for col in DETAIL_COLS if f'{col}12' in aggregated_cells)
    aggregated_cells['G57'] = prov_g57
    status_cells['G57'] = "Tổng hợp từ F12:Y12"
    # Denom57 = sum(F11:W11)
    prov_denom57 = sum(aggregated_cells[f'{col}11'] for col in POP_COLS)
    prov_h57 = round(prov_g57 / prov_denom57 * 100, 2) if prov_denom57 > 0 else 0.0
    aggregated_cells['H57'] = prov_h57
    status_cells['H57'] = f"G57 / sum(F11:W11) * 100 ({prov_g57} / {prov_denom57})"

    # Hàng 58: Tổng số học sinh (11-18)
    # SL: G58 = sum(F15:Y17, F20:Y22, F25:Y27, F30:Y32, F40:Y42) cấp tỉnh
    prov_g58 = sum(aggregated_cells[f'{col}{r}'] for r in rows_g58 for col in DETAIL_COLS)
    aggregated_cells['G58'] = prov_g58
    status_cells['G58'] = "Tổng hợp từ các lớp 6,7,8,9,THPT đang học tại địa bàn và ngoài tỉnh"
    # Denom58 = sum(F13:W13)
    prov_denom58 = sum(aggregated_cells[f'{col}13'] for col in POP_COLS)
    prov_h58 = round(prov_g58 / prov_denom58 * 100, 2) if prov_denom58 > 0 else 0.0
    aggregated_cells['H58'] = prov_h58
    status_cells['H58'] = f"G58 / sum(F13:W13) * 100 ({prov_g58} / {prov_denom58})"

    # Hàng 59: Tổng số đối tượng 18 tuổi hoàn thành chương trình GD THCS
    # SL: G59 = sum(W35:Y37) cấp tỉnh
    prov_g59 = sum(aggregated_cells[f'{col}{r}'] for r in range(35, 38) for col in cols_18)
    aggregated_cells['G59'] = prov_g59
    status_cells['G59'] = "Tổng hợp từ W35:Y37"
    # Denom59 = W13
    prov_w13 = aggregated_cells['W13']
    prov_h59 = round(prov_g59 / prov_w13 * 100, 2) if prov_w13 > 0 else 0.0
    aggregated_cells['H59'] = prov_h59
    status_cells['H59'] = f"G59 / W13 * 100 ({prov_g59} / {prov_w13}) - Đã chuẩn hóa thang 0-100"

    return {
        'total_communes': total_communes_count,
        'commune_data': commune_data,
        'aggregated_cells': aggregated_cells,
        'valid_commune_counts': valid_commune_counts,
        'status_cells': status_cells,
        'source_z_discrepancies': source_z_discrepancies
    }

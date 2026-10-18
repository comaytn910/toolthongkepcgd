import openpyxl
from core.fast_excel import fast_read_sheet

# Các cột tuổi chi tiết
AGE_COLS_6_10 = ['F', 'G', 'H', 'I', 'J'] # 6, 7, 8, 9, 10 tuổi
GROUP_COL_6_10 = 'K'                       # Tổng (6-10)

AGE_COLS_11_14 = ['L', 'M', 'N', 'O']     # 11, 12, 13, 14 tuổi
GROUP_COL_11_14 = 'P'                      # Tổng (11-14)

ALL_AGE_COLS = AGE_COLS_6_10 + AGE_COLS_11_14
ALL_ACTIVE_COLS = AGE_COLS_6_10 + [GROUP_COL_6_10] + AGE_COLS_11_14 + [GROUP_COL_11_14]

# Các hàng số lượng trong Phần 1 (Hàng 8 đến 56, trừ Hàng 14 tính từ Hàng 8 - Hàng 11)
RAW_COUNT_ROWS = [8, 9, 10, 11, 12, 13] + list(range(15, 57))

def parse_m1_cell_value(val):
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

def process_1_m1(valid_files):
    """
    Tổng hợp biểu 1.M1: Thống kê trẻ em 6 đến 14 tuổi PCGD và GD bắt buộc TH.
    Quy tắc:
    1. Phân biệt nhóm 6-10 tuổi (F:J, Tổng K) và 11-14 tuổi (L:O, Tổng P).
    2. Không cộng trùng tổng nhóm với chi tiết.
    3. Cột K và P cấp tỉnh tính lại từ chi tiết các tuổi (K = SUM(F:J), P = SUM(L:O)).
    4. Đối chiếu K nguồn và P nguồn của từng xã để phát hiện sai lệch.
    5. Hàng 14 (Số PPC) tính lại = Hàng 8 (Tổng số) - Hàng 11 (Khuyết tật).
    6. Tính chuẩn xác Bảng tiêu chí Phần 2 (Hàng 58 đến 63) với tỷ lệ theo thang 0-100.
    """
    commune_data = {}
    source_kp_discrepancies = []
    
    aggregated_cells = {}
    valid_commune_counts = {}
    status_cells = {}
    
    total_communes_count = len(valid_files)
    
    # 1. Đọc dữ liệu từ từng xã
    for file_info in valid_files:
        c_name = file_info['commune_name']
        f_path = file_info['filepath']
        
        fast_cells = fast_read_sheet(f_path, '1.M1')
        c_matrix = {}
        for r in range(8, 64):
            for col in ALL_ACTIVE_COLS + ['F', 'G']:
                coord = f"{col}{r}"
                cell_item = fast_cells.get(coord, {'val': None, 'type': 'EMPTY', 'formula': None, 'raw': None})
                dtype, val = parse_m1_cell_value(cell_item['val'])
                c_matrix[coord] = {
                    'type': dtype,
                    'val': val,
                    'raw': cell_item['formula'] if cell_item['formula'] else cell_item['val'],
                    'is_formula': cell_item['formula'] is not None
                }
                
        commune_data[c_name] = c_matrix
        
    # 2. Tổng hợp các cột tuổi chi tiết (F..J và L..O) cho các hàng số lượng
    for r in RAW_COUNT_ROWS:
        for col in ALL_AGE_COLS:
            coord = f"{col}{r}"
            col_sum = 0.0
            valid_c = 0
            has_error = False
            has_empty = False
            
            for c_name, matrix in commune_data.items():
                c_cell = matrix[coord]
                if c_cell['type'] == 'NUMBER':
                    col_sum += c_cell['val']
                    valid_c += 1
                elif c_cell['type'] == 'EMPTY':
                    has_empty = True
                elif c_cell['type'] == 'ERROR_TYPE':
                    has_error = True
                    
            aggregated_cells[coord] = col_sum
            valid_commune_counts[coord] = valid_c
            
            if has_error or has_empty:
                status_cells[coord] = f"Chưa đủ dữ liệu (Đã cộng: {valid_c}/{total_communes_count} xã)"
            else:
                status_cells[coord] = f"Đầy đủ ({valid_c}/{total_communes_count} xã)"

    # 3. Tính lại Hàng 14 (Số phải phổ cập = Hàng 8 - Hàng 11) cho các cột tuổi chi tiết
    for col in ALL_AGE_COLS:
        coord_14 = f"{col}14"
        ppc_val = aggregated_cells[f"{col}8"] - aggregated_cells[f"{col}11"]
        aggregated_cells[coord_14] = ppc_val
        valid_commune_counts[coord_14] = min(valid_commune_counts[f"{col}8"], valid_commune_counts[f"{col}11"])
        status_cells[coord_14] = "Đầy đủ" if valid_commune_counts[coord_14] == total_communes_count else "Chưa đủ dữ liệu"

    # 4. Tính lại Cột K (Tổng 6-10t = SUM F:J) và Cột P (Tổng 11-14t = SUM L:O)
    # cho tất cả các hàng từ 8 đến 56
    for r in range(8, 57):
        # Cột K
        coord_k = f"K{r}"
        recalculated_k = sum(aggregated_cells[f"{col}{r}"] for col in AGE_COLS_6_10)
        aggregated_cells[coord_k] = recalculated_k
        valid_commune_counts[coord_k] = min(valid_commune_counts[f"{col}{r}"] for col in AGE_COLS_6_10)
        status_cells[coord_k] = "Đầy đủ" if valid_commune_counts[coord_k] == total_communes_count else "Chưa đủ dữ liệu"
        
        # Cột P
        coord_p = f"P{r}"
        recalculated_p = sum(aggregated_cells[f"{col}{r}"] for col in AGE_COLS_11_14)
        aggregated_cells[coord_p] = recalculated_p
        valid_commune_counts[coord_p] = min(valid_commune_counts[f"{col}{r}"] for col in AGE_COLS_11_14)
        status_cells[coord_p] = "Đầy đủ" if valid_commune_counts[coord_p] == total_communes_count else "Chưa đủ dữ liệu"
        
        # Đối chiếu K nguồn và P nguồn của từng xã
        for c_name, matrix in commune_data.items():
            # Kiểm tra K
            src_k = matrix[coord_k]['val']
            exp_k = sum(matrix[f"{col}{r}"]['val'] for col in AGE_COLS_6_10 if matrix[f"{col}{r}"]['type'] == 'NUMBER')
            if isinstance(src_k, (int, float)):
                diff_k = src_k - exp_k
                if abs(diff_k) > 0.01:
                    source_kp_discrepancies.append({
                        'commune': c_name,
                        'coord': coord_k,
                        'row': r,
                        'group': '6-10t',
                        'source_val': src_k,
                        'expected_val': exp_k,
                        'diff': diff_k,
                        'raw_formula': matrix[coord_k]['raw']
                    })
                    
            # Kiểm tra P
            src_p = matrix[coord_p]['val']
            exp_p = sum(matrix[f"{col}{r}"]['val'] for col in AGE_COLS_11_14 if matrix[f"{col}{r}"]['type'] == 'NUMBER')
            if isinstance(src_p, (int, float)):
                diff_p = src_p - exp_p
                if abs(diff_p) > 0.01:
                    source_kp_discrepancies.append({
                        'commune': c_name,
                        'coord': coord_p,
                        'row': r,
                        'group': '11-14t',
                        'source_val': src_p,
                        'expected_val': exp_p,
                        'diff': diff_p,
                        'raw_formula': matrix[coord_p]['raw']
                    })

    # 5. TÍNH BẢNG TIÊU CHÍ PHẦN 2 (Hàng 58 đến 63)
    criteria_results = {}
    
    # Hàng 58: Trẻ 6 tuổi vào lớp 1
    # Tử số F58 = SUM(F15:F17). Mẫu số = F14 (Số PPC 6 tuổi).
    f58_qty = aggregated_cells['F15'] + aggregated_cells['F16'] + aggregated_cells['F17']
    f58_den = aggregated_cells['F14']
    f58_rate = round(f58_qty / f58_den * 100, 2) if f58_den > 0 else 0.0
    aggregated_cells['F58'] = f58_qty
    aggregated_cells['G58'] = f58_rate
    criteria_results[58] = {
        'name': 'Trẻ 6 tuổi vào lớp 1',
        'qty': f58_qty,
        'rate': f58_rate,
        'rule': 'Tử số = F15+F16+F17, Mẫu số = F14 (PPC 6t)'
    }
    
    # Hàng 59: Trẻ 11 tuổi HTCTTH
    # Tử số F59 = SUM(L40:L42). Mẫu số = L14 (Số PPC 11 tuổi).
    f59_qty = aggregated_cells['L40'] + aggregated_cells['L41'] + aggregated_cells['L42']
    f59_den = aggregated_cells['L14']
    f59_rate = round(f59_qty / f59_den * 100, 2) if f59_den > 0 else 0.0
    aggregated_cells['F59'] = f59_qty
    aggregated_cells['G59'] = f59_rate
    criteria_results[59] = {
        'name': 'Trẻ 11 tuổi HTCTTH',
        'qty': f59_qty,
        'rate': f59_rate,
        'rule': 'Tử số = L40+L41+L42, Mẫu số = L14 (PPC 11t)'
    }
    
    # Hàng 60: Trẻ 11 tuổi đang học ở TH
    # Tử số F60 = SUM các lớp 1..5 của tuổi 11: L15..L17 + L20..L22 + L25..L27 + L30..L32 + L35..L37
    f60_qty = sum(aggregated_cells[f"L{r}"] for r in [15, 16, 17, 20, 21, 22, 25, 26, 27, 30, 31, 32, 35, 36, 37])
    f60_den = aggregated_cells['L14']
    f60_rate = round(f60_qty / f60_den * 100, 2) if f60_den > 0 else 0.0
    aggregated_cells['F60'] = f60_qty
    aggregated_cells['G60'] = f60_rate
    criteria_results[60] = {
        'name': 'Trẻ 11 tuổi đang học ở TH',
        'qty': f60_qty,
        'rate': f60_rate,
        'rule': 'Tử số = L15..L37 (học lớp 1-5), Mẫu số = L14 (PPC 11t)'
    }
    
    # Hàng 61: Trẻ đến 14 tuổi HTCTTH
    # Tử số F61 = P40 + P41 + P42 (HTCTTH nhóm 11-14 tuổi). Mẫu số = P14 (Tổng PPC 11-14t).
    f61_qty = aggregated_cells['P40'] + aggregated_cells['P41'] + aggregated_cells['P42']
    f61_den = aggregated_cells['P14']
    f61_rate = round(f61_qty / f61_den * 100, 2) if f61_den > 0 else 0.0
    aggregated_cells['F61'] = f61_qty
    aggregated_cells['G61'] = f61_rate
    criteria_results[61] = {
        'name': 'Trẻ đến 14 tuổi HTCTTH',
        'qty': f61_qty,
        'rate': f61_rate,
        'rule': 'Tử số = P40+P41+P42, Mẫu số = P14 (PPC 11-14t)'
    }
    
    # Hàng 62: Trẻ 14 tuổi HTCTTH (chỉ tính trẻ 14 tuổi)
    # Tử số F62 = O40 + O41 + O42. Mẫu số = O14 (Số PPC 14t).
    f62_qty = aggregated_cells['O40'] + aggregated_cells['O41'] + aggregated_cells['O42']
    f62_den = aggregated_cells['O14']
    f62_rate = round(f62_qty / f62_den * 100, 2) if f62_den > 0 else 0.0
    aggregated_cells['F62'] = f62_qty
    aggregated_cells['G62'] = f62_rate
    criteria_results[62] = {
        'name': 'Trẻ 14 tuổi HTCTTH (chỉ tính trẻ 14 tuổi)',
        'qty': f62_qty,
        'rate': f62_rate,
        'rule': 'Tử số = O40+O41+O42, Mẫu số = O14 (PPC 14t)'
    }
    
    # Hàng 63: Trẻ KT có khả năng HT được tiếp cận GD
    # Tử số F63 = K13 + P13 (tiếp cận GD). Mẫu số = K12 + P12 (có khả năng HT).
    f63_qty = aggregated_cells['K13'] + aggregated_cells['P13']
    f63_den = aggregated_cells['K12'] + aggregated_cells['P12']
    f63_rate = round(f63_qty / f63_den * 100, 2) if f63_den > 0 else 0.0
    aggregated_cells['F63'] = f63_qty
    aggregated_cells['G63'] = f63_rate
    criteria_results[63] = {
        'name': 'Trẻ KT có khả năng HT được tiếp cận GD',
        'qty': f63_qty,
        'rate': f63_rate,
        'rule': 'Tử số = K13+P13, Mẫu số = K12+P12 (KT có khả năng HT)'
    }
    
    return {
        'sheet_name': '1.M1',
        'aggregated_cells': aggregated_cells,
        'valid_commune_counts': valid_commune_counts,
        'status_cells': status_cells,
        'criteria_results': criteria_results,
        'source_kp_discrepancies': source_kp_discrepancies,
        'total_communes_count': total_communes_count,
        'commune_names': list(commune_data.keys()),
        'commune_data': commune_data
    }

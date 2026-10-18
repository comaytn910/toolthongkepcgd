import openpyxl
from core.fast_excel import fast_read_sheet

DISABILITY_COLS = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']
DATA_ROWS = [7, 8, 9, 10, 11, 12] # Tuổi 0 đến 5
TOTAL_ROW = 13                    # Hàng tổng cộng

def parse_m6_cell_value(val):
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

def process_0_m6(valid_files):
    """
    Tổng hợp biểu 0.M6: Báo cáo thống kê đối tượng khuyết tật MN.
    Quy tắc:
    1. Tổng hợp theo năm sinh, tuổi và 8 dạng khuyết tật (Cột D đến K).
    2. Cột C (Tổng số người KT) tính lại = SUM(D:K) của 8 dạng tật, đối chiếu với C nguồn.
    3. Cột L (Số HS KT tiếp cận GD) cộng dồn từ các xã.
    4. Cột M (Tỉ lệ): Tính lại = L / C * 100 theo đúng thang 0-100.
    5. Hàng 13 (Tổng cộng): Tính lại = SUM các hàng 7-12, Tỉ lệ M13 = L13 / C13 * 100.
    """
    commune_data = {}
    source_c_discrepancies = []
    
    aggregated_cells = {}
    valid_commune_counts = {}
    status_cells = {}
    
    total_communes_count = len(valid_files)
    
    # 1. Đọc dữ liệu từ từng xã
    for file_info in valid_files:
        c_name = file_info['commune_name']
        f_path = file_info['filepath']
        
        fast_cells = fast_read_sheet(f_path, '0.M6')
        c_matrix = {}
        for r in DATA_ROWS + [TOTAL_ROW]:
            for col in ['A', 'B', 'C', 'L', 'M'] + DISABILITY_COLS:
                coord = f"{col}{r}"
                cell_item = fast_cells.get(coord, {'val': None, 'type': 'EMPTY', 'formula': None, 'raw': None})
                dtype, val = parse_m6_cell_value(cell_item['val'])
                c_matrix[coord] = {
                    'type': dtype,
                    'val': val,
                    'raw': cell_item['formula'] if cell_item['formula'] else cell_item['val'],
                    'is_formula': cell_item['formula'] is not None
                }
                
        commune_data[c_name] = c_matrix
        
    # 2. Tổng hợp các cột số lượng cho từng hàng tuổi (7 đến 12):
    # Dạng tật D..K và Tiếp cận GD L
    for r in DATA_ROWS:
        cols_to_sum = DISABILITY_COLS + ['L']
        for col in cols_to_sum:
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
                
        # 3. Tính lại Cột C (Tổng số người khuyết tật) = SUM(D:K)
        coord_c = f"C{r}"
        recalculated_c = sum(aggregated_cells[f"{col}{r}"] for col in DISABILITY_COLS)
        aggregated_cells[coord_c] = recalculated_c
        min_valid_c = min(valid_commune_counts[f"{col}{r}"] for col in DISABILITY_COLS)
        valid_commune_counts[coord_c] = min_valid_c
        
        has_incomplete = any("Chưa đủ dữ liệu" in status_cells[f"{col}{r}"] for col in DISABILITY_COLS)
        status_cells[coord_c] = f"Chưa đủ dữ liệu ({min_valid_c}/{total_communes_count} xã)" if has_incomplete else f"Đầy đủ ({min_valid_c}/{total_communes_count} xã)"
        
        # Đối chiếu C nguồn của xã
        for c_name, matrix in commune_data.items():
            src_c = matrix[coord_c]['val']
            commune_exp_c = sum(matrix[f"{col}{r}"]['val'] for col in DISABILITY_COLS if matrix[f"{col}{r}"]['type'] == 'NUMBER')
            if isinstance(src_c, (int, float)):
                diff = src_c - commune_exp_c
                if abs(diff) > 0.01:
                    source_c_discrepancies.append({
                        'commune': c_name,
                        'coord': coord_c,
                        'row': r,
                        'source_c_val': src_c,
                        'expected_c_val': commune_exp_c,
                        'diff': diff,
                        'raw_formula': matrix[coord_c]['raw']
                    })
                    
        # 4. Tính lại Tỉ lệ M = L / C * 100
        coord_m = f"M{r}"
        c_val = aggregated_cells[coord_c]
        l_val = aggregated_cells[f"L{r}"]
        if c_val > 0:
            rate_m = round(l_val / c_val * 100, 2)
            aggregated_cells[coord_m] = rate_m
        else:
            aggregated_cells[coord_m] = 0.0
        valid_commune_counts[coord_m] = min(valid_commune_counts[coord_c], valid_commune_counts[f"L{r}"])
        status_cells[coord_m] = "Đầy đủ" if valid_commune_counts[coord_m] == total_communes_count else "Chưa đủ dữ liệu"
        
        # Cố định A và B
        birth_year = 2026 - (r - 7)
        age = r - 7
        aggregated_cells[f"A{r}"] = birth_year
        aggregated_cells[f"B{r}"] = age

    # 5. TỔNG HỢP HÀNG 13 (TỔNG CỘNG)
    aggregated_cells['A13'] = 'x'
    aggregated_cells['B13'] = 'Tổng:'
    
    # Tổng từng dạng tật D13..K13 và L13
    for col in DISABILITY_COLS + ['L']:
        coord_tot = f"{col}13"
        tot_val = sum(aggregated_cells[f"{col}{r}"] for r in DATA_ROWS)
        aggregated_cells[coord_tot] = tot_val
        valid_commune_counts[coord_tot] = min(valid_commune_counts[f"{col}{r}"] for r in DATA_ROWS)
        status_cells[coord_tot] = "Đầy đủ" if valid_commune_counts[coord_tot] == total_communes_count else "Chưa đủ dữ liệu"
        
    # Cột C13 (Tổng số người KT) = SUM(C7:C12) = SUM(D13:K13)
    recalculated_c13 = sum(aggregated_cells[f"C{r}"] for r in DATA_ROWS)
    aggregated_cells['C13'] = recalculated_c13
    valid_commune_counts['C13'] = min(valid_commune_counts[f"C{r}"] for r in DATA_ROWS)
    status_cells['C13'] = "Đầy đủ" if valid_commune_counts['C13'] == total_communes_count else "Chưa đủ dữ liệu"
    
    # Đối chiếu C13 nguồn của các xã
    for c_name, matrix in commune_data.items():
        src_c13 = matrix['C13']['val']
        commune_exp_c13 = sum(matrix[f"C{r}"]['val'] for r in DATA_ROWS if matrix[f"C{r}"]['type'] == 'NUMBER')
        if isinstance(src_c13, (int, float)):
            diff = src_c13 - commune_exp_c13
            if abs(diff) > 0.01:
                source_c_discrepancies.append({
                    'commune': c_name,
                    'coord': 'C13',
                    'row': 13,
                    'source_c_val': src_c13,
                    'expected_c_val': commune_exp_c13,
                    'diff': diff,
                    'raw_formula': matrix['C13']['raw']
                })
                
    # Cột M13: Tỉ lệ chung toàn tỉnh = L13 / C13 * 100
    if recalculated_c13 > 0:
        aggregated_cells['M13'] = round(aggregated_cells['L13'] / recalculated_c13 * 100, 2)
    else:
        aggregated_cells['M13'] = 0.0
    valid_commune_counts['M13'] = min(valid_commune_counts['C13'], valid_commune_counts['L13'])
    status_cells['M13'] = "Đầy đủ" if valid_commune_counts['M13'] == total_communes_count else "Chưa đủ dữ liệu"
    
    return {
        'sheet_name': '0.M6',
        'aggregated_cells': aggregated_cells,
        'valid_commune_counts': valid_commune_counts,
        'status_cells': status_cells,
        'source_c_discrepancies': source_c_discrepancies,
        'total_communes_count': total_communes_count,
        'commune_names': list(commune_data.keys()),
        'commune_data': commune_data
    }

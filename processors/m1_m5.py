import openpyxl
from core.fast_excel import fast_read_sheet

DISABILITY_COLS = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']
# Hàng 8 đến 16: tương ứng năm sinh 2020..2012, độ tuổi 6..14
DATA_ROWS = list(range(8, 17))
TOTAL_ROW = 17

def parse_m5_cell_value(val):
    if val is None:
        return 'EMPTY', None
    if isinstance(val, (int, float)):
        if val > 5000000:
            return 'EMPTY', None
        return 'NUMBER', float(val)
    val_str = str(val).strip()
    if val_str == '':
        return 'EMPTY', None
    if val_str.lower() == 'x':
        return 'NOT_APPLICABLE', 'x'
    try:
        clean_num = val_str.replace(',', '')
        v = float(clean_num)
        if v > 5000000:
            return 'EMPTY', None
        return 'NUMBER', v
    except ValueError:
        return 'ERROR_TYPE', val_str

def process_1_m5(valid_files):
    """
    Tổng hợp Biểu 1.M5: Thống kê đối tượng khuyết tật cấp Tiểu học.
    Độ tuổi: 6 đến 14 tuổi (Năm sinh 2020 đến 2012).
    Hàng 8 đến 16: Chi tiết từng năm sinh/độ tuổi.
    Hàng 17: Tổng cộng.
    
    Quy tắc nghiệp vụ:
    1. Tổng hợp theo năm sinh, độ tuổi và 8 dạng tật (D: Vận động, E: Nghe nói, F: Nhìn,
       G: Thần kinh tâm thần, H: Trí tuệ, I: Rối loạn tự kỷ, J: Học tập, K: Khác).
    2. Cột C (Tổng số người KT) tính lại = SUM(D:K) của 8 dạng tật, đối chiếu với C nguồn.
    3. Cột L (Số HS KT tiếp cận giáo dục) cộng dồn từ chi tiết từng xã.
    4. Cột M (Tỉ lệ %): Tính lại = L / C * 100 theo đúng thang 0-100.
    5. Hàng 17 (Tổng cộng): Tính lại từ SUM các hàng chi tiết 8-16 cấp tỉnh, 
       khắc phục triệt để lỗi mất dữ liệu khi xã gõ đè số 0 vào hàng 17 (ví dụ xã Yên Phong).
    """
    commune_data = {}
    source_c_discrepancies = []
    h17_discrepancies = []
    
    aggregated_cells = {}
    valid_commune_counts = {}
    status_cells = {}
    
    total_communes_count = len(valid_files)
    
    # 1. Đọc dữ liệu từ từng xã
    for file_info in valid_files:
        c_name = file_info['commune_name']
        f_path = file_info['filepath']
        
        fast_cells = fast_read_sheet(f_path, '1.M5')
        c_matrix = {}
        for r in DATA_ROWS + [TOTAL_ROW]:
            for col in ['A', 'B', 'C', 'L', 'M'] + DISABILITY_COLS:
                coord = f"{col}{r}"
                cell_item = fast_cells.get(coord, {'val': None, 'type': 'EMPTY', 'formula': None, 'raw': None})
                dtype, val = parse_m5_cell_value(cell_item['val'])
                c_matrix[coord] = {
                    'type': dtype,
                    'val': val,
                    'raw': cell_item['formula'] if cell_item['formula'] else cell_item['val'],
                    'is_formula': cell_item['formula'] is not None
                }
                
        commune_data[c_name] = c_matrix
        
        # Kiểm tra đối chiếu hàng 17 nguồn của xã so với tổng chi tiết hàng 8..16
        sum_c_detail = 0.0
        sum_l_detail = 0.0
        for r in DATA_ROWS:
            for d_col in DISABILITY_COLS:
                if c_matrix[f"{d_col}{r}"]['type'] == 'NUMBER':
                    sum_c_detail += c_matrix[f"{d_col}{r}"]['val']
            if c_matrix[f"L{r}"]['type'] == 'NUMBER':
                sum_l_detail += c_matrix[f"L{r}"]['val']
                
        src_c17 = c_matrix['C17']['val'] if c_matrix['C17']['type'] == 'NUMBER' else 0.0
        src_l17 = c_matrix['L17']['val'] if c_matrix['L17']['type'] == 'NUMBER' else 0.0
        
        if (abs(sum_c_detail - src_c17) > 0.01) or (abs(sum_l_detail - src_l17) > 0.01):
            h17_discrepancies.append({
                'commune': c_name,
                'detail_c': sum_c_detail,
                'detail_l': sum_l_detail,
                'src_c17': src_c17,
                'src_l17': src_l17,
                'diff_c': src_c17 - sum_c_detail,
                'diff_l': src_l17 - sum_l_detail,
                'reason': 'Hàng 17 nguồn của xã không khớp với tổng các dòng chi tiết hàng 8-16 (Đã tính chuẩn từ chi tiết)'
            })
            
    # 2. Tổng hợp các cột số lượng cho từng hàng tuổi (8 đến 16):
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
            
            if has_error:
                status_cells[coord] = f"Có lỗi sai kiểu ở một số xã (Đã cộng {valid_c}/{total_communes_count} xã)"
            elif has_empty and valid_c < total_communes_count:
                status_cells[coord] = f"Chưa đủ dữ liệu (Đã cộng {valid_c}/{total_communes_count} xã)"
            else:
                status_cells[coord] = "Đầy đủ"
                
        # 3. Tính lại Cột C (Tổng số trẻ KT hàng r) = SUM(D:K)
        expected_c = sum(aggregated_cells[f"{col}{r}"] for col in DISABILITY_COLS)
        aggregated_cells[f"C{r}"] = expected_c
        valid_commune_counts[f"C{r}"] = total_communes_count
        status_cells[f"C{r}"] = "Tính lại chuẩn từ 8 dạng tật D:K"
        
        # 4. Tính lại Cột M (Tỉ lệ tiếp cận GD hàng r) = L / C * 100
        l_val = aggregated_cells[f"L{r}"]
        if expected_c > 0:
            rate_m = round((l_val / expected_c) * 100.0, 2)
            aggregated_cells[f"M{r}"] = rate_m
            status_cells[f"M{r}"] = f"Tính lại chuẩn = L{r}/C{r}*100 ({l_val:,.0f}/{expected_c:,.0f})"
        else:
            aggregated_cells[f"M{r}"] = 0.0
            status_cells[f"M{r}"] = "Mẫu số C bằng 0"
        valid_commune_counts[f"M{r}"] = total_communes_count

        # Giữ nguyên A (Năm sinh) và B (Độ tuổi)
        birth_year = 2028 - r
        age = r - 2
        aggregated_cells[f"A{r}"] = birth_year
        aggregated_cells[f"B{r}"] = age
        status_cells[f"A{r}"] = "Chỉ tiêu cố định"
        status_cells[f"B{r}"] = "Chỉ tiêu cố định"
        valid_commune_counts[f"A{r}"] = total_communes_count
        valid_commune_counts[f"B{r}"] = total_communes_count

    # 5. Tính Hàng 17 (Tổng cộng toàn tỉnh)
    # Tổng các dạng tật D..K và tiếp cận L = SUM các hàng 8..16
    for col in DISABILITY_COLS + ['L']:
        sum_col_17 = sum(aggregated_cells[f"{col}{r}"] for r in DATA_ROWS)
        aggregated_cells[f"{col}{TOTAL_ROW}"] = sum_col_17
        valid_commune_counts[f"{col}{TOTAL_ROW}"] = total_communes_count
        status_cells[f"{col}{TOTAL_ROW}"] = f"Tổng cộng cấp tỉnh = SUM({col}8:{col}16)"

    # Cột C17 = SUM(C8..C16) = SUM(D17..K17)
    total_c17 = sum(aggregated_cells[f"C{r}"] for r in DATA_ROWS)
    aggregated_cells[f"C{TOTAL_ROW}"] = total_c17
    valid_commune_counts[f"C{TOTAL_ROW}"] = total_communes_count
    status_cells[f"C{TOTAL_ROW}"] = "Tổng cộng cấp tỉnh = SUM(C8:C16)"

    # Cột M17 (Tỉ lệ cấp tỉnh) = L17 / C17 * 100
    total_l17 = aggregated_cells[f"L{TOTAL_ROW}"]
    if total_c17 > 0:
        rate_m17 = round((total_l17 / total_c17) * 100.0, 2)
        aggregated_cells[f"M{TOTAL_ROW}"] = rate_m17
        status_cells[f"M{TOTAL_ROW}"] = f"Tính lại chuẩn = L17/C17*100 ({total_l17:,.0f}/{total_c17:,.0f})"
    else:
        aggregated_cells[f"M{TOTAL_ROW}"] = 0.0
        status_cells[f"M{TOTAL_ROW}"] = "Mẫu số C17 bằng 0"
    valid_commune_counts[f"M{TOTAL_ROW}"] = total_communes_count

    aggregated_cells[f"A{TOTAL_ROW}"] = ""
    aggregated_cells[f"B{TOTAL_ROW}"] = "Tổng"
    status_cells[f"A{TOTAL_ROW}"] = "Cố định"
    status_cells[f"B{TOTAL_ROW}"] = "Cố định"
    valid_commune_counts[f"A{TOTAL_ROW}"] = total_communes_count
    valid_commune_counts[f"B{TOTAL_ROW}"] = total_communes_count

    return {
        'sheet_code': '1.M5',
        'sheet_name': 'Đối tượng khuyết tật cấp Tiểu học',
        'total_communes_count': total_communes_count,
        'commune_names': [f['commune_name'] for f in valid_files],
        'aggregated_cells': aggregated_cells,
        'valid_commune_counts': valid_commune_counts,
        'status_cells': status_cells,
        'commune_data': commune_data,
        'h17_discrepancies': h17_discrepancies,
        'summary_metrics': {
            'total_disabled': total_c17,
            'total_access': total_l17,
            'rate_access': aggregated_cells[f"M{TOTAL_ROW}"],
            'by_disability': {col: aggregated_cells[f"{col}{TOTAL_ROW}"] for col in DISABILITY_COLS}
        }
    }

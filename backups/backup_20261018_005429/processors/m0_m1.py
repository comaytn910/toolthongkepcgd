import openpyxl
from openpyxl.utils import get_column_letter
import os

# 22 chỉ tiêu số lượng (hàng)
COUNT_ROWS_5_25 = [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 17, 18, 19, 20, 21, 23, 24, 25]
COUNT_ROWS_26_29 = [26, 27, 29]
ALL_COUNT_ROWS = COUNT_ROWS_5_25 + COUNT_ROWS_26_29 # Đúng 22 chỉ tiêu

# Cột tuổi
AGE_COLS_ALL = ['D', 'E', 'F', 'G', 'H', 'I', 'J'] # 0t đến 6t
AGE_COLS_0_5 = ['D', 'E', 'F', 'G', 'H', 'I']      # 0t đến 5t (dùng tính K hàng 5-25)
AGE_COLS_4_6 = ['H', 'I', 'J']                     # 4t đến 6t (dùng tính K hàng 26, 27, 29)

def parse_cell_value(val):
    """
    Phân biệt rõ ràng:
    - None hoặc '' -> ('EMPTY', None)
    - 'x' hoặc 'X' -> ('NOT_APPLICABLE', 'x')
    - float/int -> ('NUMBER', float(val))
    - str chứa số -> ('NUMBER', float) nếu ép kiểu được
    - chuỗi lỗi hoặc văn bản khác -> ('ERROR_TYPE', str(val))
    """
    if val is None:
        return 'EMPTY', None
    if isinstance(val, (int, float)):
        return 'NUMBER', float(val)
    val_str = str(val).strip()
    if val_str == '':
        return 'EMPTY', None
    if val_str.lower() == 'x':
        return 'NOT_APPLICABLE', 'x'
    # Thử ép kiểu số
    try:
        clean_num = val_str.replace(',', '')
        return 'NUMBER', float(clean_num)
    except ValueError:
        return 'ERROR_TYPE', val_str

def process_0_m1(valid_files):
    """
    Xử lý tổng hợp biểu 0.M1 theo đúng các nguyên tắc:
    1. Không che thiếu dữ liệu.
    2. Tính lại cột K từ các cột tuổi đã tổng hợp, đối chiếu K nguồn và ghi sai lệch.
    3. Tỷ lệ tính lại từ cùng tập xã có dữ liệu hợp lệ.
    4. Ghi nhận trạng thái riêng cho các chỉ tiêu chưa chốt quy tắc.
    """
    # Lưu vết dữ liệu từng xã
    commune_data = {}
    source_k_discrepancies = []
    
    # Ma trận tổng hợp: key = f"{col}{row}"
    aggregated_cells = {}
    valid_commune_counts = {}
    status_cells = {}
    
    # 1. Đọc dữ liệu thô từ từng xã
    for file_info in valid_files:
        c_name = file_info['commune_name']
        f_path = file_info['filepath']
        
        wb_data = openpyxl.load_workbook(f_path, data_only=True)
        wb_formula = openpyxl.load_workbook(f_path, data_only=False)
        ws_d = wb_data['0.M1']
        ws_f = wb_formula['0.M1']
        
        c_matrix = {}
        for r in range(5, 30):
            for col in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
                coord = f"{col}{r}"
                cell_d = ws_d[coord]
                cell_f = ws_f[coord]
                dtype, val = parse_cell_value(cell_d.value)
                raw_formula = cell_f.value if isinstance(cell_f.value, str) and cell_f.value.startswith('=') else None
                c_matrix[coord] = {
                    'type': dtype,
                    'val': val,
                    'raw': cell_f.value,
                    'is_formula': raw_formula is not None
                }
                
        # Đọc thêm hàng 31-40
        for r in range(31, 41):
            for col in ['A', 'B', 'C']:
                coord = f"{col}{r}"
                c_matrix[coord] = {
                    'val': ws_d[coord].value,
                    'raw': ws_f[coord].value
                }
                
        commune_data[c_name] = c_matrix
        wb_data.close()
        wb_formula.close()
        
    total_communes_count = len(valid_files)
    
    # 2. Tổng hợp 22 chỉ tiêu số lượng (Cột D đến J)
    for r in ALL_COUNT_ROWS:
        cols_to_sum = AGE_COLS_ALL if r in COUNT_ROWS_5_25 else AGE_COLS_4_6
        # Với hàng 26, 27, 29: D, E, F, G là 'x'
        if r in COUNT_ROWS_26_29:
            for col in ['D', 'E', 'F', 'G']:
                aggregated_cells[f"{col}{r}"] = 'x'
                valid_commune_counts[f"{col}{r}"] = total_communes_count
                status_cells[f"{col}{r}"] = "Không áp dụng"
                
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

    # 3. Tính lại CỘT K cho 22 chỉ tiêu số lượng & Đối chiếu K nguồn
    for r in ALL_COUNT_ROWS:
        coord_k = f"K{r}"
        if r in COUNT_ROWS_5_25:
            # K = SUM(D:I)
            recalculated_k = sum(aggregated_cells[f"{col}{r}"] for col in AGE_COLS_0_5 if isinstance(aggregated_cells[f"{col}{r}"], (int, float)))
            k_min_valid = min(valid_commune_counts[f"{col}{r}"] for col in AGE_COLS_0_5)
        else:
            # K = SUM(H:J)
            recalculated_k = sum(aggregated_cells[f"{col}{r}"] for col in AGE_COLS_4_6 if isinstance(aggregated_cells[f"{col}{r}"], (int, float)))
            k_min_valid = min(valid_commune_counts[f"{col}{r}"] for col in AGE_COLS_4_6)
            
        aggregated_cells[coord_k] = recalculated_k
        valid_commune_counts[coord_k] = k_min_valid
        
        # Đối chiếu với K nguồn của từng xã
        for c_name, matrix in commune_data.items():
            source_k_val = matrix[coord_k]['val']
            # Tính K riêng của xã đó xem có khớp với ô K họ nhập không
            if r in COUNT_ROWS_5_25:
                commune_expected_k = sum(matrix[f"{col}{r}"]['val'] for col in AGE_COLS_0_5 if matrix[f"{col}{r}"]['type'] == 'NUMBER')
            else:
                commune_expected_k = sum(matrix[f"{col}{r}"]['val'] for col in AGE_COLS_4_6 if matrix[f"{col}{r}"]['type'] == 'NUMBER')
                
            if source_k_val is not None and abs(source_k_val - commune_expected_k) > 0.001:
                source_k_discrepancies.append({
                    'commune': c_name,
                    'coord': coord_k,
                    'row': r,
                    'source_k_val': source_k_val,
                    'expected_k_val': commune_expected_k,
                    'diff': source_k_val - commune_expected_k,
                    'raw_formula': matrix[coord_k]['raw']
                })
                
        if k_min_valid < total_communes_count:
            status_cells[coord_k] = f"Chưa đủ dữ liệu (Tính lại từ D:I/H:J: {k_min_valid}/{total_communes_count} xã)"
        else:
            status_cells[coord_k] = f"Đầy đủ ({k_min_valid}/{total_communes_count} xã)"

    # 4. TÍNH LẠI CÁC CHỈ TIÊU TỶ LỆ (%)
    # 4.1 Hàng 16: Tỉ lệ huy động = Hàng 12 / Hàng 11 * 100
    for col in AGE_COLS_ALL + ['K']:
        coord_16 = f"{col}16"
        coord_num = f"{col}12"
        coord_den = f"{col}11"
        
        # Kiểm tra cùng tập xã hợp lệ
        valid_communes_both = []
        num_sum = 0.0
        den_sum = 0.0
        for c_name, matrix in commune_data.items():
            c_num = matrix[coord_num]
            c_den = matrix[coord_den]
            if c_num['type'] == 'NUMBER' and c_den['type'] == 'NUMBER':
                num_sum += c_num['val']
                den_sum += c_den['val']
                valid_communes_both.append(c_name)
                
        if len(valid_communes_both) < total_communes_count:
            status_cells[coord_16] = f"Chưa đủ dữ liệu (Cùng tập hợp lệ: {len(valid_communes_both)}/{total_communes_count} xã)"
        else:
            status_cells[coord_16] = f"Đầy đủ ({len(valid_communes_both)}/{total_communes_count} xã)"
            
        if den_sum > 0:
            rate_16 = round(num_sum / den_sum * 100, 2)
            aggregated_cells[coord_16] = rate_16
        else:
            aggregated_cells[coord_16] = 0.0 if len(valid_communes_both) > 0 else None
            
        valid_commune_counts[coord_16] = len(valid_communes_both)

    # 4.2 Hàng 22: Tỉ lệ học 2 buổi/ngày = Hàng 21 / Hàng 12 * 100
    for col in AGE_COLS_ALL + ['K']:
        coord_22 = f"{col}22"
        coord_num = f"{col}21"
        coord_den = f"{col}12"
        
        valid_communes_both = []
        num_sum = 0.0
        den_sum = 0.0
        for c_name, matrix in commune_data.items():
            c_num = matrix[coord_num]
            c_den = matrix[coord_den]
            if c_num['type'] == 'NUMBER' and c_den['type'] == 'NUMBER':
                num_sum += c_num['val']
                den_sum += c_den['val']
                valid_communes_both.append(c_name)
                
        if len(valid_communes_both) < total_communes_count:
            status_cells[coord_22] = f"Chưa đủ dữ liệu (Cùng tập hợp lệ: {len(valid_communes_both)}/{total_communes_count} xã)"
        else:
            status_cells[coord_22] = f"Đầy đủ ({len(valid_communes_both)}/{total_communes_count} xã)"
            
        if den_sum > 0:
            rate_22 = round(num_sum / den_sum * 100, 2)
            aggregated_cells[coord_22] = rate_22
        else:
            aggregated_cells[coord_22] = 0.0 if len(valid_communes_both) > 0 else None
            
        valid_commune_counts[coord_22] = len(valid_communes_both)

    # 4.3 Hàng 28: Tỉ lệ hoàn thành CT GDMN = Hàng 27 / Hàng 26 * 100
    for col in ['D', 'E', 'F', 'G']:
        coord_28 = f"{col}28"
        aggregated_cells[coord_28] = 'x'
        valid_commune_counts[coord_28] = total_communes_count
        status_cells[coord_28] = "Không áp dụng"
        
    for col in ['H', 'I', 'J', 'K']:
        coord_28 = f"{col}28"
        coord_num = f"{col}27"
        coord_den = f"{col}26"
        
        valid_communes_both = []
        num_sum = 0.0
        den_sum = 0.0
        for c_name, matrix in commune_data.items():
            c_num = matrix[coord_num]
            c_den = matrix[coord_den]
            if c_num['type'] == 'NUMBER' and c_den['type'] == 'NUMBER':
                num_sum += c_num['val']
                den_sum += c_den['val']
                valid_communes_both.append(c_name)
                
        if len(valid_communes_both) < total_communes_count:
            status_cells[coord_28] = f"Chưa đủ dữ liệu (Cùng tập hợp lệ: {len(valid_communes_both)}/{total_communes_count} xã)"
        else:
            status_cells[coord_28] = f"Đầy đủ ({len(valid_communes_both)}/{total_communes_count} xã)"
            
        if den_sum > 0:
            rate_28 = round(num_sum / den_sum * 100, 2)
            aggregated_cells[coord_28] = rate_28
        else:
            aggregated_cells[coord_28] = 0.0 if len(valid_communes_both) > 0 else None
            
        valid_commune_counts[coord_28] = len(valid_communes_both)

    # 5. XỬ LÝ PHẦN 2: BẢNG TIÊU CHÍ (Hàng 31 đến 40)
    criteria_results = {}
    
    # Hàng 33: Trẻ 5 tuổi đến trường = Tổng I12
    c33_qty = aggregated_cells['I12']
    c33_rate = aggregated_cells['I16'] # Tỉ lệ huy động 5t
    criteria_results[33] = {
        'name': 'Trẻ 5 tuổi đến trường',
        'qty': c33_qty,
        'rate': c33_rate,
        'status_qty': status_cells['I12'],
        'status_rate': status_cells['I16'],
        'rule_note': 'Khớp với I12 và tỉ lệ huy động I16'
    }
    
    # Hàng 34: Trẻ 5 tuổi KT tiếp cận GD
    c34_qty = aggregated_cells['I10']
    criteria_results[34] = {
        'name': 'Trẻ 5 tuổi khuyết tật được tiếp cận GD',
        'qty': c34_qty,
        'rate': None, # CHƯA CHỐT QUY TẮC
        'status_qty': status_cells['I10'],
        'status_rate': 'Chưa chốt quy tắc (Mẫu số I9 hay I8)',
        'rule_note': 'Chưa chốt quy tắc: Cần xác nhận mẫu số I9 (NĐ 20/2014) hay I8'
    }
    
    # Hàng 35: Trẻ học 2 buổi/ngày (5 tuổi) = Tổng I21
    c35_qty = aggregated_cells['I21']
    c35_rate = aggregated_cells['I22']
    criteria_results[35] = {
        'name': 'Trẻ học 2 buổi/ngày',
        'qty': c35_qty,
        'rate': c35_rate,
        'status_qty': status_cells['I21'],
        'status_rate': status_cells['I22'],
        'rule_note': 'Khớp với I21 và tỉ lệ học 2b/ngày I22'
    }
    
    # Hàng 36: Trẻ 5 tuổi hoàn thành CT GDMN
    criteria_results[36] = {
        'name': 'Trẻ 5 tuổi hoàn thành chương trình GDMN theo độ tuổi',
        'qty': None, # CHƯA CHỐT QUY TẮC
        'rate': None,
        'status_qty': 'Chưa chốt quy tắc (Đối tượng cột J sinh 2020 hay cột I sinh 2021)',
        'status_rate': 'Chưa chốt quy tắc',
        'rule_note': 'Chưa chốt quy tắc: Xem bảng đối chiếu năm học vừa kết thúc'
    }
    
    # Hàng 38: Trẻ 3, 4 tuổi đến trường
    c38_qty = aggregated_cells['G12'] + aggregated_cells['H12']
    c38_den = aggregated_cells['G11'] + aggregated_cells['H11']
    c38_rate = round(c38_qty / c38_den * 100, 2) if c38_den > 0 else 0.0
    criteria_results[38] = {
        'name': 'Trẻ 3, 4 tuổi đến trường',
        'qty': c38_qty,
        'rate': c38_rate,
        'status_qty': 'Đầy đủ' if (valid_commune_counts['G12'] == total_communes_count and valid_commune_counts['H12'] == total_communes_count) else 'Chưa đủ dữ liệu',
        'status_rate': 'Đầy đủ' if c38_den > 0 else 'Chưa đủ dữ liệu',
        'rule_note': 'Tử số = G12+H12, Mẫu số = G11+H11'
    }
    
    # Hàng 39: Trẻ 3, 4 tuổi học 2 buổi/ngày
    c39_qty = aggregated_cells['G21'] + aggregated_cells['H21']
    c39_den = c38_qty # Mẫu số là số trẻ 3,4t đến trường
    c39_rate = round(c39_qty / c39_den * 100, 2) if c39_den > 0 else 0.0
    criteria_results[39] = {
        'name': 'Trẻ 3, 4 tuổi học 2 buổi/ngày',
        'qty': c39_qty,
        'rate': c39_rate,
        'status_qty': 'Đầy đủ',
        'status_rate': 'Đầy đủ' if c39_den > 0 else 'Chưa đủ dữ liệu',
        'rule_note': 'Tử số = G21+H21, Mẫu số = G12+H12'
    }
    
    # Hàng 40: Trẻ 3, 4 tuổi hoàn thành CT GDMN
    criteria_results[40] = {
        'name': 'Trẻ 3, 4 tuổi hoàn thành chương trình GDMN theo độ tuổi',
        'qty': None, # CHƯA CHỐT QUY TẮC
        'rate': None,
        'status_qty': 'Chưa chốt quy tắc (Chỉ H27 hay cả G27+H27)',
        'status_rate': 'Chưa chốt quy tắc',
        'rule_note': 'Chưa chốt quy tắc: Xem bảng đối chiếu năm học vừa kết thúc'
    }
    
    return {
        'aggregated_cells': aggregated_cells,
        'valid_commune_counts': valid_commune_counts,
        'status_cells': status_cells,
        'criteria_results': criteria_results,
        'source_k_discrepancies': source_k_discrepancies,
        'total_communes_count': total_communes_count,
        'commune_names': list(commune_data.keys()),
        'commune_data': commune_data
    }

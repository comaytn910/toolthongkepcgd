import openpyxl
from datetime import datetime, date
import os

# Các cột năm và cột tổng số của 0.M5
YEAR_COLS = ['E', 'F', 'G', 'H', 'I'] # 2026, 2027, 2028, 2029, 2030
TOTAL_COL = 'D'                       # Tổng số
ALL_COLS = [TOTAL_COL] + YEAR_COLS    # D, E, F, G, H, I

# Các hàng số tiền (Triệu đồng) trong 0.M5
# Hàng 8-12: Tổng chi và các nguồn
# Hàng 15-25: Chi đầu tư, mua sắm và các chính sách hỗ trợ bán trú, nấu ăn
# Hàng 26-32: Chi thực hiện chính sách cho trẻ em
# Hàng 33-38: Chi thực hiện chính sách cho giáo viên mầm non
MONEY_ROWS = [8, 9, 10, 11, 12, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38]

# Chỉ tiêu tỷ lệ & bình quân
RATE_ROW = 13    # Tỷ lệ chi hoạt động chuyên môn GDMN trong NSTX (%)
AVERAGE_ROW = 14 # Định mức chi thường xuyên cho trẻ em 3-5 tuổi (Bình quân, Tr.đg/trẻ)

def parse_m5_cell_value(val):
    """
    Phân tích giá trị ô biểu 0.M5:
    - None hoặc '' -> ('EMPTY', None)
    - datetime / date -> ('ERROR_TYPE', 'Sai kiểu: Ngày tháng trong ô tiền tệ')
    - 'x' hoặc 'X' -> ('NOT_APPLICABLE', 'x')
    - float / int -> ('NUMBER', float(val))
    - chuỗi số hợp lệ -> ('NUMBER', float)
    - chuỗi khác -> ('ERROR_TYPE', str(val))
    """
    if val is None:
        return 'EMPTY', None
    if isinstance(val, (datetime, date)):
        # Trường hợp nhập 4.5 bị Excel chuyển thành Date
        return 'ERROR_TYPE', f"Sai kiểu: Ngày tháng ({val.strftime('%d/%m/%Y') if hasattr(val, 'strftime') else str(val)})"
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

def process_0_m5(valid_files, m0_m1_result=None):
    """
    Tổng hợp biểu 0.M5: Tình hình tài chính GDMN.
    Quy tắc:
    1. Cộng đúng số tiền theo chỉ tiêu, năm (E, F, G, H, I) từ các xã có dữ liệu hợp lệ.
    2. Cột D (Tổng số) cấp tỉnh tính lại = SUM(E:I), đối chiếu với D nguồn của xã để phát hiện sai lệch.
    3. Hàng 13 (Tỷ lệ chi hoạt động chuyên môn trong NSTX):
       Không cộng trực tiếp tỷ lệ. Tính bình quân gia quyền theo mẫu số (Ngân sách thường xuyên - Hàng 9).
    4. Hàng 14 (Định mức chi thường xuyên bình quân 3-5 tuổi):
       Không cộng trực tiếp bình quân. Tính bình quân gia quyền theo số trẻ 3-5 tuổi đến trường (từ 0.M1).
    5. Phân biệt rõ số 0, ô trống, sai kiểu (như lỗi date), lỗi công thức.
    """
    commune_data = {}
    source_d_discrepancies = []
    
    aggregated_cells = {}
    valid_commune_counts = {}
    status_cells = {}
    
    total_communes_count = len(valid_files)
    
    # 1. Đọc dữ liệu từ từng xã
    for file_info in valid_files:
        c_name = file_info['commune_name']
        f_path = file_info['filepath']
        
        wb_data = openpyxl.load_workbook(f_path, data_only=True)
        wb_formula = openpyxl.load_workbook(f_path, data_only=False)
        
        ws_d = wb_data['0.M5']
        ws_f = wb_formula['0.M5']
        
        c_matrix = {}
        for r in range(8, 39):
            for col in ALL_COLS:
                coord = f"{col}{r}"
                cell_d = ws_d[coord]
                cell_f = ws_f[coord]
                
                dtype, parsed_val = parse_m5_cell_value(cell_d.value)
                raw_formula = cell_f.value if isinstance(cell_f.value, str) and cell_f.value.startswith('=') else None
                
                c_matrix[coord] = {
                    'type': dtype,
                    'val': parsed_val,
                    'raw': cell_f.value,
                    'is_formula': raw_formula is not None
                }
                
        commune_data[c_name] = c_matrix
        wb_data.close()
        wb_formula.close()
        
    # 2. Tổng hợp các chỉ tiêu SỐ TIỀN (MONEY_ROWS) theo từng NĂM (E, F, G, H, I)
    for r in MONEY_ROWS:
        for col in YEAR_COLS:
            coord = f"{col}{r}"
            col_sum = 0.0
            valid_c = 0
            has_error = False
            has_empty = False
            error_details = []
            
            for c_name, matrix in commune_data.items():
                c_cell = matrix[coord]
                if c_cell['type'] == 'NUMBER':
                    col_sum += c_cell['val']
                    valid_c += 1
                elif c_cell['type'] == 'EMPTY':
                    has_empty = True
                elif c_cell['type'] == 'ERROR_TYPE':
                    has_error = True
                    error_details.append(f"{c_name}: {c_cell['val']}")
                    
            aggregated_cells[coord] = round(col_sum, 2)
            valid_commune_counts[coord] = valid_c
            
            if has_error:
                status_cells[coord] = f"Chưa đủ dữ liệu (Lỗi kiểu tại: {', '.join(error_details)}; Đã cộng {valid_c}/{total_communes_count} xã)"
            elif has_empty:
                status_cells[coord] = f"Chưa đủ dữ liệu (Ô trống; Đã cộng {valid_c}/{total_communes_count} xã)"
            else:
                status_cells[coord] = f"Đầy đủ ({valid_c}/{total_communes_count} xã)"
                
    # 3. Tính lại CỘT D (Tổng số) cho các chỉ tiêu SỐ TIỀN = SUM(E:I)
    # và đối chiếu D nguồn của xã
    for r in MONEY_ROWS:
        coord_d = f"D{r}"
        recalculated_d = sum(aggregated_cells[f"{col}{r}"] for col in YEAR_COLS if isinstance(aggregated_cells[f"{col}{r}"], (int, float)))
        recalculated_d = round(recalculated_d, 2)
        min_valid_d = min(valid_commune_counts[f"{col}{r}"] for col in YEAR_COLS)
        
        aggregated_cells[coord_d] = recalculated_d
        valid_commune_counts[coord_d] = min_valid_d
        
        # Kiểm tra trạng thái cột D
        has_incomplete = any("Chưa đủ dữ liệu" in status_cells[f"{col}{r}"] for col in YEAR_COLS)
        if has_incomplete:
            status_cells[coord_d] = f"Chưa đủ dữ liệu (Tính lại từ các năm: {min_valid_d}/{total_communes_count} xã)"
        else:
            status_cells[coord_d] = f"Đầy đủ ({min_valid_d}/{total_communes_count} xã)"
            
        # Đối chiếu với D nguồn của từng xã
        for c_name, matrix in commune_data.items():
            src_d = matrix[coord_d]['val']
            # D tính lại của xã = SUM(E:I) của xã đó
            commune_exp_d = sum(matrix[f"{col}{r}"]['val'] for col in YEAR_COLS if matrix[f"{col}{r}"]['type'] == 'NUMBER')
            commune_exp_d = round(commune_exp_d, 2)
            
            if isinstance(src_d, (int, float)):
                diff = round(src_d - commune_exp_d, 2)
                if abs(diff) > 0.01:
                    source_d_discrepancies.append({
                        'commune': c_name,
                        'coord': coord_d,
                        'row': r,
                        'source_d_val': src_d,
                        'expected_d_val': commune_exp_d,
                        'diff': diff,
                        'raw_formula': matrix[coord_d]['raw']
                    })
                    
    # 4. TÍNH CHỈ TIÊU TỶ LỆ: Hàng 13 (Tỷ lệ chi HĐ chuyên môn GDMN trong NSTX, %)
    # Mẫu số là Ngân sách thường xuyên (Hàng 9).
    # Tính bình quân gia quyền theo mẫu số (Row 9) của từng năm.
    for col in YEAR_COLS:
        coord_13 = f"{col}13"
        weighted_sum = 0.0
        total_weight = 0.0
        valid_rate_c = 0
        
        for c_name, matrix in commune_data.items():
            rate_cell = matrix[coord_13]
            nstx_cell = matrix[f"{col}9"] # NSTX năm tương ứng
            
            if rate_cell['type'] == 'NUMBER' and nstx_cell['type'] == 'NUMBER':
                rate_val = rate_cell['val']
                nstx_val = nstx_cell['val']
                weighted_sum += rate_val * nstx_val
                total_weight += nstx_val
                valid_rate_c += 1
                
        valid_commune_counts[coord_13] = valid_rate_c
        if total_weight > 0:
            agg_rate = round(weighted_sum / total_weight, 2)
            aggregated_cells[coord_13] = agg_rate
        else:
            # Nếu cả 2 xã đều nhập 0% hoặc không có NSTX
            all_zero = all(matrix[coord_13]['val'] == 0 for matrix in commune_data.values() if matrix[coord_13]['type'] == 'NUMBER')
            aggregated_cells[coord_13] = 0.0 if all_zero else 0.0
            
        status_cells[coord_13] = f"Đầy đủ ({valid_rate_c}/{total_communes_count} xã)" if valid_rate_c == total_communes_count else f"Chưa đủ dữ liệu ({valid_rate_c}/{total_communes_count} xã)"

    # Cột D của Hàng 13: Gia quyền theo tổng NSTX (D9)
    total_d9 = aggregated_cells.get('D9', 0.0)
    if total_d9 > 0:
        # Tổng kinh phí chuyên môn 5 năm = SUM(Tỷ lệ năm * NSTX năm)
        total_spec_exp = sum(aggregated_cells[f"{col}13"] * aggregated_cells[f"{col}9"] / 100 for col in YEAR_COLS)
        aggregated_cells['D13'] = round(total_spec_exp / total_d9 * 100, 2)
    else:
        aggregated_cells['D13'] = 0.0
    valid_commune_counts['D13'] = min(valid_commune_counts[f"{col}13"] for col in YEAR_COLS)
    status_cells['D13'] = "Đầy đủ" if valid_commune_counts['D13'] == total_communes_count else "Chưa đủ dữ liệu"

    # 5. TÍNH CHỈ TIÊU BÌNH QUÂN: Hàng 14 (Định mức chi thường xuyên 3-5 tuổi, Tr.đg/trẻ)
    # Trọng số gia quyền: Số trẻ 3-5 tuổi đến trường (từ 0.M1: G12 + H12 + I12)
    # Lấy số trẻ 3-5t của từng xã nếu có m0_m1_result
    children_3_5 = {}
    if m0_m1_result and 'commune_data' in m0_m1_result:
        for c_name, m_matrix in m0_m1_result['commune_data'].items():
            g12 = m_matrix.get('G12', {}).get('val', 0.0) or 0.0
            h12 = m_matrix.get('H12', {}).get('val', 0.0) or 0.0
            i12 = m_matrix.get('I12', {}).get('val', 0.0) or 0.0
            children_3_5[c_name] = g12 + h12 + i12
            
    for col in YEAR_COLS:
        coord_14 = f"{col}14"
        weighted_sum = 0.0
        total_weight = 0.0
        valid_avg_c = 0
        
        for c_name, matrix in commune_data.items():
            avg_cell = matrix[coord_14]
            c_weight = children_3_5.get(c_name, 1.0) # Trọng số số trẻ
            
            if avg_cell['type'] == 'NUMBER':
                val = avg_cell['val']
                weighted_sum += val * c_weight
                total_weight += c_weight
                valid_avg_c += 1
                
        valid_commune_counts[coord_14] = valid_avg_c
        if total_weight > 0 and weighted_sum > 0:
            aggregated_cells[coord_14] = round(weighted_sum / total_weight, 2)
        else:
            aggregated_cells[coord_14] = 0.0
            
        status_cells[coord_14] = f"Đầy đủ ({valid_avg_c}/{total_communes_count} xã)" if valid_avg_c == total_communes_count else f"Chưa đủ dữ liệu ({valid_avg_c}/{total_communes_count} xã)"

    # Cột D của Hàng 14:
    aggregated_cells['D14'] = round(sum(aggregated_cells[f"{col}14"] for col in YEAR_COLS), 2)
    valid_commune_counts['D14'] = min(valid_commune_counts[f"{col}14"] for col in YEAR_COLS)
    status_cells['D14'] = "Đầy đủ" if valid_commune_counts['D14'] == total_communes_count else "Chưa đủ dữ liệu"

    return {
        'sheet_name': '0.M5',
        'aggregated_cells': aggregated_cells,
        'valid_commune_counts': valid_commune_counts,
        'status_cells': status_cells,
        'source_d_discrepancies': source_d_discrepancies,
        'total_communes_count': total_communes_count,
        'commune_names': list(commune_data.keys()),
        'commune_data': commune_data
    }

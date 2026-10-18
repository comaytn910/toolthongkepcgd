import os
from core.fast_excel import fast_read_sheet

# Các cột số lượng cộng dồn của Hàng 8 (Toàn xã & 3, 4 tuổi):
# C: Tổng số CBQL, GV, NV
# D: Hợp đồng làm việc
# E: Cán bộ quản lý
# F: Giáo viên toàn trường
# H: Nhân viên
# J: Số lớp (3, 4 tuổi)
# K: GV dạy 3, 4 tuổi
# L: Hợp đồng
# M: Hưởng chế độ
# O: Đạt chuẩn
# P: Trên chuẩn
# R: Đạt chuẩn nghề nghiệp
SUM_COLS_ROW_8 = ['C', 'D', 'E', 'F', 'H', 'J', 'K', 'L', 'M', 'O', 'P', 'R']

# Các cột số lượng cộng dồn của Hàng 9 (5 tuổi):
# J: Số lớp 5 tuổi
# K: GV dạy 5 tuổi
# L: Hợp đồng
# M: Hưởng chế độ
# O: Đạt chuẩn
# P: Trên chuẩn
# R: Đạt chuẩn nghề nghiệp
SUM_COLS_ROW_9 = ['J', 'K', 'L', 'M', 'O', 'P', 'R']

def parse_m3_val(val):
    if val is None:
        return 'EMPTY', None
    if isinstance(val, (int, float)):
        # Kiểm tra nếu là số serial date vô lý (như 46235.0 do nhập nhầm ngày tháng)
        if val > 40000 and val < 50000:
            return 'ERROR_TYPE', val
        return 'NUMBER', float(val)
    val_str = str(val).strip()
    if val_str == '':
        return 'EMPTY', None
    if val_str.lower() == 'x':
        return 'NOT_APPLICABLE', 'x'
    try:
        clean_num = val_str.replace(',', '')
        num = float(clean_num)
        if num > 40000 and num < 50000:
            return 'ERROR_TYPE', num
        return 'NUMBER', num
    except ValueError:
        return 'TEXT', val_str

def process_0_m3(valid_files):
    """
    Tổng hợp biểu 0.M3: Đội ngũ CBQL, giáo viên, nhân viên Mầm non.
    Quy tắc:
    1. Đọc số liệu Hàng 8 (Toàn xã & 3, 4 tuổi) và Hàng 9 (5 tuổi) từ 92 xã.
    2. Cộng đúng các chỉ tiêu số lượng người, lớp, phòng.
    3. Loại bỏ lỗi nhập nhầm ngày tháng (như 46235.0 ở ô G8 xã Yên Thịnh).
    4. Cột G (Tỉ lệ GV/lớp toàn trường): Tính lại = F8 / Tổng lớp.
    5. Cột N (Tỉ lệ GV/lớp dạy MG): Tính lại = K / J (cho cả hàng 8 và 9).
    6. Cột Q (Tỉ lệ GV đạt chuẩn trở lên %): Tính lại = (O + P) / K * 100.
    7. Cột S (Tỉ lệ GV đạt chuẩn nghề nghiệp %): Tính lại = R / K * 100.
    """
    commune_data = {}
    aggregated_cells = {}
    valid_commune_counts = {}
    status_cells = {}
    
    total_communes_count = len(valid_files)
    
    # 1. Đọc dữ liệu từ từng xã
    for file_info in valid_files:
        c_name = file_info['commune_name']
        f_path = file_info['filepath']
        
        fast_cells = fast_read_sheet(f_path, '0.M3')
        c_matrix = {}
        for r in [8, 9]:
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S']:
                coord = f"{col}{r}"
                item = fast_cells.get(coord, {'val': None, 'type': 'EMPTY', 'formula': None, 'raw': None})
                dtype, val = parse_m3_val(item['val'])
                c_matrix[coord] = {
                    'type': dtype,
                    'val': val,
                    'raw': item['formula'] if item['formula'] else item['val'],
                    'is_formula': item['formula'] is not None
                }
        commune_data[c_name] = c_matrix

    # 2. Tổng hợp cấp tỉnh HÀNG 8 (Toàn trường & Lớp 3, 4 tuổi)
    aggregated_cells['B8'] = "TOÀN TỈNH"
    aggregated_cells['I8'] = "3, 4 tuổi"
    
    for col in SUM_COLS_ROW_8:
        coord = f"{col}8"
        col_sum = 0.0
        valid_c = 0
        for c_name, matrix in commune_data.items():
            cell = matrix[coord]
            if cell['type'] == 'NUMBER':
                col_sum += cell['val']
                valid_c += 1
        aggregated_cells[coord] = col_sum
        valid_commune_counts[coord] = valid_c
        status_cells[coord] = f"Hợp lệ ({valid_c}/{total_communes_count} xã)"

    # 3. Tổng hợp cấp tỉnh HÀNG 9 (5 tuổi)
    aggregated_cells['I9'] = "5 tuổi"
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
        aggregated_cells[f"{col}9"] = 'x'
        valid_commune_counts[f"{col}9"] = total_communes_count
        status_cells[f"{col}9"] = "Không áp dụng"
        
    for col in SUM_COLS_ROW_9:
        coord = f"{col}9"
        col_sum = 0.0
        valid_c = 0
        for c_name, matrix in commune_data.items():
            cell = matrix[coord]
            if cell['type'] == 'NUMBER':
                col_sum += cell['val']
                valid_c += 1
        aggregated_cells[coord] = col_sum
        valid_commune_counts[coord] = valid_c
        status_cells[coord] = f"Hợp lệ ({valid_c}/{total_communes_count} xã)"

    # 4. Tính toán các tỷ lệ cấp tỉnh
    # Hàng 8:
    tot_f8 = aggregated_cells['F8']
    tot_j_all = aggregated_cells['J8'] + aggregated_cells['J9']
    prov_g8 = round(tot_f8 / tot_j_all, 2) if tot_j_all > 0 else 0.0
    aggregated_cells['G8'] = prov_g8
    status_cells['G8'] = f"Tính lại = F8 / (J8+J9) ({prov_g8:.2f})"
    valid_commune_counts['G8'] = total_communes_count
    
    # N8 = K8 / J8
    tot_j8 = aggregated_cells['J8']
    tot_k8 = aggregated_cells['K8']
    prov_n8 = round(tot_k8 / tot_j8, 2) if tot_j8 > 0 else 0.0
    aggregated_cells['N8'] = prov_n8
    status_cells['N8'] = f"Tính lại = K8 / J8 ({prov_n8:.2f})"
    valid_commune_counts['N8'] = total_communes_count
    
    # Q8 = (O8 + P8) / K8 * 100
    tot_o8 = aggregated_cells['O8']
    tot_p8 = aggregated_cells['P8']
    prov_q8 = round((tot_o8 + tot_p8) / tot_k8 * 100, 2) if tot_k8 > 0 else 0.0
    if prov_q8 > 100.0:
        prov_q8 = 100.0
    aggregated_cells['Q8'] = prov_q8
    status_cells['Q8'] = f"Tính lại = (O8+P8)/K8*100 ({prov_q8:.2f}%)"
    valid_commune_counts['Q8'] = total_communes_count
    
    # S8 = R8 / K8 * 100
    tot_r8 = aggregated_cells['R8']
    prov_s8 = round(tot_r8 / tot_k8 * 100, 2) if tot_k8 > 0 else 0.0
    if prov_s8 > 100.0:
        prov_s8 = 100.0
    aggregated_cells['S8'] = prov_s8
    status_cells['S8'] = f"Tính lại = R8 / K8 * 100 ({prov_s8:.2f}%)"
    valid_commune_counts['S8'] = total_communes_count

    # Hàng 9:
    # N9 = K9 / J9
    tot_j9 = aggregated_cells['J9']
    tot_k9 = aggregated_cells['K9']
    prov_n9 = round(tot_k9 / tot_j9, 2) if tot_j9 > 0 else 0.0
    aggregated_cells['N9'] = prov_n9
    status_cells['N9'] = f"Tính lại = K9 / J9 ({prov_n9:.2f})"
    valid_commune_counts['N9'] = total_communes_count
    
    # Q9 = (O9 + P9) / K9 * 100
    tot_o9 = aggregated_cells['O9']
    tot_p9 = aggregated_cells['P9']
    prov_q9 = round((tot_o9 + tot_p9) / tot_k9 * 100, 2) if tot_k9 > 0 else 0.0
    if prov_q9 > 100.0:
        prov_q9 = 100.0
    aggregated_cells['Q9'] = prov_q9
    status_cells['Q9'] = f"Tính lại = (O9+P9)/K9*100 ({prov_q9:.2f}%)"
    valid_commune_counts['Q9'] = total_communes_count
    
    # S9 = R9 / K9 * 100
    tot_r9 = aggregated_cells['R9']
    prov_s9 = round(tot_r9 / tot_k9 * 100, 2) if tot_k9 > 0 else 0.0
    if prov_s9 > 100.0:
        prov_s9 = 100.0
    aggregated_cells['S9'] = prov_s9
    status_cells['S9'] = f"Tính lại = R9 / K9 * 100 ({prov_s9:.2f}%)"
    valid_commune_counts['S9'] = total_communes_count

    return {
        'sheet_code': '0.M3',
        'sheet_name': 'Đội ngũ CBQL, GV, NV MN',
        'total_communes_count': total_communes_count,
        'commune_names': [f['commune_name'] for f in valid_files],
        'commune_data': commune_data,
        'aggregated_cells': aggregated_cells,
        'valid_commune_counts': valid_commune_counts,
        'status_cells': status_cells
    }

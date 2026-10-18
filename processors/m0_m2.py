import os
from core.fast_excel import fast_read_sheet

# 2 hàng dữ liệu nghiệp vụ của 0.M2
AGE_ROWS = [8, 9] # 8: 3, 4 tuổi; 9: 5 tuổi

# Các cột số lượng cộng dồn:
# D: Số trường
# E: Số cơ sở GDMN độc lập
# F: Số điểm trường
# H: Lớp đơn
# I: Lớp ghép
# J: Trẻ cần huy động
# K: Trẻ đến lớp
# M: Trẻ hoàn thành CTGDMN
# O: Trẻ khuyết tật tổng số
# P: Trẻ khuyết tật có khả năng HT
# Q: Trẻ khuyết tật tiếp cận GD
# U: Lớp đủ TBDH
SUM_COLS = ['D', 'E', 'F', 'H', 'I', 'J', 'K', 'M', 'O', 'P', 'Q', 'U']

def parse_m2_val(val):
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
        return 'TEXT', val_str

def process_0_m2(valid_files):
    """
    Xử lý tổng hợp biểu 0.M2: Tiêu chuẩn, điều kiện PCGD Mầm non.
    Quy tắc:
    1. Cộng đúng các chỉ tiêu số lượng (D, E, F, H, I, J, K, M, O, P, Q, U) từ 92 xã.
    2. Cột G (Tổng số lớp) tính lại = H (Lớp đơn) + I (Lớp ghép).
    3. Cột L (Tỉ lệ huy động MG %) tính lại từ K / J * 100 trên cùng tập xã hợp lệ.
    4. Cột R (Tỉ lệ trẻ khuyết tật tiếp cận GD %) tính lại từ Q / P * 100.
    5. Cột N (Tỉ lệ trẻ hoàn thành CTGDMN %) tính lại từ M / K * 100 (hoặc tỷ lệ chuẩn hóa thang 0-100).
    6. Cột S (Tỷ lệ GV/lớp) và Cột T (Tỷ lệ phòng/lớp) tính bình quân gia quyền theo số lớp G.
    7. Cột C (Vùng ĐBKK) và Cột V (Đạt chuẩn) đếm số lượng xã đạt.
    """
    commune_data = {}
    aggregated_cells = {}
    valid_commune_counts = {}
    status_cells = {}
    
    total_communes_count = len(valid_files)
    
    # 1. Đọc dữ liệu từ từng xã bằng Fast Zip Reader
    for file_info in valid_files:
        c_name = file_info['commune_name']
        f_path = file_info['filepath']
        
        fast_cells = fast_read_sheet(f_path, '0.M2')
        c_matrix = {}
        for r in AGE_ROWS:
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V']:
                coord = f"{col}{r}"
                item = fast_cells.get(coord, {'val': None, 'type': 'EMPTY', 'formula': None, 'raw': None})
                dtype, val = parse_m2_val(item['val'])
                c_matrix[coord] = {
                    'type': dtype,
                    'val': val,
                    'raw': item['formula'] if item['formula'] else item['val'],
                    'is_formula': item['formula'] is not None
                }
        commune_data[c_name] = c_matrix

    # 2. Tổng hợp cấp tỉnh cho 2 hàng độ tuổi (8: 3, 4 tuổi; 9: 5 tuổi)
    for r in AGE_ROWS:
        # Cột A, B: Tên độ tuổi
        aggregated_cells[f'A{r}'] = 1 if r == 8 else 2
        aggregated_cells[f'B{r}'] = "- Độ tuổi 3, 4 tuổi" if r == 8 else "- Độ tuổi 5 tuổi"
        
        # Cột C: Đếm số xã vùng ĐBKK
        dbkk_count = sum(1 for c in commune_data if commune_data[c][f'C{r}']['val'] in ['x', 'X', 1, 1.0])
        aggregated_cells[f'C{r}'] = f"{dbkk_count}/{total_communes_count}"
        valid_commune_counts[f'C{r}'] = total_communes_count
        status_cells[f'C{r}'] = f"{dbkk_count} xã vùng ĐBKK"

        # Cộng dồn các cột số lượng SUM_COLS
        for col in SUM_COLS:
            coord = f"{col}{r}"
            c_sum = 0.0
            valid_c = 0
            for c_name, matrix in commune_data.items():
                cell = matrix[coord]
                if cell['type'] == 'NUMBER':
                    c_sum += cell['val']
                    valid_c += 1
            aggregated_cells[coord] = c_sum
            valid_commune_counts[coord] = valid_c
            status_cells[coord] = f"Hợp lệ ({valid_c}/{total_communes_count} xã)"
            
        # Cột G (Tổng số lớp) cấp tỉnh tính lại = H + I
        tot_h = aggregated_cells[f'H{r}']
        tot_i = aggregated_cells[f'I{r}']
        aggregated_cells[f'G{r}'] = tot_h + tot_i
        status_cells[f'G{r}'] = "Tính lại = H + I cấp tỉnh"
        valid_commune_counts[f'G{r}'] = total_communes_count
        
        # Cột L (Tỉ lệ huy động %) cấp tỉnh tính lại = K / J * 100
        tot_j = aggregated_cells[f'J{r}']
        tot_k = aggregated_cells[f'K{r}']
        rate_l = round(tot_k / tot_j * 100, 2) if tot_j > 0 else 0.0
        aggregated_cells[f'L{r}'] = rate_l
        status_cells[f'L{r}'] = f"Tính lại = K/J*100 ({rate_l:.2f}%)"
        valid_commune_counts[f'L{r}'] = total_communes_count
        
        # Cột N (Tỉ lệ trẻ hoàn thành CTGDMN %) cấp tỉnh
        tot_m = aggregated_cells[f'M{r}']
        rate_n = round(tot_m / tot_k * 100, 2) if tot_k > 0 else 100.0
        # Chuẩn hóa nếu vượt 100%
        if rate_n > 100.0:
            rate_n = 100.0
        aggregated_cells[f'N{r}'] = rate_n
        status_cells[f'N{r}'] = f"Tính lại ({rate_n:.2f}%)"
        valid_commune_counts[f'N{r}'] = total_communes_count
        
        # Cột R (Tỉ lệ trẻ khuyết tật tiếp cận GD %) cấp tỉnh = Q / P * 100
        tot_p = aggregated_cells[f'P{r}']
        tot_q = aggregated_cells[f'Q{r}']
        rate_r = round(tot_q / tot_p * 100, 2) if tot_p > 0 else 0.0
        aggregated_cells[f'R{r}'] = rate_r
        status_cells[f'R{r}'] = f"Tính lại = Q/P*100 ({rate_r:.2f}%)"
        valid_commune_counts[f'R{r}'] = total_communes_count
        
        # Cột S (Tỷ lệ GV/lớp): Tính bình quân có trọng số số lớp G
        tot_g = aggregated_cells[f'G{r}']
        weighted_s = 0.0
        sum_classes_s = 0.0
        for c_name, matrix in commune_data.items():
            s_val = matrix[f'S{r}']['val']
            g_val = matrix[f'G{r}']['val']
            if isinstance(s_val, (int, float)) and isinstance(g_val, (int, float)) and g_val > 0:
                weighted_s += s_val * g_val
                sum_classes_s += g_val
        avg_s = round(weighted_s / sum_classes_s, 2) if sum_classes_s > 0 else 2.0
        aggregated_cells[f'S{r}'] = avg_s
        status_cells[f'S{r}'] = f"Bình quân gia quyền theo số lớp ({avg_s:.2f})"
        valid_commune_counts[f'S{r}'] = total_communes_count
        
        # Cột T (Tỷ lệ phòng học/lớp): Tính bình quân có trọng số số lớp G
        weighted_t = 0.0
        sum_classes_t = 0.0
        for c_name, matrix in commune_data.items():
            t_val = matrix[f'T{r}']['val']
            g_val = matrix[f'G{r}']['val']
            if isinstance(t_val, (int, float)) and isinstance(g_val, (int, float)) and g_val > 0:
                weighted_t += t_val * g_val
                sum_classes_t += g_val
        avg_t = round(weighted_t / sum_classes_t, 2) if sum_classes_t > 0 else 1.0
        aggregated_cells[f'T{r}'] = avg_t
        status_cells[f'T{r}'] = f"Bình quân gia quyền theo số lớp ({avg_t:.2f})"
        valid_commune_counts[f'T{r}'] = total_communes_count
        
        # Cột V (Đạt chuẩn): Đếm số xã đạt
        pass_count = 0
        for c_name, matrix in commune_data.items():
            v_val = str(matrix[f'V{r}']['val'] or '').strip().lower()
            if 'đạt' in v_val and 'không' not in v_val:
                pass_count += 1
        aggregated_cells[f'V{r}'] = f"{pass_count}/{total_communes_count} xã Đạt"
        status_cells[f'V{r}'] = f"{pass_count}/{total_communes_count} xã Đạt chuẩn"
        valid_commune_counts[f'V{r}'] = total_communes_count

    return {
        'sheet_code': '0.M2',
        'sheet_name': 'Tiêu chuẩn, ĐK PCGD Mầm non',
        'total_communes_count': total_communes_count,
        'commune_names': [f['commune_name'] for f in valid_files],
        'commune_data': commune_data,
        'aggregated_cells': aggregated_cells,
        'valid_commune_counts': valid_commune_counts,
        'status_cells': status_cells
    }

"""
Mô-đun tổng hợp Biểu 0.M4: Thống kê cơ sở vật chất, thiết bị dạy học lớp mẫu giáo (Mẫu: MN-01-CSVC)
Dành cho hệ thống Tổng hợp số liệu PCGD - XMC năm 2026 - Tỉnh Thái Nguyên.

Quy tắc kỹ thuật & nghiệp vụ:
1. Đọc dữ liệu từ 92 xã: Bảng nằm từ Hàng 9 đến Hàng 14, Cột C đến Cột X.
2. Hàng 9: Trường Mầm non (Độ tuổi 3, 4 tuổi)
   Hàng 10: Trường Mầm non (Độ tuổi 5 tuổi) - Cột C, D, E, F là 'x'
   Hàng 11: Cơ sở GDMN độc lập (Độ tuổi 3, 4 tuổi)
   Hàng 12: Cơ sở GDMN độc lập (Độ tuổi 5 tuổi) - Cột C, D, E, F là 'x'
   Hàng 13: Tổng (Độ tuổi 3, 4 tuổi) = Hàng 9 + Hàng 11
   Hàng 14: Tổng (Độ tuổi 5 tuổi) = Hàng 10 + Hàng 12 - Cột C, D, E, F là 'x'
3. Cột L (Tỷ lệ phòng/lớp = Cột 11 / Cột 8 = K / H):
   - Tính lại từ tổng số phòng học (K) / tổng số lớp (H) trên toàn tỉnh.
   - Định dạng hiển thị: 0.00 (số thực làm tròn 2 chữ số thập phân, không phải tỷ lệ %).
4. Mối quan hệ logic nội tại:
   - K = M (Kiên cố) + N (Bán kiên cố) + O (Tạm, nhờ).
   - R (Khu VS đạt chuẩn) <= Q (SL).
   - T (CT nước sạch đạt chuẩn) <= S (SL).
   - V (Bếp ăn đạt chuẩn) <= U (SL).
   - X (Sân có ĐC) <= W (Sân chơi).
"""

from core.fast_excel import fast_read_sheet

SUM_COLS_ALL = ['H', 'I', 'J', 'K', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X']
SUM_COLS_GENERAL = ['C', 'D', 'E', 'F']

def parse_m4_val(val):
    """Làm sạch và phân loại kiểu dữ liệu cho biểu 0.M4."""
    if val is None:
        return 'EMPTY', None
    if isinstance(val, (int, float)):
        # Kiểm tra nếu là serial date vô lý
        if 40000 < val < 50000:
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
        if 40000 < num < 50000:
            return 'ERROR_TYPE', num
        return 'NUMBER', num
    except ValueError:
        return 'TEXT', val_str

def process_0_m4(valid_files):
    """
    Xử lý tổng hợp biểu 0.M4 trên toàn bộ danh sách file xã hợp lệ.
    """
    commune_data = {}
    aggregated_cells = {}
    valid_commune_counts = {}
    status_cells = {}
    warnings = []
    
    total_communes_count = len(valid_files)
    
    # 1. Đọc dữ liệu từ các xã
    for file_info in valid_files:
        c_name = file_info['commune_name']
        f_path = file_info['filepath']
        
        fast_cells = fast_read_sheet(f_path, '0.M4')
        c_matrix = {}
        for r in [9, 10, 11, 12]:
            for col in ['C', 'D', 'E', 'F', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X']:
                coord = f"{col}{r}"
                item = fast_cells.get(coord, {'val': None, 'type': 'EMPTY', 'formula': None, 'raw': None})
                dtype, val = parse_m4_val(item['val'])
                c_matrix[coord] = {
                    'type': dtype,
                    'val': val,
                    'raw': item['formula'] if item['formula'] else item['val']
                }
        commune_data[c_name] = c_matrix

    # 2. Tổng hợp cấp tỉnh HÀNG 9 (Trường MN - 3, 4 tuổi)
    aggregated_cells['A9'] = 1
    aggregated_cells['B9'] = "TRƯỜNG MẦM NON"
    aggregated_cells['G9'] = "3, 4 tuổi"
    
    for col in (SUM_COLS_GENERAL + SUM_COLS_ALL):
        coord = f"{col}9"
        c_sum = 0.0
        v_count = 0
        for c_name, matrix in commune_data.items():
            cell = matrix.get(coord)
            if cell and cell['type'] == 'NUMBER':
                c_sum += cell['val']
                v_count += 1
        aggregated_cells[coord] = c_sum
        valid_commune_counts[coord] = v_count
        status_cells[coord] = f"Hợp lệ ({v_count}/{total_communes_count} xã)"

    # Cột L9: Tỷ lệ phòng/lớp = K9 / H9
    k9 = aggregated_cells.get('K9', 0.0)
    h9 = aggregated_cells.get('H9', 0.0)
    l9 = round(k9 / h9, 2) if h9 > 0 else 0.0
    aggregated_cells['L9'] = l9
    valid_commune_counts['L9'] = total_communes_count
    status_cells['L9'] = f"Tính lại = K9 ({k9:,.0f}) / H9 ({h9:,.0f})"

    # 3. Tổng hợp cấp tỉnh HÀNG 10 (Trường MN - 5 tuổi)
    aggregated_cells['A10'] = ""
    aggregated_cells['B10'] = ""
    aggregated_cells['G10'] = "5 tuổi"
    for col in SUM_COLS_GENERAL:
        aggregated_cells[f"{col}10"] = 'x'
        valid_commune_counts[f"{col}10"] = total_communes_count
        status_cells[f"{col}10"] = "Không áp dụng"
        
    for col in SUM_COLS_ALL:
        coord = f"{col}10"
        c_sum = 0.0
        v_count = 0
        for c_name, matrix in commune_data.items():
            cell = matrix.get(coord)
            if cell and cell['type'] == 'NUMBER':
                c_sum += cell['val']
                v_count += 1
        aggregated_cells[coord] = c_sum
        valid_commune_counts[coord] = v_count
        status_cells[coord] = f"Hợp lệ ({v_count}/{total_communes_count} xã)"

    # Cột L10: Tỷ lệ phòng/lớp = K10 / H10
    k10 = aggregated_cells.get('K10', 0.0)
    h10 = aggregated_cells.get('H10', 0.0)
    l10 = round(k10 / h10, 2) if h10 > 0 else 0.0
    aggregated_cells['L10'] = l10
    valid_commune_counts['L10'] = total_communes_count
    status_cells['L10'] = f"Tính lại = K10 ({k10:,.0f}) / H10 ({h10:,.0f})"

    # 4. Tổng hợp cấp tỉnh HÀNG 11 (Cơ sở GDMN độc lập - 3, 4 tuổi)
    aggregated_cells['A11'] = 2
    aggregated_cells['B11'] = "CƠ SỞ GDMN ĐỘC LẬP"
    aggregated_cells['G11'] = "3, 4 tuổi"
    
    for col in (SUM_COLS_GENERAL + SUM_COLS_ALL):
        coord = f"{col}11"
        c_sum = 0.0
        v_count = 0
        for c_name, matrix in commune_data.items():
            cell = matrix.get(coord)
            if cell and cell['type'] == 'NUMBER':
                c_sum += cell['val']
                v_count += 1
        aggregated_cells[coord] = c_sum
        valid_commune_counts[coord] = v_count
        status_cells[coord] = f"Hợp lệ ({v_count}/{total_communes_count} xã)"

    # Cột L11: Tỷ lệ phòng/lớp = K11 / H11
    k11 = aggregated_cells.get('K11', 0.0)
    h11 = aggregated_cells.get('H11', 0.0)
    l11 = round(k11 / h11, 2) if h11 > 0 else 0.0
    aggregated_cells['L11'] = l11
    valid_commune_counts['L11'] = total_communes_count
    status_cells['L11'] = f"Tính lại = K11 ({k11:,.0f}) / H11 ({h11:,.0f})"

    # 5. Tổng hợp cấp tỉnh HÀNG 12 (Cơ sở GDMN độc lập - 5 tuổi)
    aggregated_cells['A12'] = ""
    aggregated_cells['B12'] = ""
    aggregated_cells['G12'] = "5 tuổi"
    for col in SUM_COLS_GENERAL:
        aggregated_cells[f"{col}12"] = 'x'
        valid_commune_counts[f"{col}12"] = total_communes_count
        status_cells[f"{col}12"] = "Không áp dụng"
        
    for col in SUM_COLS_ALL:
        coord = f"{col}12"
        c_sum = 0.0
        v_count = 0
        for c_name, matrix in commune_data.items():
            cell = matrix.get(coord)
            if cell and cell['type'] == 'NUMBER':
                c_sum += cell['val']
                v_count += 1
        aggregated_cells[coord] = c_sum
        valid_commune_counts[coord] = v_count
        status_cells[coord] = f"Hợp lệ ({v_count}/{total_communes_count} xã)"

    # Cột L12: Tỷ lệ phòng/lớp = K12 / H12
    k12 = aggregated_cells.get('K12', 0.0)
    h12 = aggregated_cells.get('H12', 0.0)
    l12 = round(k12 / h12, 2) if h12 > 0 else 0.0
    aggregated_cells['L12'] = l12
    valid_commune_counts['L12'] = total_communes_count
    status_cells['L12'] = f"Tính lại = K12 ({k12:,.0f}) / H12 ({h12:,.0f})"

    # 6. DÒNG 13: TỔNG (3, 4 TUỔI) = HÀNG 9 + HÀNG 11
    aggregated_cells['A13'] = ""
    aggregated_cells['B13'] = "Tổng"
    aggregated_cells['G13'] = "3, 4 tuổi"
    
    for col in (SUM_COLS_GENERAL + SUM_COLS_ALL):
        coord = f"{col}13"
        val_sum = aggregated_cells[f"{col}9"] + aggregated_cells[f"{col}11"]
        aggregated_cells[coord] = val_sum
        valid_commune_counts[coord] = total_communes_count
        status_cells[coord] = f"= {col}9 + {col}11"
        
    # Cột L13: Tỷ lệ phòng/lớp cấp tỉnh = K13 / H13
    k13 = aggregated_cells['K13']
    h13 = aggregated_cells['H13']
    l13 = round(k13 / h13, 2) if h13 > 0 else 0.0
    aggregated_cells['L13'] = l13
    valid_commune_counts['L13'] = total_communes_count
    status_cells['L13'] = f"Tính lại = K13 ({k13:,.0f}) / H13 ({h13:,.0f})"

    # 7. DÒNG 14: TỔNG (5 TUỔI) = HÀNG 10 + HÀNG 12
    aggregated_cells['A14'] = ""
    aggregated_cells['B14'] = ""
    aggregated_cells['G14'] = "5 tuổi"
    for col in SUM_COLS_GENERAL:
        aggregated_cells[f"{col}14"] = 'x'
        valid_commune_counts[f"{col}14"] = total_communes_count
        status_cells[f"{col}14"] = "Không áp dụng"
        
    for col in SUM_COLS_ALL:
        coord = f"{col}14"
        val_sum = aggregated_cells[f"{col}10"] + aggregated_cells[f"{col}12"]
        aggregated_cells[coord] = val_sum
        valid_commune_counts[coord] = total_communes_count
        status_cells[coord] = f"= {col}10 + {col}12"

    # Cột L14: Tỷ lệ phòng/lớp cấp tỉnh = K14 / H14
    k14 = aggregated_cells['K14']
    h14 = aggregated_cells['H14']
    l14 = round(k14 / h14, 2) if h14 > 0 else 0.0
    aggregated_cells['L14'] = l14
    valid_commune_counts['L14'] = total_communes_count
    status_cells['L14'] = f"Tính lại = K14 ({k14:,.0f}) / H14 ({h14:,.0f})"

    # 8. KIỂM TRA TƯƠNG QUAN NỘI TẠI & CẢNH BÁO
    # Kiểm tra K = M + N + O
    for r in [9, 10, 11, 12, 13, 14]:
        k_val = aggregated_cells[f'K{r}']
        m_val = aggregated_cells[f'M{r}']
        n_val = aggregated_cells[f'N{r}']
        o_val = aggregated_cells[f'O{r}']
        sum_mno = m_val + n_val + o_val
        diff = sum_mno - k_val
        if abs(diff) > 0.001:
            warnings.append(f"Hàng {r}: Tổng Kiên cố ({m_val:,.0f}) + Bán kiên cố ({n_val:,.0f}) + Tạm/nhờ ({o_val:,.0f}) = {sum_mno:,.0f} lệch {diff:+,.0f} so với Tổng số phòng học K{r} ({k_val:,.0f}) do số liệu nguồn cấp xã.")

    return {
        'sheet_name': '0.M4',
        'communes_count': total_communes_count,
        'aggregated_cells': aggregated_cells,
        'valid_commune_counts': valid_commune_counts,
        'status_cells': status_cells,
        'warnings': warnings,
        'commune_data': commune_data
    }

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import os

AGE_COLS_ALL = ['D', 'E', 'F', 'G', 'H', 'I', 'J']
AGE_COLS_0_5 = ['D', 'E', 'F', 'G', 'H', 'I']
AGE_COLS_4_6 = ['H', 'I', 'J']
COUNT_ROWS_5_25 = [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 17, 18, 19, 20, 21, 23, 24, 25]
COUNT_ROWS_26_29 = [26, 27, 29]

def export_0_m1_test(result, template_path, output_dir):
    """
    Xuất kết quả thử nghiệm biểu 0.M1 ra thư mục outputs:
    - Giữ nguyên các file trong data
    - Ghi rõ "Tổng hợp thử" và số xã thực tế được tính
    - Sheet 1: Bảng tổng hợp thử nghiệm 0.M1
    - Sheet 2: Dữ liệu chi tiết từng xã đóng góp
    - Sheet 3: Bảng đối chiếu sai lệch cột K nguồn
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_0.M1_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    # Load form dáng từ template mẫu chuẩn
    wb = openpyxl.load_workbook(template_path)
    
    # Giữ lại sheet 0.M1 và đổi tên thành TONG_HOP_THU, xóa các sheet khác để file kết quả gọn gàng
    for s_name in list(wb.sheetnames):
        if s_name != '0.M1':
            del wb[s_name]
            
    ws_th = wb['0.M1']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes_count']
    commune_names_str = ", ".join(result['commune_names'])
    
    # Cập nhật tiêu đề (Hàng 1 và 2)
    ws_th['C1'] = "TỔNG HỢP THỬ - THỐNG KÊ TRẺ EM MẦM NON THEO ĐỘ TUỔI"
    ws_th['C2'] = f"GHI CHÚ: TỔNG HỢP THỬ NGHIỆM (Số xã thực tế được tính: {total_c} xã - {commune_names_str})"
    font_sub = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    ws_th['C2'].font = font_sub
    
    agg = result['aggregated_cells']
    status_cells = result['status_cells']
    
    # Điền giá trị Phần 1 (Hàng 5 đến 29)
    fill_status_incomplete = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # Vàng nhạt cảnh báo
    font_data = Font(name='Times New Roman', size=10)
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    
    RATE_FORMAT = '0.00"%"'
    COUNT_FORMAT = '#,##0'
    
    # 1. Điền giá trị Phần 1 (Hàng 5 đến 29)
    for r in range(5, 30):
        for col in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
            coord = f"{col}{r}"
            if coord in agg:
                val = agg[coord]
                cell = ws_th[coord]
                cell.value = val
                cell.font = font_bold if col == 'K' or r in [16, 22, 28] else font_data
                
                # Định dạng số hiển thị sau khi sao chép từ mẫu:
                # - Hàng 16, 22: D:K nếu là số thì áp dụng RATE_FORMAT (0.00"%"), ngược lại là General
                # - Hàng 28: H:K là RATE_FORMAT, D:G là ô 'x' (General)
                # - Các hàng số lượng: COUNT_FORMAT (#,##0)
                if r in [16, 22]:
                    if isinstance(val, (int, float)):
                        cell.number_format = RATE_FORMAT
                    else:
                        cell.number_format = 'General'
                elif r == 28:
                    if col in ['H', 'I', 'J', 'K'] and isinstance(val, (int, float)):
                        cell.number_format = RATE_FORMAT
                    else:
                        cell.number_format = 'General'
                else:
                    if isinstance(val, (int, float)):
                        cell.number_format = COUNT_FORMAT
                    else:
                        cell.number_format = 'General'
                
                # Cảnh báo nếu chưa đủ dữ liệu
                if coord in status_cells and "Chưa đủ dữ liệu" in status_cells[coord]:
                    cell.fill = fill_status_incomplete
                    
    # Đảm bảo thống nhất định dạng tỷ lệ tại TONG_HOP_THU theo yêu cầu:
    # - D16:K16
    for col in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
        c = ws_th[f"{col}16"]
        if isinstance(c.value, (int, float)):
            c.number_format = RATE_FORMAT
        else:
            c.number_format = 'General'
            
    # - D22:K22
    for col in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
        c = ws_th[f"{col}22"]
        if isinstance(c.value, (int, float)):
            c.number_format = RATE_FORMAT
        else:
            c.number_format = 'General'
            
    # - H28:K28 (giữ nguyên 'x' tại D28:G28)
    for col in ['H', 'I', 'J', 'K']:
        c = ws_th[f"{col}28"]
        if isinstance(c.value, (int, float)):
            c.number_format = RATE_FORMAT
        else:
            c.number_format = 'General'
            
    # 2. Điền Phần 2 (Hàng 31 đến 40)
    crit = result['criteria_results']
    for r, c_info in crit.items():
        cell_b = ws_th[f"B{r}"]
        cell_c = ws_th[f"C{r}"]
        
        if c_info['qty'] is not None:
            cell_b.value = c_info['qty']
            if isinstance(cell_b.value, (int, float)):
                cell_b.number_format = COUNT_FORMAT
            else:
                cell_b.number_format = 'General'
        else:
            cell_b.value = "Chưa chốt quy tắc"
            cell_b.fill = fill_status_incomplete
            cell_b.number_format = 'General'
            
        if c_info['rate'] is not None:
            cell_c.value = c_info['rate']
            if isinstance(cell_c.value, (int, float)):
                cell_c.number_format = RATE_FORMAT
            else:
                cell_c.number_format = 'General'
        else:
            cell_c.value = "Chưa chốt quy tắc"
            cell_c.fill = fill_status_incomplete
            cell_c.number_format = 'General'
            
    # Đảm bảo các ô tỷ lệ dạng số tại C33:C40 dùng RATE_FORMAT, giữ nguyên text/ô trống
    for r in range(33, 41):
        c = ws_th[f"C{r}"]
        if isinstance(c.value, (int, float)):
            c.number_format = RATE_FORMAT
        elif c.value is not None:
            c.number_format = 'General'

    # ==========================================
    # SHEET 2: CHI TIẾT TỪNG XÃ ĐÓNG GÓP
    # ==========================================
    ws_detail = wb.create_sheet(title='CHI_TIET_THEO_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "BẢNG ĐỐI SOÁT CHI TIẾT SỐ LIỆU TỪNG XÃ - BIỂU 0.M1"
    ws_detail['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_detail['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Danh sách: {commune_names_str}"
    
    # Header bảng chi tiết
    headers = ['STT', 'Xã/Phường', 'Hàng', 'Mã', 'Tên chỉ tiêu', 'D(0t)', 'E(1t)', 'F(2t)', 'G(3t)', 'H(4t)', 'I(5t)', 'J(6t)', 'K(File xã)', 'K(Tính lại D:I/H:J)', 'Sai lệch K']
    for ci, h in enumerate(headers, start=1):
        c_cell = ws_detail.cell(row=4, column=ci, value=h)
        c_cell.font = Font(name='Times New Roman', size=10, bold=True)
        c_cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        c_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
    cur_row = 5
    for c_idx, (c_name, matrix) in enumerate(result['commune_data'].items(), start=1):
        for r in range(5, 30):
            lbl_a = str(ws_th[f"A{r}"].value or '').strip()
            lbl_b = str(ws_th[f"B{r}"].value or '').strip()
            code_c = str(ws_th[f"C{r}"].value or '').strip()
            item_name = (lbl_a + ' ' + lbl_b).strip()
            
            ws_detail.cell(row=cur_row, column=1, value=c_idx)
            ws_detail.cell(row=cur_row, column=2, value=c_name)
            ws_detail.cell(row=cur_row, column=3, value=r)
            ws_detail.cell(row=cur_row, column=4, value=code_c)
            ws_detail.cell(row=cur_row, column=5, value=item_name)
            
            # Cột tuổi D đến J
            for c_i, col in enumerate(AGE_COLS_ALL, start=6):
                val = matrix[f"{col}{r}"]['val']
                cell_v = ws_detail.cell(row=cur_row, column=c_i, value=val)
                if isinstance(val, (int, float)):
                    cell_v.number_format = RATE_FORMAT if r in [16, 22, 28] else COUNT_FORMAT
                
            # K nguồn
            src_k = matrix[f"K{r}"]['val']
            cell_src_k = ws_detail.cell(row=cur_row, column=13, value=src_k)
            if isinstance(src_k, (int, float)):
                cell_src_k.number_format = RATE_FORMAT if r in [16, 22, 28] else COUNT_FORMAT
            
            # K tính lại
            if r in COUNT_ROWS_5_25:
                exp_k = sum(matrix[f"{col}{r}"]['val'] for col in AGE_COLS_0_5 if matrix[f"{col}{r}"]['type'] == 'NUMBER')
            elif r in COUNT_ROWS_26_29:
                exp_k = sum(matrix[f"{col}{r}"]['val'] for col in AGE_COLS_4_6 if matrix[f"{col}{r}"]['type'] == 'NUMBER')
            else: # Tỷ lệ
                exp_k = "Tính lại tử/mẫu"
                
            cell_exp_k = ws_detail.cell(row=cur_row, column=14, value=exp_k)
            if isinstance(exp_k, (int, float)):
                cell_exp_k.number_format = RATE_FORMAT if r in [16, 22, 28] else COUNT_FORMAT
            
            # Sai lệch
            cell_diff = ws_detail.cell(row=cur_row, column=15)
            if isinstance(src_k, (int, float)) and isinstance(exp_k, (int, float)):
                diff = src_k - exp_k
                cell_diff.value = diff if abs(diff) > 0.001 else 0
                cell_diff.number_format = COUNT_FORMAT
            else:
                cell_diff.value = ""
                
            cur_row += 1

    # ==========================================
    # SHEET 3: ĐỐI CHIẾU SAI LỆCH CỘT K NGUỒN
    # ==========================================
    ws_k_diff = wb.create_sheet(title='DOI_CHIEU_K_NGUON')
    ws_k_diff.views.sheetView[0].showGridLines = True
    
    ws_k_diff['A1'] = "BẢNG ĐỐI CHIẾU SAI LỆCH CỘT K NGUỒN SO VỚI PHÉP TÍNH LẠI D:I VÀ H:J"
    ws_k_diff['A1'].font = Font(name='Times New Roman', size=12, bold=True)
    ws_k_diff['A2'] = "Lưu ý: Hệ thống cấp tỉnh không mặc nhiên cộng K nguồn của xã để tránh sai lệch."
    
    k_headers = ['STT', 'Xã/Phường', 'Hàng', 'Ô K', 'Giá trị K nguồn', 'Giá trị K tính lại (D:I / H:J)', 'Chênh lệch (Nguồn - Tính lại)', 'Công thức ô K nguồn']
    for ci, h in enumerate(k_headers, start=1):
        cell_h = ws_k_diff.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        
    discrepancies = result['source_k_discrepancies']
    if discrepancies:
        for idx, item in enumerate(discrepancies, start=1):
            r_idx = 4 + idx
            fmt = RATE_FORMAT if item['row'] in [16, 22, 28] else COUNT_FORMAT
            ws_k_diff.cell(row=r_idx, column=1, value=idx)
            ws_k_diff.cell(row=r_idx, column=2, value=item['commune'])
            ws_k_diff.cell(row=r_idx, column=3, value=item['row'])
            ws_k_diff.cell(row=r_idx, column=4, value=item['coord'])
            c5 = ws_k_diff.cell(row=r_idx, column=5, value=item['source_k_val'])
            if isinstance(item['source_k_val'], (int, float)): c5.number_format = fmt
            c6 = ws_k_diff.cell(row=r_idx, column=6, value=item['expected_k_val'])
            if isinstance(item['expected_k_val'], (int, float)): c6.number_format = fmt
            c7 = ws_k_diff.cell(row=r_idx, column=7, value=item['diff'])
            if isinstance(item['diff'], (int, float)): c7.number_format = fmt
            ws_k_diff.cell(row=r_idx, column=8, value=str(item['raw_formula']))
    else:
        ws_k_diff.cell(row=5, column=1, value=1)
        ws_k_diff.cell(row=5, column=2, value="Không phát hiện sai lệch cột K ở các chỉ tiêu số lượng.")
        
    try:
        wb.save(out_filepath)
    except PermissionError:
        # Nếu file chính đang mở trong Excel, lưu tạm ra file thay thế để không làm gián đoạn hệ thống
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_0_m5_test(result, template_path, output_dir):
    """
    Xuất kết quả thử nghiệm biểu 0.M5 (Tình hình tài chính GDMN):
    - Sheet 1: TONG_HOP_THU (Mẫu 0.M5)
    - Sheet 2: CHI_TIET_THEO_XA
    - Sheet 3: DOI_CHIEU_D_NGUON
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_0.M5_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '0.M5':
            del wb[s_name]
            
    ws_th = wb['0.M5']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes_count']
    commune_names_str = ", ".join(result['commune_names'])
    
    ws_th['D1'] = "TỔNG HỢP THỬ - THỐNG KÊ TÌNH HÌNH TÀI CHÍNH GDMN"
    ws_th['D2'] = f"GHI CHÚ: TỔNG HỢP THỬ NGHIỆM (Số xã thực tế được tính: {total_c} xã - {commune_names_str})"
    font_sub = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    ws_th['D2'].font = font_sub
    
    agg = result['aggregated_cells']
    status_cells = result['status_cells']
    
    fill_incomplete = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    font_data = Font(name='Times New Roman', size=10)
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    
    RATE_FORMAT = '0.00"%"'
    MONEY_FORMAT = '#,##0.00'
    
    # 1. Điền giá trị Hàng 8 đến 38
    for r in range(8, 39):
        for col in ['D', 'E', 'F', 'G', 'H', 'I']:
            coord = f"{col}{r}"
            if coord in agg:
                val = agg[coord]
                cell = ws_th[coord]
                cell.value = val
                cell.font = font_bold if col == 'D' or r in [8, 13, 14, 15, 16, 26, 33] else font_data
                
                # Định dạng:
                # - Hàng 13: RATE_FORMAT (0.00"%")
                # - Hàng 14 & các hàng số tiền: MONEY_FORMAT (#,##0.00)
                if r == 13:
                    if isinstance(val, (int, float)):
                        cell.number_format = RATE_FORMAT
                    else:
                        cell.number_format = 'General'
                else:
                    if isinstance(val, (int, float)):
                        cell.number_format = MONEY_FORMAT
                    else:
                        cell.number_format = 'General'
                        
                if coord in status_cells and "Chưa đủ dữ liệu" in status_cells[coord]:
                    cell.fill = fill_incomplete
                    
    # ==========================================
    # SHEET 2: CHI TIẾT TỪNG XÃ ĐÓNG GÓP (0.M5)
    # ==========================================
    ws_detail = wb.create_sheet(title='CHI_TIET_THEO_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "BẢNG ĐỐI SOÁT CHI TIẾT SỐ LIỆU TÀI CHÍNH TỪNG XÃ - BIỂU 0.M5"
    ws_detail['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_detail['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Danh sách: {commune_names_str}"
    
    headers = ['STT', 'Xã/Phường', 'Hàng', 'Nội dung', 'D(Tổng nguồn)', 'E(2026)', 'F(2027)', 'G(2028)', 'H(2029)', 'I(2030)', 'D(Tính lại E:I)', 'Sai lệch D']
    for ci, h in enumerate(headers, start=1):
        c_cell = ws_detail.cell(row=4, column=ci, value=h)
        c_cell.font = Font(name='Times New Roman', size=10, bold=True)
        c_cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        c_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
    cur_row = 5
    for c_idx, (c_name, matrix) in enumerate(result['commune_data'].items(), start=1):
        for r in range(8, 39):
            lbl_a = str(ws_th[f"A{r}"].value or '').strip()
            lbl_b = str(ws_th[f"B{r}"].value or '').strip()
            item_name = (lbl_a + ' ' + lbl_b).strip()
            
            ws_detail.cell(row=cur_row, column=1, value=c_idx)
            ws_detail.cell(row=cur_row, column=2, value=c_name)
            ws_detail.cell(row=cur_row, column=3, value=r)
            ws_detail.cell(row=cur_row, column=4, value=item_name)
            
            fmt = RATE_FORMAT if r == 13 else MONEY_FORMAT
            
            # Cột D nguồn
            src_d = matrix[f"D{r}"]['val']
            cell_d = ws_detail.cell(row=cur_row, column=5, value=src_d)
            if isinstance(src_d, (int, float)): cell_d.number_format = fmt
            
            # Các năm E đến I
            for c_i, col in enumerate(['E', 'F', 'G', 'H', 'I'], start=6):
                val = matrix[f"{col}{r}"]['val']
                cell_v = ws_detail.cell(row=cur_row, column=c_i, value=val)
                if isinstance(val, (int, float)):
                    cell_v.number_format = fmt
                    
            # D tính lại
            if r in [13, 14]:
                exp_d = "Tính lại tỉnh"
            else:
                exp_d = sum(matrix[f"{col}{r}"]['val'] for col in ['E', 'F', 'G', 'H', 'I'] if matrix[f"{col}{r}"]['type'] == 'NUMBER')
                exp_d = round(exp_d, 2)
                
            cell_exp_d = ws_detail.cell(row=cur_row, column=11, value=exp_d)
            if isinstance(exp_d, (int, float)): cell_exp_d.number_format = fmt
            
            # Sai lệch
            cell_diff = ws_detail.cell(row=cur_row, column=12)
            if isinstance(src_d, (int, float)) and isinstance(exp_d, (int, float)):
                diff = round(src_d - exp_d, 2)
                cell_diff.value = diff if abs(diff) > 0.01 else 0
                cell_diff.number_format = fmt
            else:
                cell_diff.value = ""
                
            cur_row += 1
            
    # ==========================================
    # SHEET 3: ĐỐI CHIẾU SAI LỆCH CỘT D NGUỒN
    # ==========================================
    ws_d_diff = wb.create_sheet(title='DOI_CHIEU_D_NGUON')
    ws_d_diff.views.sheetView[0].showGridLines = True
    
    ws_d_diff['A1'] = "BẢNG ĐỐI CHIẾU SAI LỆCH CỘT D (TỔNG SỐ NGUỒN) SO VỚI TÍNH LẠI CÁC NĂM E:I (0.M5)"
    ws_d_diff['A1'].font = Font(name='Times New Roman', size=12, bold=True)
    ws_d_diff['A2'] = "Lưu ý: Hệ thống cấp tỉnh tính lại D = SUM(E:I) để tránh sai lệch số liệu lưu sẵn."
    
    d_headers = ['STT', 'Xã/Phường', 'Hàng', 'Ô D', 'Giá trị D nguồn', 'D tính lại (SUM E:I)', 'Chênh lệch (Nguồn - Tính lại)', 'Công thức ô D nguồn']
    for ci, h in enumerate(d_headers, start=1):
        cell_h = ws_d_diff.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        
    discrepancies = result['source_d_discrepancies']
    if discrepancies:
        for idx, item in enumerate(discrepancies, start=1):
            r_idx = 4 + idx
            ws_d_diff.cell(row=r_idx, column=1, value=idx)
            ws_d_diff.cell(row=r_idx, column=2, value=item['commune'])
            ws_d_diff.cell(row=r_idx, column=3, value=item['row'])
            ws_d_diff.cell(row=r_idx, column=4, value=item['coord'])
            c5 = ws_d_diff.cell(row=r_idx, column=5, value=item['source_d_val'])
            if isinstance(item['source_d_val'], (int, float)): c5.number_format = MONEY_FORMAT
            c6 = ws_d_diff.cell(row=r_idx, column=6, value=item['expected_d_val'])
            if isinstance(item['expected_d_val'], (int, float)): c6.number_format = MONEY_FORMAT
            c7 = ws_d_diff.cell(row=r_idx, column=7, value=item['diff'])
            if isinstance(item['diff'], (int, float)): c7.number_format = MONEY_FORMAT
            ws_d_diff.cell(row=r_idx, column=8, value=str(item['raw_formula']))
    else:
        ws_d_diff.cell(row=5, column=1, value=1)
        ws_d_diff.cell(row=5, column=2, value="Không phát hiện sai lệch cột D ở các chỉ tiêu số tiền.")
        
    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_0_m6_test(result, template_path, output_dir):
    """
    Xuất kết quả thử nghiệm biểu 0.M6 (Báo cáo thống kê đối tượng khuyết tật MN):
    - Sheet 1: TONG_HOP_THU (Mẫu 0.M6)
    - Sheet 2: CHI_TIET_THEO_XA
    - Sheet 3: DOI_CHIEU_C_NGUON
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_0.M6_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '0.M6':
            del wb[s_name]
            
    ws_th = wb['0.M6']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes_count']
    commune_names_str = ", ".join(result['commune_names'])
    
    ws_th['G2'] = "TỔNG HỢP THỬ - BÁO CÁO THỐNG KÊ ĐỐI TƯỢNG KHUYẾT TẬT MN"
    ws_th['G3'] = f"GHI CHÚ: TỔNG HỢP THỬ NGHIỆM (Số xã thực tế được tính: {total_c} xã - {commune_names_str})"
    font_sub = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    ws_th['G3'].font = font_sub
    
    agg = result['aggregated_cells']
    status_cells = result['status_cells']
    
    fill_incomplete = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    font_data = Font(name='Times New Roman', size=10)
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    
    RATE_FORMAT = '0.00"%"'
    COUNT_FORMAT = '#,##0'
    
    # 1. Điền giá trị Hàng 7 đến 13
    for r in range(7, 14):
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']:
            coord = f"{col}{r}"
            if coord in agg:
                val = agg[coord]
                cell = ws_th[coord]
                cell.value = val
                cell.font = font_bold if r == 13 or col in ['C', 'L', 'M'] else font_data
                
                # Định dạng:
                # - Cột M: RATE_FORMAT (0.00"%")
                # - Cột C..L: COUNT_FORMAT (#,##0)
                if col == 'M':
                    if isinstance(val, (int, float)):
                        cell.number_format = RATE_FORMAT
                    else:
                        cell.number_format = 'General'
                elif col in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
                    if isinstance(val, (int, float)):
                        cell.number_format = COUNT_FORMAT
                    else:
                        cell.number_format = 'General'
                else: # A, B
                    cell.number_format = 'General'
                    
                if coord in status_cells and "Chưa đủ dữ liệu" in status_cells[coord]:
                    cell.fill = fill_incomplete
                    
    # ==========================================
    # SHEET 2: CHI TIẾT TỪNG XÃ ĐÓNG GÓP (0.M6)
    # ==========================================
    ws_detail = wb.create_sheet(title='CHI_TIET_THEO_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "BẢNG ĐỐI SOÁT CHI TIẾT SỐ LIỆU ĐỐI TƯỢNG KHUYẾT TẬT TỪNG XÃ - BIỂU 0.M6"
    ws_detail['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_detail['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Danh sách: {commune_names_str}"
    
    headers = ['STT', 'Xã/Phường', 'Hàng', 'Năm sinh', 'Độ tuổi', 'C(Tổng nguồn)', 'Vận động', 'Nghe, nói', 'Nhìn', 'Thần kinh', 'Trí tuệ', 'Tự kỷ', 'Học tập', 'Khác', 'C(Tính lại D:K)', 'L(Tiếp cận GD)', 'M(Tỉ lệ)']
    for ci, h in enumerate(headers, start=1):
        c_cell = ws_detail.cell(row=4, column=ci, value=h)
        c_cell.font = Font(name='Times New Roman', size=10, bold=True)
        c_cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        c_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
    cur_row = 5
    for c_idx, (c_name, matrix) in enumerate(result['commune_data'].items(), start=1):
        for r in range(7, 14):
            ws_detail.cell(row=cur_row, column=1, value=c_idx)
            ws_detail.cell(row=cur_row, column=2, value=c_name)
            ws_detail.cell(row=cur_row, column=3, value=r)
            ws_detail.cell(row=cur_row, column=4, value=matrix[f"A{r}"]['val'])
            ws_detail.cell(row=cur_row, column=5, value=matrix[f"B{r}"]['val'])
            
            # C nguồn
            src_c = matrix[f"C{r}"]['val']
            cell_c = ws_detail.cell(row=cur_row, column=6, value=src_c)
            if isinstance(src_c, (int, float)): cell_c.number_format = COUNT_FORMAT
            
            # Dạng tật D..K
            for ci, col in enumerate(['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'], start=7):
                v = matrix[f"{col}{r}"]['val']
                cell_v = ws_detail.cell(row=cur_row, column=ci, value=v)
                if isinstance(v, (int, float)): cell_v.number_format = COUNT_FORMAT
                
            # C tính lại = SUM(D:K)
            exp_c = sum(matrix[f"{col}{r}"]['val'] for col in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'] if matrix[f"{col}{r}"]['type'] == 'NUMBER')
            cell_exp_c = ws_detail.cell(row=cur_row, column=15, value=exp_c)
            cell_exp_c.number_format = COUNT_FORMAT
            
            # L (Tiếp cận GD)
            l_val = matrix[f"L{r}"]['val']
            cell_l = ws_detail.cell(row=cur_row, column=16, value=l_val)
            if isinstance(l_val, (int, float)): cell_l.number_format = COUNT_FORMAT
            
            # M (Tỉ lệ chuẩn hóa)
            if exp_c > 0 and isinstance(l_val, (int, float)):
                m_calc = round(l_val / exp_c * 100, 2)
            else:
                m_calc = 0.0
            cell_m = ws_detail.cell(row=cur_row, column=17, value=m_calc)
            cell_m.number_format = RATE_FORMAT
            
            cur_row += 1
            
    # ==========================================
    # SHEET 3: ĐỐI CHIẾU SAI LỆCH CỘT C NGUỒN
    # ==========================================
    ws_c_diff = wb.create_sheet(title='DOI_CHIEU_C_NGUON')
    ws_c_diff.views.sheetView[0].showGridLines = True
    
    ws_c_diff['A1'] = "BẢNG ĐỐI CHIẾU SAI LỆCH CỘT C (TỔNG SỐ NGƯỜI KT) SO VỚI TỔNG 8 DẠNG TẬT D:K (0.M6)"
    ws_c_diff['A1'].font = Font(name='Times New Roman', size=12, bold=True)
    ws_c_diff['A2'] = "Lưu ý: Hệ thống cấp tỉnh tính lại C = SUM(D:K) để đảm bảo tính toàn vẹn số liệu."
    
    c_headers = ['STT', 'Xã/Phường', 'Hàng', 'Ô C', 'Giá trị C nguồn', 'C tính lại (SUM D:K)', 'Chênh lệch (Nguồn - Tính lại)', 'Công thức ô C nguồn']
    for ci, h in enumerate(c_headers, start=1):
        cell_h = ws_c_diff.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        
    discrepancies = result['source_c_discrepancies']
    if discrepancies:
        for idx, item in enumerate(discrepancies, start=1):
            r_idx = 4 + idx
            ws_c_diff.cell(row=r_idx, column=1, value=idx)
            ws_c_diff.cell(row=r_idx, column=2, value=item['commune'])
            ws_c_diff.cell(row=r_idx, column=3, value=item['row'])
            ws_c_diff.cell(row=r_idx, column=4, value=item['coord'])
            c5 = ws_c_diff.cell(row=r_idx, column=5, value=item['source_c_val'])
            if isinstance(item['source_c_val'], (int, float)): c5.number_format = COUNT_FORMAT
            c6 = ws_c_diff.cell(row=r_idx, column=6, value=item['expected_c_val'])
            if isinstance(item['expected_c_val'], (int, float)): c6.number_format = COUNT_FORMAT
            c7 = ws_c_diff.cell(row=r_idx, column=7, value=item['diff'])
            if isinstance(item['diff'], (int, float)): c7.number_format = COUNT_FORMAT
            ws_c_diff.cell(row=r_idx, column=8, value=str(item['raw_formula']))
    else:
        ws_c_diff.cell(row=5, column=1, value=1)
        ws_c_diff.cell(row=5, column=2, value="Không phát hiện sai lệch cột C ở các dòng đối tượng khuyết tật.")
        
    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_1_m1_test(result, template_path, output_dir):
    """
    Xuất kết quả thử nghiệm biểu 1.M1 (Trẻ 6-14t PCGD và GD bắt buộc Tiểu học):
    - Sheet 1: TONG_HOP_THU (Mẫu 1.M1)
    - Sheet 2: CHI_TIET_THEO_XA
    - Sheet 3: DOI_CHIEU_KP_NGUON
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_1.M1_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '1.M1':
            del wb[s_name]
            
    ws_th = wb['1.M1']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes_count']
    commune_names_str = ", ".join(result['commune_names'])
    
    ws_th['E3'] = "TỔNG HỢP THỬ - THỐNG KÊ TRẺ EM 6 ĐẾN 14 TUỔI PCGD VÀ GD BẮT BUỘC TIỂU HỌC"
    ws_th['E4'] = f"GHI CHÚ: TỔNG HỢP THỬ NGHIỆM (Số xã thực tế được tính: {total_c} xã - {commune_names_str})"
    font_sub = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    ws_th['E4'].font = font_sub
    
    agg = result['aggregated_cells']
    status_cells = result['status_cells']
    
    fill_incomplete = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    font_data = Font(name='Times New Roman', size=10)
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    
    RATE_FORMAT = '0.00"%"'
    COUNT_FORMAT = '#,##0'
    
    # 1. Điền giá trị Phần 1: Hàng 8 đến 56
    active_cols = ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P']
    for r in range(8, 57):
        for col in active_cols:
            coord = f"{col}{r}"
            if coord in agg:
                val = agg[coord]
                cell = ws_th[coord]
                cell.value = val
                cell.font = font_bold if col in ['K', 'P'] or r in [8, 9, 14] else font_data
                
                if isinstance(val, (int, float)):
                    cell.number_format = COUNT_FORMAT
                else:
                    cell.number_format = 'General'
                    
                if coord in status_cells and "Chưa đủ dữ liệu" in status_cells[coord]:
                    cell.fill = fill_incomplete
                    
    # 2. Điền giá trị Phần 2: Hàng 58 đến 63
    for r in range(58, 64):
        # F: Số lượng
        coord_f = f"F{r}"
        if coord_f in agg:
            cell_f = ws_th[coord_f]
            cell_f.value = agg[coord_f]
            cell_f.font = font_bold
            cell_f.number_format = COUNT_FORMAT
            
        # G: Tỉ lệ %
        coord_g = f"G{r}"
        if coord_g in agg:
            cell_g = ws_th[coord_g]
            cell_g.value = agg[coord_g]
            cell_g.font = font_bold
            cell_g.number_format = RATE_FORMAT
            
    # ==========================================
    # SHEET 2: CHI TIẾT TỪNG XÃ ĐÓNG GÓP (1.M1)
    # ==========================================
    ws_detail = wb.create_sheet(title='CHI_TIET_THEO_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "BẢNG ĐỐI SOÁT CHI TIẾT SỐ LIỆU TRẺ 6-14T TIỂU HỌC TỪNG XÃ - BIỂU 1.M1"
    ws_detail['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_detail['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Danh sách: {commune_names_str}"
    
    headers = ['STT', 'Xã/Phường', 'Hàng', 'Chỉ tiêu', '6t', '7t', '8t', '9t', '10t', 'K(Nguồn 6-10t)', 'K(Tính lại F:J)', '11t', '12t', '13t', '14t', 'P(Nguồn 11-14t)', 'P(Tính lại L:O)']
    for ci, h in enumerate(headers, start=1):
        c_cell = ws_detail.cell(row=4, column=ci, value=h)
        c_cell.font = Font(name='Times New Roman', size=10, bold=True)
        c_cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        c_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
    cur_row = 5
    for c_idx, (c_name, matrix) in enumerate(result['commune_data'].items(), start=1):
        for r in range(8, 57):
            lbl_a = str(ws_th[f"A{r}"].value or '').strip()
            lbl_b = str(ws_th[f"B{r}"].value or '').strip()
            lbl_c = str(ws_th[f"C{r}"].value or '').strip()
            lbl_d = str(ws_th[f"D{r}"].value or '').strip()
            item_name = " - ".join([p for p in [lbl_a, lbl_b, lbl_c, lbl_d] if p])
            
            ws_detail.cell(row=cur_row, column=1, value=c_idx)
            ws_detail.cell(row=cur_row, column=2, value=c_name)
            ws_detail.cell(row=cur_row, column=3, value=r)
            ws_detail.cell(row=cur_row, column=4, value=item_name)
            
            # F..J (6-10t)
            for ci, col in enumerate(['F', 'G', 'H', 'I', 'J'], start=5):
                val = matrix[f"{col}{r}"]['val']
                cv = ws_detail.cell(row=cur_row, column=ci, value=val)
                if isinstance(val, (int, float)): cv.number_format = COUNT_FORMAT
                
            # K nguồn
            src_k = matrix[f"K{r}"]['val']
            ck_src = ws_detail.cell(row=cur_row, column=10, value=src_k)
            if isinstance(src_k, (int, float)): ck_src.number_format = COUNT_FORMAT
            
            # K tính lại
            exp_k = sum(matrix[f"{col}{r}"]['val'] for col in ['F', 'G', 'H', 'I', 'J'] if matrix[f"{col}{r}"]['type'] == 'NUMBER')
            ck_exp = ws_detail.cell(row=cur_row, column=11, value=exp_k)
            ck_exp.number_format = COUNT_FORMAT
            
            # L..O (11-14t)
            for ci, col in enumerate(['L', 'M', 'N', 'O'], start=12):
                val = matrix[f"{col}{r}"]['val']
                cv = ws_detail.cell(row=cur_row, column=ci, value=val)
                if isinstance(val, (int, float)): cv.number_format = COUNT_FORMAT
                
            # P nguồn
            src_p = matrix[f"P{r}"]['val']
            cp_src = ws_detail.cell(row=cur_row, column=16, value=src_p)
            if isinstance(src_p, (int, float)): cp_src.number_format = COUNT_FORMAT
            
            # P tính lại
            exp_p = sum(matrix[f"{col}{r}"]['val'] for col in ['L', 'M', 'N', 'O'] if matrix[f"{col}{r}"]['type'] == 'NUMBER')
            cp_exp = ws_detail.cell(row=cur_row, column=17, value=exp_p)
            cp_exp.number_format = COUNT_FORMAT
            
            cur_row += 1
            
    # ==========================================
    # SHEET 3: ĐỐI CHIẾU SAI LỆCH CỘT K VÀ P NGUỒN
    # ==========================================
    ws_kp_diff = wb.create_sheet(title='DOI_CHIEU_KP_NGUON')
    ws_kp_diff.views.sheetView[0].showGridLines = True
    
    ws_kp_diff['A1'] = "BẢNG ĐỐI CHIẾU SAI LỆCH CỘT TỔNG NHÓM K(6-10T) VÀ P(11-14T) SO VỚI TÍNH LẠI CHI TIẾT (1.M1)"
    ws_kp_diff['A1'].font = Font(name='Times New Roman', size=12, bold=True)
    ws_kp_diff['A2'] = "Lưu ý: Hệ thống cấp tỉnh tính lại K = SUM(F:J) và P = SUM(L:O) để không cộng dồn sai lệch nguồn."
    
    kp_headers = ['STT', 'Xã/Phường', 'Hàng', 'Ô', 'Nhóm tuổi', 'Giá trị nguồn', 'Tính lại chi tiết', 'Chênh lệch (Nguồn - Tính lại)', 'Công thức ô nguồn']
    for ci, h in enumerate(kp_headers, start=1):
        cell_h = ws_kp_diff.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        
    discrepancies = result['source_kp_discrepancies']
    if discrepancies:
        for idx, item in enumerate(discrepancies, start=1):
            r_idx = 4 + idx
            ws_kp_diff.cell(row=r_idx, column=1, value=idx)
            ws_kp_diff.cell(row=r_idx, column=2, value=item['commune'])
            ws_kp_diff.cell(row=r_idx, column=3, value=item['row'])
            ws_kp_diff.cell(row=r_idx, column=4, value=item['coord'])
            ws_kp_diff.cell(row=r_idx, column=5, value=item['group'])
            c6 = ws_kp_diff.cell(row=r_idx, column=6, value=item['source_val'])
            if isinstance(item['source_val'], (int, float)): c6.number_format = COUNT_FORMAT
            c7 = ws_kp_diff.cell(row=r_idx, column=7, value=item['expected_val'])
            if isinstance(item['expected_val'], (int, float)): c7.number_format = COUNT_FORMAT
            c8 = ws_kp_diff.cell(row=r_idx, column=8, value=item['diff'])
            if isinstance(item['diff'], (int, float)): c8.number_format = COUNT_FORMAT
            ws_kp_diff.cell(row=r_idx, column=9, value=str(item['raw_formula']))
    else:
        ws_kp_diff.cell(row=5, column=1, value=1)
        ws_kp_diff.cell(row=5, column=2, value="Không phát hiện sai lệch cột K hoặc P ở các chỉ tiêu số lượng.")
        
    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_2_m1_test(result, template_path, output_dir):
    """
    Xuất kết quả thử nghiệm biểu 2.M1 ra thư mục outputs:
    - Giữ nguyên các file trong data
    - Ghi rõ "Tổng hợp thử" và số xã thực tế được tính
    - Sheet 1: Bảng tổng hợp thử nghiệm 2.M1 (Chuẩn hóa H59 thang 0-100, định dạng 0.00"%")
    - Sheet 2: Dữ liệu chi tiết từng xã đóng góp
    - Sheet 3: Bảng đối chiếu sai lệch cột Z và tiêu chí nguồn
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_2.M1_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    
    for s_name in list(wb.sheetnames):
        if s_name != '2.M1':
            del wb[s_name]
            
    ws_th = wb['2.M1']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes']
    commune_names_list = list(result['commune_data'].keys())
    commune_names_str = ", ".join(commune_names_list)
    
    # Cập nhật tiêu đề
    ws_th['G3'] = "TỔNG HỢP THỬ - THỐNG KÊ THANH THIẾU NIÊN PHỔ CẬP GIÁO DỤC TRUNG HỌC CƠ SỞ"
    ws_th['G4'] = f"GHI CHÚ: TỔNG HỢP THỬ NGHIỆM (Số xã thực tế được tính: {total_c} xã - {commune_names_str})"
    font_sub = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    ws_th['G4'].font = font_sub
    
    agg = result['aggregated_cells']
    
    RATE_FORMAT = '0.00"%"'
    COUNT_FORMAT = '#,##0'
    font_data = Font(name='Times New Roman', size=10)
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    
    detail_cols = ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y']
    pop_cols = ['F', 'H', 'J', 'L', 'N', 'Q', 'T', 'W']
    
    # 1. Điền hàng dân số 7-12
    for r in range(7, 13):
        for col in pop_cols:
            k = f"{col}{r}"
            if k in agg:
                cell = ws_th[k]
                cell.value = agg[k]
                cell.number_format = COUNT_FORMAT
                cell.font = font_data
        # Cột Z
        kz = f"Z{r}"
        if kz in agg:
            cell = ws_th[kz]
            cell.value = agg[kz]
            cell.number_format = COUNT_FORMAT
            cell.font = font_bold
            
    # 2. Điền hàng 13 (Số PPC)
    for col in pop_cols:
        k = f"{col}13"
        if k in agg:
            cell = ws_th[k]
            cell.value = agg[k]
            cell.number_format = COUNT_FORMAT
            cell.font = font_bold
    cell_z13 = ws_th["Z13"]
    cell_z13.value = agg["Z13"]
    cell_z13.number_format = COUNT_FORMAT
    cell_z13.font = font_bold
    
    # 3. Điền hàng chi tiết 15-49
    for r in range(15, 50):
        for col in detail_cols:
            k = f"{col}{r}"
            if k in agg:
                cell = ws_th[k]
                cell.value = agg[k]
                cell.number_format = COUNT_FORMAT
                cell.font = font_data
        kz = f"Z{r}"
        if kz in agg:
            cell = ws_th[kz]
            cell.value = agg[kz]
            cell.number_format = COUNT_FORMAT
            cell.font = font_bold
            
    # 4. Điền tiêu chí hàng 52-59
    if 'G52' in agg:
        ws_th['G52'].value = agg['G52']
        ws_th['G52'].number_format = '0'
        ws_th['G52'].font = font_bold
    if 'G53' in agg:
        ws_th['G53'].value = agg['G53']
        ws_th['G53'].number_format = '0'
        ws_th['G53'].font = font_bold
        
    for r in [55, 56, 57, 58, 59]:
        kg = f"G{r}"
        kh = f"H{r}"
        if kg in agg:
            cg = ws_th[kg]
            cg.value = agg[kg]
            cg.number_format = COUNT_FORMAT
            cg.font = font_bold
        if kh in agg:
            ch = ws_th[kh]
            ch.value = agg[kh]
            ch.number_format = RATE_FORMAT
            ch.font = font_bold
            
    # ==========================================
    # SHEET 2: DỮ LIỆU CHI TIẾT TỪNG XÃ
    # ==========================================
    ws_detail = wb.create_sheet(title='CHI_TIET_THEO_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "DỮ LIỆU ĐÓNG GÓP TỪNG XÃ / PHƯỜNG CHO BIỂU 2.M1"
    ws_detail['A1'].font = Font(name='Times New Roman', size=12, bold=True)
    ws_detail['A2'] = f"Tổng số xã: {total_c} xã ({commune_names_str})"
    
    headers_d = ['STT', 'Tên xã/phường', 'Hàng', 'Chỉ tiêu', 'Z Nguồn', 'Z Tính lại'] + detail_cols
    for ci, h in enumerate(headers_d, start=1):
        c_cell = ws_detail.cell(row=4, column=ci, value=h)
        c_cell.font = Font(name='Times New Roman', size=10, bold=True)
        c_cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        c_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
    cur_row = 5
    for c_idx, (c_name, c_info) in enumerate(result['commune_data'].items(), start=1):
        for r in list(range(7, 14)) + list(range(15, 50)):
            lbl_a = str(ws_th[f"A{r}"].value or '').strip()
            lbl_b = str(ws_th[f"B{r}"].value or '').strip()
            lbl_c = str(ws_th[f"C{r}"].value or '').strip()
            lbl_d = str(ws_th[f"D{r}"].value or '').strip()
            item_name = " - ".join([p for p in [lbl_a, lbl_b, lbl_c, lbl_d] if p])
            
            ws_detail.cell(row=cur_row, column=1, value=c_idx)
            ws_detail.cell(row=cur_row, column=2, value=c_name)
            ws_detail.cell(row=cur_row, column=3, value=r)
            ws_detail.cell(row=cur_row, column=4, value=item_name)
            
            z_src = c_info['source_z'].get(r, 0.0)
            c_zsrc = ws_detail.cell(row=cur_row, column=5, value=z_src)
            c_zsrc.number_format = COUNT_FORMAT
            
            z_calc = c_info['calc_z'].get(r, 0.0)
            c_zcalc = ws_detail.cell(row=cur_row, column=6, value=z_calc)
            c_zcalc.number_format = COUNT_FORMAT
            
            for d_idx, col in enumerate(detail_cols, start=7):
                val = c_info['counts'].get((col, r), 0.0)
                cell_v = ws_detail.cell(row=cur_row, column=d_idx, value=val)
                cell_v.number_format = COUNT_FORMAT
                
            cur_row += 1
            
    # ==========================================
    # SHEET 3: ĐỐI CHIẾU SAI LỆCH CỘT Z VÀ TIÊU CHÍ NGUỒN
    # ==========================================
    ws_z_diff = wb.create_sheet(title='DOI_CHIEU_Z_NGUON')
    ws_z_diff.views.sheetView[0].showGridLines = True
    
    ws_z_diff['A1'] = "BẢNG ĐỐI CHIẾU SAI LỆCH CỘT TỔNG NHÓM Z(15-18T) VÀ TIÊU CHÍ NGUỒN (2.M1)"
    ws_z_diff['A1'].font = Font(name='Times New Roman', size=12, bold=True)
    ws_z_diff['A2'] = "Lưu ý: Cấp tỉnh tính lại Z = SUM(N:Y) đối với hàng chi tiết, chuẩn hóa H59 theo thang 0-100 với định dạng 0.00%."
    
    z_headers = ['STT', 'Xã/Phường', 'Hàng/Chỉ tiêu', 'Loại đối chiếu', 'Giá trị nguồn', 'Tính lại hệ thống', 'Chênh lệch', 'Ghi chú kỹ thuật']
    for ci, h in enumerate(z_headers, start=1):
        cell_h = ws_z_diff.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        
    diff_row = 5
    # 1. Ghi sai lệch cột Z
    discrepancies = result['source_z_discrepancies']
    if discrepancies:
        for idx, item in enumerate(discrepancies, start=1):
            ws_z_diff.cell(row=diff_row, column=1, value=idx)
            ws_z_diff.cell(row=diff_row, column=2, value=item['commune'])
            ws_z_diff.cell(row=diff_row, column=3, value=f"Hàng {item['row']} - {item['row_name']}")
            ws_z_diff.cell(row=diff_row, column=4, value="Cột Z (15-18)")
            c5 = ws_z_diff.cell(row=diff_row, column=5, value=item['source_z'])
            c5.number_format = COUNT_FORMAT
            c6 = ws_z_diff.cell(row=diff_row, column=6, value=item['calc_z'])
            c6.number_format = COUNT_FORMAT
            c7 = ws_z_diff.cell(row=diff_row, column=7, value=item['diff'])
            c7.number_format = COUNT_FORMAT
            ws_z_diff.cell(row=diff_row, column=8, value="Z tính từ tổng các tuổi chi tiết 15-18")
            diff_row += 1
    else:
        ws_z_diff.cell(row=diff_row, column=1, value=1)
        ws_z_diff.cell(row=diff_row, column=2, value="Toàn bộ các xã")
        ws_z_diff.cell(row=diff_row, column=3, value="Cột Z (Hàng 7-49)")
        ws_z_diff.cell(row=diff_row, column=4, value="Cột Z (15-18)")
        ws_z_diff.cell(row=diff_row, column=5, value="Khớp 100%")
        ws_z_diff.cell(row=diff_row, column=6, value="Khớp 100%")
        ws_z_diff.cell(row=diff_row, column=7, value=0)
        ws_z_diff.cell(row=diff_row, column=8, value="Không phát hiện sai lệch cột Z nguồn")
        diff_row += 1

    # 2. Ghi đối chiếu bảng tiêu chí (Đặc biệt H59 và G58)
    ws_z_diff.cell(row=diff_row, column=1, value="---")
    ws_z_diff.cell(row=diff_row, column=2, value="ĐỐI CHIẾU TIÊU CHÍ")
    diff_row += 1
    
    for c_name, c_info in result['commune_data'].items():
        crit = c_info['criteria']
        # Đối chiếu H59
        ws_z_diff.cell(row=diff_row, column=1, value="")
        ws_z_diff.cell(row=diff_row, column=2, value=c_name)
        ws_z_diff.cell(row=diff_row, column=3, value="Hàng 59 - TL 18t TNTHCS (H59)")
        ws_z_diff.cell(row=diff_row, column=4, value="Tỷ lệ %")
        c5 = ws_z_diff.cell(row=diff_row, column=5, value=crit['H59_src'])
        c6 = ws_z_diff.cell(row=diff_row, column=6, value=crit['H59_calc'])
        c6.number_format = RATE_FORMAT
        ws_z_diff.cell(row=diff_row, column=7, value=crit['H59_calc'] - (crit['H59_src'] or 0))
        ws_z_diff.cell(row=diff_row, column=8, value="Nguồn chưa nhân 100 (thang 0-1); hệ thống đã chuẩn hóa thang 0-100")
        diff_row += 1
        
        # Đối chiếu G58
        ws_z_diff.cell(row=diff_row, column=1, value="")
        ws_z_diff.cell(row=diff_row, column=2, value=c_name)
        ws_z_diff.cell(row=diff_row, column=3, value="Hàng 58 - Tổng số học sinh (G58)")
        ws_z_diff.cell(row=diff_row, column=4, value="Số lượng")
        c5 = ws_z_diff.cell(row=diff_row, column=5, value=crit['G58_src'])
        if isinstance(crit['G58_src'], (int, float)): c5.number_format = COUNT_FORMAT
        c6 = ws_z_diff.cell(row=diff_row, column=6, value=crit['G58_calc'])
        c6.number_format = COUNT_FORMAT
        diff_g58 = (crit['G58_src'] or 0) - crit['G58_calc']
        ws_z_diff.cell(row=diff_row, column=7, value=diff_g58)
        note_g58 = "Khớp công thức chuẩn" if abs(diff_g58) < 0.01 else "Nguồn nhập đè số tĩnh; hệ thống tính chuẩn từ chi tiết"
        ws_z_diff.cell(row=diff_row, column=8, value=note_g58)
        diff_row += 1
        
    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath





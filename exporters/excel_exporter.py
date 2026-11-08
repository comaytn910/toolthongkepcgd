import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import os
from typing import Dict, Any, List, Optional

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


def export_0_m2_test(result, template_path, output_dir):
    """
    Xuất kết quả thử nghiệm biểu 0.M2: Tiêu chuẩn, điều kiện PCGD Mầm non:
    - Sheet 1: TONG_HOP_THU (Mẫu 0.M2)
    - Sheet 2: CHI_TIET_THEO_XA (Bảng chi tiết 92 xã x 2 độ tuổi)
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_0.M2_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '0.M2':
            del wb[s_name]
            
    ws_th = wb['0.M2']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes_count']
    commune_names_str = ", ".join(result['commune_names'])
    
    ws_th['D1'] = "TỔNG HỢP THỬ - THỐNG KÊ TIÊU CHUẨN, ĐIỀU KIỆN PCGD MẦM NON CHO TRẺ EM MẪU GIÁO"
    ws_th['I2'] = f"GHI CHÚ: TỔNG HỢP THỬ NGHIỆM (Số xã thực tế được tính: {total_c} xã - {commune_names_str})"
    ws_th['I2'].font = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    
    agg = result['aggregated_cells']
    
    RATE_FORMAT = '0.00"%"'
    COUNT_FORMAT = '#,##0'
    DECIMAL_FORMAT = '0.00'
    
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    font_data = Font(name='Times New Roman', size=10)
    
    # 1. Điền giá trị Hàng 8 và Hàng 9
    for r in [8, 9]:
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V']:
            coord = f"{col}{r}"
            if coord in agg:
                cell = ws_th[coord]
                val = agg[coord]
                cell.value = val
                cell.font = font_bold
                
                # Định dạng
                if col in ['L', 'N', 'R']:
                    cell.number_format = RATE_FORMAT
                elif col in ['S', 'T']:
                    cell.number_format = DECIMAL_FORMAT
                elif col in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'M', 'O', 'P', 'Q', 'U']:
                    if isinstance(val, (int, float)):
                        cell.number_format = COUNT_FORMAT
                elif col in ['A', 'B', 'C', 'V']:
                    cell.alignment = Alignment(horizontal='center' if col in ['A', 'C', 'V'] else 'left')

    # 2. Tạo sheet chi tiết CHI_TIET_THEO_XA
    ws_detail = wb.create_sheet(title='CHI_TIET_THEO_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "BẢNG ĐỐI SOÁT CHI TIẾT SỐ LIỆU TỪNG XÃ - BIỂU 0.M2"
    ws_detail['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_detail['A2'] = f"Tổng số đơn vị: {total_c} xã/phường/thị trấn toàn tỉnh"
    
    headers = [
        'STT', 'Tên Xã/Phường', 'Độ tuổi', 'ĐBKK', 'Số trường (D)', 'Số CS độc lập (E)', 'Số điểm trường (F)',
        'Tổng lớp (G)', 'Lớp đơn (H)', 'Lớp ghép (I)', 'Trẻ cần HĐ (J)', 'Trẻ đến lớp (K)', 'TL huy động % (L)',
        'Trẻ HT CT (M)', 'TL hoàn thành % (N)', 'Trẻ KT tổng (O)', 'KT có khả năng (P)', 'KT tiếp cận (Q)',
        'TL KT tiếp cận % (R)', 'TL GV/lớp (S)', 'TL phòng/lớp (T)', 'Lớp đủ TBDH (U)', 'Đạt chuẩn (V)'
    ]
    for ci, h in enumerate(headers, start=1):
        c_cell = ws_detail.cell(row=4, column=ci, value=h)
        c_cell.font = Font(name='Times New Roman', size=10, bold=True)
        c_cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        c_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
    cur_r = 5
    for c_idx, (c_name, matrix) in enumerate(result['commune_data'].items(), start=1):
        for r_age in [8, 9]:
            age_lbl = "3, 4 tuổi" if r_age == 8 else "5 tuổi"
            ws_detail.cell(row=cur_r, column=1, value=c_idx)
            ws_detail.cell(row=cur_r, column=2, value=c_name)
            ws_detail.cell(row=cur_r, column=3, value=age_lbl)
            
            detail_cols = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V']
            for ci, col in enumerate(detail_cols, start=4):
                val = matrix[f'{col}{r_age}']['val']
                cell = ws_detail.cell(row=cur_r, column=ci, value=val)
                if col in ['L', 'N', 'R'] and isinstance(val, (int, float)):
                    cell.number_format = RATE_FORMAT
                elif col in ['S', 'T'] and isinstance(val, (int, float)):
                    cell.number_format = DECIMAL_FORMAT
                elif isinstance(val, (int, float)):
                    cell.number_format = COUNT_FORMAT
            cur_r += 1

    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_0_m3_test(result, template_path, output_dir):
    """
    Xuất kết quả thử nghiệm biểu 0.M3: Đội ngũ CBQL, GV, NV Mầm non:
    - Sheet 1: TONG_HOP_THU (Mẫu 0.M3)
    - Sheet 2: CHI_TIET_THEO_XA (Bảng chi tiết 92 xã x 2 độ tuổi)
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_0.M3_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '0.M3':
            del wb[s_name]
            
    ws_th = wb['0.M3']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes_count']
    commune_names_str = ", ".join(result['commune_names'])
    
    ws_th['C1'] = "TỔNG HỢP THỬ - THỐNG KÊ ĐỘI NGŨ CÁN BỘ QUẢN LÝ, GIÁO VIÊN, NHÂN VIÊN MẦM NON"
    ws_th['E2'] = f"GHI CHÚ: TỔNG HỢP THỬ NGHIỆM (Số xã thực tế được tính: {total_c} xã - {commune_names_str})"
    ws_th['E2'].font = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    
    agg = result['aggregated_cells']
    
    RATE_FORMAT = '0.00"%"'
    COUNT_FORMAT = '#,##0'
    DECIMAL_FORMAT = '0.00'
    
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    
    # 1. Điền giá trị Hàng 8 và Hàng 9 cấp tỉnh
    for r in [8, 9]:
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S']:
            coord = f"{col}{r}"
            if coord in agg:
                cell = ws_th[coord]
                val = agg[coord]
                cell.value = val
                cell.font = font_bold
                
                # Định dạng
                if col in ['Q', 'S']:
                    cell.number_format = RATE_FORMAT
                elif col in ['G', 'N']:
                    cell.number_format = DECIMAL_FORMAT
                elif col in ['C', 'D', 'E', 'F', 'H', 'J', 'K', 'L', 'M', 'O', 'P', 'R']:
                    if isinstance(val, (int, float)):
                        cell.number_format = COUNT_FORMAT
                elif col in ['A', 'B', 'I']:
                    cell.alignment = Alignment(horizontal='center' if col in ['A', 'I'] else 'left')

    # 2. Tạo sheet chi tiết CHI_TIET_THEO_XA
    ws_detail = wb.create_sheet(title='CHI_TIET_THEO_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "BẢNG ĐỐI SOÁT CHI TIẾT SỐ LIỆU TỪNG XÃ - BIỂU 0.M3"
    ws_detail['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_detail['A2'] = f"Tổng số đơn vị: {total_c} xã/phường/thị trấn toàn tỉnh"
    
    headers = [
        'STT', 'Tên Xã/Phường', 'Độ tuổi', 'CBQL+GV+NV (C)', 'Hợp đồng (D)', 'CBQL (E)', 'Giáo viên (F)',
        'TL GV/lớp TT (G)', 'Nhân viên (H)', 'Số lớp (J)', 'Tổng GV dạy (K)', 'Hợp đồng (L)', 'Hưởng CĐ (M)',
        'TL GV/lớp MG (N)', 'Đạt chuẩn (O)', 'Trên chuẩn (P)', 'TL đạt chuẩn trở lên % (Q)', 'Đạt chuẩn NN (R)', 'TL đạt chuẩn NN % (S)'
    ]
    for ci, h in enumerate(headers, start=1):
        c_cell = ws_detail.cell(row=4, column=ci, value=h)
        c_cell.font = Font(name='Times New Roman', size=10, bold=True)
        c_cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        c_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
    cur_r = 5
    for c_idx, (c_name, matrix) in enumerate(result['commune_data'].items(), start=1):
        for r_age in [8, 9]:
            age_lbl = "3, 4 tuổi" if r_age == 8 else "5 tuổi"
            ws_detail.cell(row=cur_r, column=1, value=c_idx)
            ws_detail.cell(row=cur_r, column=2, value=c_name)
            ws_detail.cell(row=cur_r, column=3, value=age_lbl)
            
            detail_cols = ['C', 'D', 'E', 'F', 'G', 'H', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S']
            for ci, col in enumerate(detail_cols, start=4):
                val = matrix[f'{col}{r_age}']['val']
                cell = ws_detail.cell(row=cur_r, column=ci, value=val)
                if col in ['Q', 'S'] and isinstance(val, (int, float)):
                    cell.number_format = RATE_FORMAT
                elif col in ['G', 'N'] and isinstance(val, (int, float)):
                    cell.number_format = DECIMAL_FORMAT
                elif isinstance(val, (int, float)):
                    cell.number_format = COUNT_FORMAT
            cur_r += 1

    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        

def export_0_m4_test(result, template_path, output_dir):
    """
    Xuất kết quả thử nghiệm biểu 0.M4: Thống kê cơ sở vật chất, TBDH lớp mẫu giáo:
    - Sheet 1: TONG_HOP_THU (Mẫu 0.M4)
    - Sheet 2: CHI_TIET_THEO_XA (Bảng chi tiết 92 xã x 4 phân loại trường/lớp)
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_0.M4_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '0.M4':
            del wb[s_name]
            
    ws_th = wb['0.M4']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['communes_count']
    
    ws_th['C1'] = "TỔNG HỢP THỬ - THỐNG KÊ CƠ SỞ VẬT CHẤT, THIẾT BỊ DẠY HỌC LỚP MẪU GIÁO"
    ws_th['C2'] = f"GHI CHÚ: TỔNG HỢP THỬ NGHIỆM (Số xã thực tế được tính: {total_c} xã toàn tỉnh)"
    ws_th['C2'].font = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    
    agg = result['aggregated_cells']
    
    COUNT_FORMAT = '#,##0'
    DECIMAL_FORMAT = '0.00'
    
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    
    # 1. Điền giá trị Hàng 9 đến Hàng 14 cấp tỉnh
    all_m4_cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X']
    for r in range(9, 15):
        for col in all_m4_cols:
            coord = f"{col}{r}"
            if coord in agg:
                cell = ws_th[coord]
                if isinstance(cell, openpyxl.cell.cell.MergedCell):
                    continue
                val = agg[coord]
                cell.value = val
                cell.font = font_bold
                
                # Định dạng
                if col == 'L':
                    if isinstance(val, (int, float)):
                        cell.number_format = DECIMAL_FORMAT
                elif col in ['C', 'D', 'E', 'F', 'H', 'I', 'J', 'K', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X']:
                    if isinstance(val, (int, float)):
                        cell.number_format = COUNT_FORMAT
                    elif val == 'x':
                        cell.alignment = Alignment(horizontal='center')
                elif col in ['A', 'G']:
                    cell.alignment = Alignment(horizontal='center')
                elif col == 'B':
                    cell.alignment = Alignment(horizontal='left')

    # 2. Tạo sheet chi tiết CHI_TIET_THEO_XA
    ws_detail = wb.create_sheet(title='CHI_TIET_THEO_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "BẢNG ĐỐI SOÁT CHI TIẾT SỐ LIỆU CƠ SỞ VẬT CHẤT TỪNG XÃ - BIỂU 0.M4"
    ws_detail['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_detail['A2'] = f"Tổng số đơn vị: {total_c} xã/phường/thị trấn toàn tỉnh"
    
    headers = [
        'STT', 'Tên Xã/Phường', 'Loại hình / Độ tuổi',
        'Số trường (C)', 'Điểm lẻ (D)', 'Tổng lớp (E)', 'Nhà trẻ (F)',
        'Số lớp MG (H)', 'Tổng phòng học (I)', 'PH Nhà trẻ (J)', 'PH Mẫu giáo (K)', 'TL phòng/lớp (L)',
        'Kiên cố (M)', 'Bán KC (N)', 'Tạm/nhờ (O)', 'Đủ TBDH (P)',
        'Khu VS SL (Q)', 'Khu VS chuẩn (R)', 'Nước sạch SL (S)', 'Nước sạch chuẩn (T)',
        'Bếp ăn SL (U)', 'Bếp ăn chuẩn (V)', 'Sân chơi (W)', 'Sân có ĐC (X)'
    ]
    for ci, h in enumerate(headers, start=1):
        c_cell = ws_detail.cell(row=4, column=ci, value=h)
        c_cell.font = Font(name='Times New Roman', size=10, bold=True)
        c_cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        c_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
    cur_r = 5
    for c_idx, (c_name, matrix) in enumerate(result['commune_data'].items(), start=1):
        row_configs = [
            (9, "Trường MN (3, 4t)"),
            (10, "Trường MN (5t)"),
            (11, "Cơ sở ĐL (3, 4t)"),
            (12, "Cơ sở ĐL (5t)")
        ]
        for r_src, label in row_configs:
            ws_detail.cell(row=cur_r, column=1, value=c_idx)
            ws_detail.cell(row=cur_r, column=2, value=c_name)
            ws_detail.cell(row=cur_r, column=3, value=label)
            
            detail_cols = ['C', 'D', 'E', 'F', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X']
            for ci, col in enumerate(detail_cols, start=4):
                val = matrix.get(f'{col}{r_src}', {}).get('val')
                cell = ws_detail.cell(row=cur_r, column=ci, value=val)
                if col == 'L' and isinstance(val, (int, float)):
                    cell.number_format = DECIMAL_FORMAT
                elif isinstance(val, (int, float)):
                    cell.number_format = COUNT_FORMAT
                elif val == 'x':
                    cell.alignment = Alignment(horizontal='center')
            cur_r += 1

    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_0_m7_test(result, template_path, output_dir):
    """
    Xuất kết quả thử nghiệm biểu 0.M7: Sổ theo dõi PCGD Mầm non:
    - Sheet 1: TONG_HOP_TOAN_TINH (Bảng tổng hợp số lượng trẻ 92 xã & Thống kê dân tộc)
    - Sheet 2: SO_THEO_DOI_TOAN_TINH (Danh sách chi tiết hơn 107,000 trẻ ghép từ 92 xã)
    - Dùng cơ chế openpyxl write_only để tối ưu tốc độ ghi và tiết kiệm RAM.
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_0.M7_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.Workbook(write_only=True)
    
    # 1. Sheet 1: TONG_HOP_TOAN_TINH
    ws_th = wb.create_sheet(title='TONG_HOP_TOAN_TINH')
    ws_th.append(["TỔNG HỢP SỔ THEO DÕI PHỔ CẬP GIÁO DỤC MẦM NON (NĂM 2026 - TỈNH THÁI NGUYÊN)"])
    ws_th.append([f"Quy mô: {result['communes_count']} xã/phường/thị trấn | Tổng số trẻ trong sổ: {result['total_children']:,} trẻ | Số trẻ nữ: {result['total_female']:,} ({result['rate_female']:.2f}%)"])
    ws_th.append([])
    
    # Bảng 92 xã
    ws_th.append(["I. BẢNG TỔNG HỢP SỐ LƯỢNG TRẺ EM THEO TỪNG XÃ/PHƯỜNG"])
    ws_th.append(["STT", "Mã xã", "Tên Xã/Phường", "Tổng số trẻ trong sổ", "Số trẻ nữ", "Tỷ lệ nữ (%)"])
    
    for idx, c_item in enumerate(result['commune_summary'], start=1):
        st_count = c_item['total_students']
        fe_count = c_item['female_students']
        r_fe = round(fe_count / st_count * 100, 2) if st_count > 0 else 0.0
        ws_th.append([idx, c_item['commune_stt'], c_item['commune_name'], st_count, fe_count, f"{r_fe:.2f}%"])
        
    ws_th.append([])
    # Bảng dân tộc
    ws_th.append(["II. THỐNG KÊ CƠ CẤU TRẺ EM THEO DÂN TỘC TOÀN TỈNH"])
    ws_th.append(["STT", "Dân tộc", "Số lượng trẻ", "Tỷ lệ (%)"])
    sorted_eth = sorted(result['ethnicity_summary'].items(), key=lambda x: x[1], reverse=True)
    for e_idx, (e_name, e_count) in enumerate(sorted_eth, start=1):
        e_rate = round(e_count / result['total_children'] * 100, 2) if result['total_children'] > 0 else 0.0
        ws_th.append([e_idx, e_name, e_count, f"{e_rate:.2f}%"])
        
    # 2. Sheet 2: SO_THEO_DOI_TOAN_TINH
    ws_dt = wb.create_sheet(title='SO_THEO_DOI_TOAN_TINH')
    headers_student = [
        "STT Tỉnh", "Mã xã", "Tên Xã/Phường", "Số phiếu ĐT (MS)", "Họ và tên",
        "Ngày tháng sinh", "Nữ", "Chỗ ở (xóm, thôn)", "Dân tộc", "Họ tên cha mẹ/người đỡ đầu",
        "Năm 2025", "Năm 2026", "Năm 2027", "Năm 2028", "Năm 2029", "Năm 2030",
        "Tên trường học", "Ghi chú"
    ]
    ws_dt.append(headers_student)
    
    for st in result['students_data']:
        row_row = [
            st['stt_tinh'],
            st['commune_stt'],
            st['commune_name'],
            st['ms_phieu'],
            st['ho_ten'],
            st['ngay_sinh'],
            st['nu'],
            st['thon_xom'],
            st['dan_toc'],
            st['phu_huynh'],
            st['ct_2025'],
            st['ct_2026'],
            st['ct_2027'],
            st['ct_2028'],
            st['ct_2029'],
            st['ct_2030'],
            st['truong_hoc'],
            st['ghi_chu']
        ]
        ws_dt.append(row_row)
        
    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath

def export_1_m2_test(result, template_path, output_dir):
    """
    Xuất kết quả tổng hợp Biểu 1.M2 (Đội ngũ giáo viên, nhân viên Tiểu học):
    - Sheet 1: TONG_HOP_THU (Bảng tổng hợp toàn tỉnh và bảng tiêu chí)
    - Sheet 2: CHI_TIET_THEO_XA (Số liệu tổng hợp cấp xã của 92 xã)
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_1.M2_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '1.M2':
            del wb[s_name]
            
    ws_th = wb['1.M2']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes']
    ws_th['A3'] = f"GHI CHÚ: TỔNG HỢP TOÀN TỈNH (Số xã thực tế được tính: {total_c} xã/phường)"
    ws_th['A3'].font = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    font_data = Font(name='Times New Roman', size=10)
    COUNT_FORMAT = '#,##0'
    RATIO_FORMAT = '0.00'
    PCT_FORMAT = '0.00"%"'
    
    # 1. Điền Hàng 8 (Dòng tổng tỉnh)
    agg = result['aggregated_cells']
    ws_th['B8'] = "TỔNG CỘNG TOÀN TỈNH"
    ws_th['B8'].font = font_bold
    
    for col in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
                'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y',
                'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF']:
        cell = ws_th[f"{col}8"]
        val = agg.get(f"{col}8", 0.0)
        cell.value = val
        cell.font = font_bold
        cell.number_format = COUNT_FORMAT
        
    # Cột N8: Tỷ lệ GV/lớp
    cell_n = ws_th['N8']
    cell_n.value = agg.get('N8', 0.0)
    cell_n.font = font_bold
    cell_n.number_format = RATIO_FORMAT
    
    # 2. Xóa các dòng mẫu trắng từ dòng 9 đến 30 để bảng tổng hợp gọn gàng
    for r in range(9, 31):
        for c in range(1, 34):
            col_letter = get_column_letter(c)
            ws_th[f"{col_letter}{r}"].value = None
            
    ws_th['B9'] = "(Số liệu chi tiết 92 xã được tổng hợp tại sheet CHI_TIET_THEO_XA)"
    ws_th['B9'].font = Font(name='Times New Roman', size=10, italic=True, color='555555')
    
    # 3. Điền Bảng tiêu chí (Hàng 33-35)
    crit = result['criteria']
    # Hàng 33: Đại học
    ws_th['I33'].value = crit['dai_hoc']['sl']
    ws_th['I33'].number_format = COUNT_FORMAT
    ws_th['I33'].font = font_bold
    ws_th['J33'].value = crit['dai_hoc']['tl']
    ws_th['J33'].number_format = PCT_FORMAT
    ws_th['J33'].font = font_bold
    
    # Hàng 34: Thạc sĩ trở lên
    ws_th['I34'].value = crit['thac_si']['sl']
    ws_th['I34'].number_format = COUNT_FORMAT
    ws_th['I34'].font = font_bold
    ws_th['J34'].value = crit['thac_si']['tl']
    ws_th['J34'].number_format = PCT_FORMAT
    ws_th['J34'].font = font_bold
    
    # Hàng 35: Chuẩn nghề nghiệp đạt trở lên
    ws_th['I35'].value = crit['chuan_nghe_nghiep']['sl']
    ws_th['I35'].number_format = COUNT_FORMAT
    ws_th['I35'].font = font_bold
    ws_th['J35'].value = crit['chuan_nghe_nghiep']['tl']
    ws_th['J35'].number_format = PCT_FORMAT
    ws_th['J35'].font = font_bold
    
    # 4. Sheet 2: CHI_TIET_THEO_XA
    ws_dt = wb.create_sheet(title='CHI_TIET_THEO_XA')
    headers_dt = [
        "STT", "Tên Xã / Phường", "Tên file nguồn",
        "Hạng 1", "Hạng 2", "Hạng 3", "2 buổi/ngày",
        "Hiệu trưởng", "P.Hiệu trưởng", "Tổng số GV", "Biên chế", "Hợp đồng", "Nữ", "Dân tộc", "Tỉ lệ GV/Lớp",
        "Trên ĐH", "ĐH", "CĐ", "THSP",
        "Tiểu học", "Âm nhạc", "Mỹ thuật", "Thể dục", "Tin học", "Ngoại ngữ", "Khác",
        "Tốt", "Khá", "Đạt", "Chưa đạt", "TPT.Đội", "Văn phòng", "T.Viện-TBDH"
    ]
    ws_dt.append(headers_dt)
    
    # Header styling
    fill_hdr = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    font_hdr = Font(name='Times New Roman', size=10, bold=True, color="FFFFFF")
    for col_idx in range(1, len(headers_dt) + 1):
        c = ws_dt.cell(row=1, column=col_idx)
        c.fill = fill_hdr
        c.font = font_hdr
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
    cols_order = [
        'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N',
        'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y',
        'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF'
    ]
    
    for row_idx, c_data in enumerate(result['commune_details'], start=2):
        row_vals = [
            c_data.get('stt', row_idx - 1),
            c_data.get('commune_name', ''),
            c_data.get('filename', '')
        ]
        for col in cols_order:
            row_vals.append(c_data.get(col, 0.0))
        ws_dt.append(row_vals)
        
        # Style dòng dữ liệu
        for c_idx in range(4, len(headers_dt) + 1):
            cell = ws_dt.cell(row=row_idx, column=c_idx)
            cell.font = font_data
            if headers_dt[c_idx - 1] == "Tỉ lệ GV/Lớp":
                cell.number_format = RATIO_FORMAT
            else:
                cell.number_format = COUNT_FORMAT
                
    # Dòng tổng cộng tỉnh ở cuối Sheet 2
    sum_row_idx = len(result['commune_details']) + 2
    sum_row_vals = ["", "TỔNG CỘNG TOÀN TỈNH", ""]
    for col in cols_order:
        sum_row_vals.append(agg.get(f"{col}8", 0.0))
    ws_dt.append(sum_row_vals)
    
    for c_idx in range(1, len(headers_dt) + 1):
        cell = ws_dt.cell(row=sum_row_idx, column=c_idx)
        cell.font = font_bold
        if c_idx >= 4:
            if headers_dt[c_idx - 1] == "Tỉ lệ GV/Lớp":
                cell.number_format = RATIO_FORMAT
            else:
                cell.number_format = COUNT_FORMAT
                
    # 5. Sheet 3: DOI_CHIEU_I_NGUON (Đối chiếu sai lệch Tổng số GV nguồn)
    if result.get('discrepancies'):
        ws_dc = wb.create_sheet(title='DOI_CHIEU_I_NGUON')
        headers_dc = [
            "STT", "Tên Xã / Phường", "Tên file nguồn",
            "Tổng GV nguồn (I8)", "Biên chế (J8)", "Hợp đồng (K8)",
            "Tổng GV tính chuẩn (J+K)", "Chênh lệch (Chuẩn - Nguồn)", "Ghi chú phân tích"
        ]
        ws_dc.append(headers_dc)
        fill_dc = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
        for col_idx in range(1, len(headers_dc) + 1):
            c = ws_dc.cell(row=1, column=col_idx)
            c.fill = fill_dc
            c.font = font_hdr
            c.alignment = Alignment(horizontal="center", vertical="center")
            
        for d_idx, d in enumerate(result['discrepancies'], start=2):
            note = "Khớp"
            if d['commune_name'] == "Phường Bắc Kạn":
                note = "Trường TH Bắc Kạn (dòng 9) có 72 GV biên chế nhưng ô I9 để 0, làm I8 thiếu 72 GV"
            else:
                note = f"Lệch {d['diff']:+.0f} GV do nguồn xã nhập đè số tĩnh hoặc công thức sum thiếu dòng"
            row_dc = [
                d['stt'], d['commune_name'], d['filename'],
                d['i_source'], d['j_val'], d['k_val'],
                d['i_recalc'], d['diff'], note
            ]
            ws_dc.append(row_dc)
            for c_idx in range(4, 9):
                ws_dc.cell(row=d_idx, column=c_idx).number_format = COUNT_FORMAT
                ws_dc.cell(row=d_idx, column=c_idx).font = font_data
        ws_dc.views.sheetView[0].showGridLines = True

    # Auto-fit columns
    for ws in [ws_th, ws_dt]:
        ws.views.sheetView[0].showGridLines = True
        
    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath

def export_1_m3_test(result, template_path, output_dir):
    """
    Xuất kết quả tổng hợp Biểu 1.M3 (Cơ sở vật chất cấp Tiểu học):
    - Sheet 1: TONG_HOP_THU (Bảng tổng hợp toàn tỉnh)
    - Sheet 2: CHI_TIET_THEO_XA (Số liệu tổng hợp cấp xã của 92 xã)
    - Sheet 3: DOI_CHIEU_D_NGUON (Đối soát xã có dòng cộng trùng lặp)
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_1.M3_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '1.M3':
            del wb[s_name]
            
    ws_th = wb['1.M3']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes']
    ws_th['A3'] = f"GHI CHÚ: TỔNG HỢP TOÀN TỈNH (Số xã thực tế được tính: {total_c} xã/phường)"
    ws_th['A3'].font = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    font_data = Font(name='Times New Roman', size=10)
    COUNT_FORMAT = '#,##0'
    AREA_FORMAT = '#,##0.0'
    PCT_FORMAT = '0.00"%"'
    
    # 1. Điền Hàng 8 (Dòng tổng tỉnh)
    agg = result['aggregated_cells']
    ws_th['B8'] = "TỔNG CỘNG TOÀN TỈNH"
    ws_th['B8'].font = font_bold
    
    # Danh sách các cột diện tích
    AREA_COLS = ['Q', 'S', 'U', 'W', 'Y', 'AA', 'AC']
    
    for c in range(3, 30):
        col = get_column_letter(c)
        cell = ws_th[f"{col}8"]
        val = agg.get(f"{col}8", 0.0)
        cell.value = val
        cell.font = font_bold
        if col == 'J':
            cell.number_format = PCT_FORMAT
        elif col in AREA_COLS:
            cell.number_format = AREA_FORMAT
        else:
            cell.number_format = COUNT_FORMAT
            
    # 2. Xóa các dòng mẫu trắng từ dòng 9 đến 35
    for r in range(9, 36):
        for c in range(1, 30):
            col_letter = get_column_letter(c)
            ws_th[f"{col_letter}{r}"].value = None
            
    ws_th['B9'] = "(Số liệu chi tiết 92 xã được tổng hợp tại sheet CHI_TIET_THEO_XA)"
    ws_th['B9'].font = Font(name='Times New Roman', size=10, italic=True, color='555555')
    
    # 3. Sheet 2: CHI_TIET_THEO_XA
    ws_dt = wb.create_sheet(title='CHI_TIET_THEO_XA')
    headers_dt = [
        "STT", "Tên Xã / Phường", "Tên file nguồn",
        "Số Đ.Tr", "Tổng số lớp", "Lớp ghép",
        "Kiên cố", "Bán kiên cố", "Tạm", "Thuê/mượn", "Tỉ lệ Ph/Lớp (%)",
        "P.HTr", "P.PHT", "P.VP", "P.Y tế", "P.Đội",
        "P.họp (SL)", "P.họp (DT)", "T.Viện (SL)", "T.Viện (DT)",
        "T.bị (SL)", "T.bị (DT)", "P.GV", "VS GV (DT)",
        "VS HS (SL)", "VS HS (DT)", "Sân chơi (SL)", "Sân chơi (DT)",
        "Sân tập (SL)", "Sân tập (DT)"
    ]
    ws_dt.append(headers_dt)
    
    fill_hdr = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    font_hdr = Font(name='Times New Roman', size=10, bold=True, color="FFFFFF")
    for col_idx in range(1, len(headers_dt) + 1):
        c = ws_dt.cell(row=1, column=col_idx)
        c.fill = fill_hdr
        c.font = font_hdr
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
    cols_order = [get_column_letter(c) for c in range(3, 30)]
    
    for row_idx, c_data in enumerate(result['commune_details'], start=2):
        row_vals = [
            c_data.get('stt', row_idx - 1),
            c_data.get('commune_name', ''),
            c_data.get('filename', '')
        ]
        for col in cols_order:
            row_vals.append(c_data.get(col, 0.0))
        ws_dt.append(row_vals)
        
        for c_idx in range(4, len(headers_dt) + 1):
            cell = ws_dt.cell(row=row_idx, column=c_idx)
            cell.font = font_data
            col_letter = cols_order[c_idx - 4]
            if col_letter == 'J':
                cell.number_format = PCT_FORMAT
            elif col_letter in AREA_COLS:
                cell.number_format = AREA_FORMAT
            else:
                cell.number_format = COUNT_FORMAT
                
    # Dòng tổng cộng tỉnh ở cuối Sheet 2
    sum_row_idx = len(result['commune_details']) + 2
    sum_row_vals = ["", "TỔNG CỘNG TOÀN TỈNH", ""]
    for col in cols_order:
        sum_row_vals.append(agg.get(f"{col}8", 0.0))
    ws_dt.append(sum_row_vals)
    
    for c_idx in range(1, len(headers_dt) + 1):
        cell = ws_dt.cell(row=sum_row_idx, column=c_idx)
        cell.font = font_bold
        if c_idx >= 4:
            col_letter = cols_order[c_idx - 4]
            if col_letter == 'J':
                cell.number_format = PCT_FORMAT
            elif col_letter in AREA_COLS:
                cell.number_format = AREA_FORMAT
            else:
                cell.number_format = COUNT_FORMAT
                
    # 4. Sheet 3: DOI_CHIEU_D_NGUON (Đối soát chênh lệch nguồn do cộng trùng)
    if result.get('discrepancies'):
        ws_dc = wb.create_sheet(title='DOI_CHIEU_D_NGUON')
        headers_dc = [
            "STT", "Tên Xã / Phường", "Tên file nguồn",
            "Số lớp nguồn (D8)", "Số lớp chuẩn hóa", "Chênh lệch", "Ghi chú phân tích nguyên nhân"
        ]
        ws_dc.append(headers_dc)
        fill_dc = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
        for col_idx in range(1, len(headers_dc) + 1):
            c = ws_dc.cell(row=1, column=col_idx)
            c.fill = fill_dc
            c.font = font_hdr
            c.alignment = Alignment(horizontal="center", vertical="center")
            
        for d_idx, d in enumerate(result['discrepancies'], start=2):
            diff = d['d8_raw'] - d['d_clean']
            row_dc = [
                d['stt'], d['commune_name'], d['filename'],
                d['d8_raw'], d['d_clean'], diff, d['note']
            ]
            ws_dc.append(row_dc)
            for c_idx in range(4, 7):
                ws_dc.cell(row=d_idx, column=c_idx).number_format = COUNT_FORMAT
                ws_dc.cell(row=d_idx, column=c_idx).font = font_data
        ws_dc.views.sheetView[0].showGridLines = True
        
    for ws in [ws_th, ws_dt]:
        ws.views.sheetView[0].showGridLines = True
        
    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath

def export_1_m4_test(result, template_path, output_dir):
    """
    Xuất kết quả tổng hợp Biểu 1.M4 (Kết quả PCGD Tiểu học toàn tỉnh):
    - Sử dụng đúng cấu trúc ma trận 116 hàng của Sở GDĐT:
      + Liệt kê chi tiết 92 xã/phường từ dòng 8 đến 99
      + Hàng tổng cộng toàn tỉnh ở dòng 100
      + Định dạng tỷ lệ 0.00"%" và số lượng #,##0
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_1.M4_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '1.M4':
            del wb[s_name]
            
    ws = wb['1.M4']
    ws.title = 'TONG_HOP_1.M4'
    
    total_c = result['total_communes']
    ws['A1'] = "BÁO CÁO KẾT QUẢ PHỔ CẬP GIÁO DỤC TIỂU HỌC NĂM 2026"
    ws['A2'] = f"GHI CHÚ: TỔNG HỢP TOÀN TỈNH (Số xã thực tế được tính: {total_c} xã/phường)"
    ws['A2'].font = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    font_data = Font(name='Times New Roman', size=10)
    COUNT_FORMAT = '#,##0'
    PCT_FORMAT = '0.00"%"'
    
    COLS_INT = ['C', 'D', 'E', 'F', 'H', 'J', 'L', 'N', 'O', 'P']
    COLS_PCT = ['G', 'I', 'K', 'M', 'Q']
    
    # 1. Điền 92 xã từ dòng 8 trở đi
    current_row = 8
    for item in result['commune_details']:
        ws[f"A{current_row}"] = item.get('stt', current_row - 7)
        ws[f"B{current_row}"] = item.get('commune_name', '')
        
        for c in COLS_INT:
            cell = ws[f"{c}{current_row}"]
            cell.value = item.get(c, 0.0)
            cell.font = font_data
            cell.number_format = COUNT_FORMAT
            
        for c in COLS_PCT:
            cell = ws[f"{c}{current_row}"]
            cell.value = item.get(c, 0.0)
            cell.font = font_data
            cell.number_format = PCT_FORMAT
            
        ws[f"R{current_row}"] = item.get('R', 'Đảm bảo')
        ws[f"S{current_row}"] = item.get('S', 'Đảm bảo')
        ws[f"T{current_row}"] = item.get('T', 'Mức độ 3')
        ws[f"U{current_row}"] = item.get('U', '')
        
        current_row += 1
        
    # 2. Điền Hàng TỔNG CỘNG TOÀN TỈNH
    sum_row = current_row
    ws[f"A{sum_row}"] = "@"
    ws[f"B{sum_row}"] = "TỔNG CỘNG TOÀN TỈNH"
    ws[f"B{sum_row}"].font = font_bold
    
    agg = result['aggregated_cells']
    
    for c in COLS_INT:
        cell = ws[f"{c}{sum_row}"]
        cell.value = agg.get(f"{c}8", 0.0)
        cell.font = font_bold
        cell.number_format = COUNT_FORMAT
        
    for c in COLS_PCT:
        cell = ws[f"{c}{sum_row}"]
        cell.value = agg.get(f"{c}8", 0.0)
        cell.font = font_bold
        cell.number_format = PCT_FORMAT
        
    ws[f"R{sum_row}"] = agg.get('R8', '')
    ws[f"S{sum_row}"] = agg.get('S8', '')
    ws[f"T{sum_row}"] = agg.get('T8', '')
    ws[f"U{sum_row}"] = ""
    for c in ['R', 'S', 'T', 'U']:
        ws[f"{c}{sum_row}"].font = font_bold
        
    # 3. Xóa các dòng mẫu thừa phía dưới
    for r in range(sum_row + 1, sum_row + 50):
        for c in range(1, 22):
            col_l = get_column_letter(c)
            ws[f"{col_l}{r}"].value = None
            
    ws.views.sheetView[0].showGridLines = True
    
    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_1_m5_test(result, template_path, output_dir):
    """
    Xuất kết quả thử nghiệm biểu 1.M5 (Thống kê đối tượng khuyết tật cấp Tiểu học):
    - Sheet 1: TONG_HOP_THU (Mẫu 1.M5 cấp tỉnh)
    - Sheet 2: CHI_TIET_THEO_XA (Ma trận chi tiết từng xã theo từng độ tuổi)
    - Sheet 3: DOI_CHIEU_H17_NGUON (Đối chiếu hàng 17 nguồn của các xã so với chi tiết)
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_1.M5_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '1.M5':
            del wb[s_name]
            
    ws_th = wb['1.M5']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes_count']
    commune_names_str = ", ".join(result['commune_names'])
    
    ws_th['E2'] = "TỔNG HỢP THỬ - THỐNG KÊ ĐỐI TƯỢNG KHUYẾT TẬT CẤP TIỂU HỌC"
    ws_th['E3'] = f"GHI CHÚ: TỔNG HỢP THỬ NGHIỆM (Số xã thực tế được tính: {total_c} xã - {commune_names_str})"
    font_sub = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    ws_th['E3'].font = font_sub
    
    agg = result['aggregated_cells']
    status_cells = result['status_cells']
    
    fill_incomplete = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    font_data = Font(name='Times New Roman', size=10)
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    
    RATE_FORMAT = '0.00"%"'
    COUNT_FORMAT = '#,##0'
    
    DISABILITY_COLS = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']
    
    # 1. Điền giá trị Hàng 8 đến 17
    for r in range(8, 18):
        for col in ['A', 'B', 'C', 'L', 'M'] + DISABILITY_COLS:
            coord = f"{col}{r}"
            if coord in agg:
                val = agg[coord]
                cell = ws_th[coord]
                cell.value = val
                cell.font = font_bold if r == 17 or col in ['C', 'L', 'M'] else font_data
                
                # Định dạng số hiển thị:
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

    ws_th.views.sheetView[0].showGridLines = True

    # ==========================================
    # SHEET 2: CHI TIẾT TỪNG XÃ ĐÓNG GÓP (1.M5)
    # ==========================================
    ws_detail = wb.create_sheet(title='CHI_TIET_THEO_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "BẢNG ĐỐI SOÁT CHI TIẾT SỐ LIỆU ĐỐI TƯỢNG KHUYẾT TẬT TỪNG XÃ - BIỂU 1.M5"
    ws_detail['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_detail['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Danh sách: {total_c} xã/phường"
    
    headers = ['STT', 'Xã/Phường', 'Hàng', 'Năm sinh', 'Độ tuổi', 'C(Tổng tật D:K)', 'Vận động', 'Nghe, nói', 'Nhìn', 'Thần kinh', 'Trí tuệ', 'Tự kỷ', 'Học tập', 'Khác', 'L(Tiếp cận GD)', 'M(Tỉ lệ %)']
    for ci, h in enumerate(headers, start=1):
        cell_h = ws_detail.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        
    cur_row = 5
    stt_idx = 1
    for c_name, matrix in result['commune_data'].items():
        for r in range(8, 18):
            ws_detail.cell(row=cur_row, column=1, value=stt_idx if r == 8 else "")
            ws_detail.cell(row=cur_row, column=2, value=c_name if r == 8 else "")
            ws_detail.cell(row=cur_row, column=3, value=r)
            ws_detail.cell(row=cur_row, column=4, value=matrix[f"A{r}"]['val'])
            ws_detail.cell(row=cur_row, column=5, value=matrix[f"B{r}"]['val'])
            
            # Tính lại C
            c_calc = sum(matrix[f"{col}{r}"]['val'] for col in DISABILITY_COLS if matrix[f"{col}{r}"]['type'] == 'NUMBER')
            c_cell = ws_detail.cell(row=cur_row, column=6, value=c_calc)
            c_cell.number_format = COUNT_FORMAT
            
            # D..K
            for ci, col in enumerate(DISABILITY_COLS, start=7):
                v = matrix[f"{col}{r}"]['val']
                cl = ws_detail.cell(row=cur_row, column=ci, value=v)
                if isinstance(v, (int, float)): cl.number_format = COUNT_FORMAT
                
            # L
            l_val = matrix[f"L{r}"]['val']
            cl_l = ws_detail.cell(row=cur_row, column=15, value=l_val)
            if isinstance(l_val, (int, float)): cl_l.number_format = COUNT_FORMAT
            
            # M
            m_val = round(l_val / c_calc * 100, 2) if (isinstance(l_val, (int, float)) and c_calc > 0) else 0.0
            cl_m = ws_detail.cell(row=cur_row, column=16, value=m_val)
            cl_m.number_format = RATE_FORMAT
            
            if r == 17:
                for c_idx in range(1, 17):
                    ws_detail.cell(row=cur_row, column=c_idx).font = font_bold
                    
            cur_row += 1
        stt_idx += 1

    # ==========================================
    # SHEET 3: ĐỐI CHIẾU HÀNG 17 NGUỒN CỦA CÁC XÃ
    # ==========================================
    ws_h17 = wb.create_sheet(title='DOI_CHIEU_H17_NGUON')
    ws_h17.views.sheetView[0].showGridLines = True
    
    ws_h17['A1'] = "BẢNG ĐỐI SOÁT HÀNG 17 (TỔNG CỘNG) NGUỒN CỦA CÁC XÃ SO VỚI TỔNG CHI TIẾT HÀNG 8-16"
    ws_h17['A1'].font = Font(name='Times New Roman', size=12, bold=True)
    ws_h17['A2'] = "Lưu ý: Hệ thống tính cấp tỉnh từ tổng chi tiết từng độ tuổi (H8-16) để tránh thất thoát dữ liệu khi xã gõ đè số 0 vào hàng 17."
    
    h17_headers = ['STT', 'Xã/Phường', 'Tổng tật chi tiết (C)', 'Tổng tật nguồn (C17)', 'Chênh lệch C', 'Tiếp cận GD chi tiết (L)', 'Tiếp cận GD nguồn (L17)', 'Chênh lệch L', 'Ghi chú giải trình']
    for ci, h in enumerate(h17_headers, start=1):
        cell_h = ws_h17.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        
    discrepancies = result.get('h17_discrepancies', [])
    if discrepancies:
        for idx, item in enumerate(discrepancies, start=1):
            r_idx = 4 + idx
            ws_h17.cell(row=r_idx, column=1, value=idx)
            ws_h17.cell(row=r_idx, column=2, value=item['commune'])
            
            c3 = ws_h17.cell(row=r_idx, column=3, value=item['detail_c'])
            c3.number_format = COUNT_FORMAT
            c4 = ws_h17.cell(row=r_idx, column=4, value=item['src_c17'])
            c4.number_format = COUNT_FORMAT
            c5 = ws_h17.cell(row=r_idx, column=5, value=item['diff_c'])
            c5.number_format = COUNT_FORMAT
            
            c6 = ws_h17.cell(row=r_idx, column=6, value=item['detail_l'])
            c6.number_format = COUNT_FORMAT
            c7 = ws_h17.cell(row=r_idx, column=7, value=item['src_l17'])
            c7.number_format = COUNT_FORMAT
            c8 = ws_h17.cell(row=r_idx, column=8, value=item['diff_l'])
            c8.number_format = COUNT_FORMAT
            
            ws_h17.cell(row=r_idx, column=9, value=item['reason'])
    else:
        ws_h17.cell(row=5, column=1, value=1)
        ws_h17.cell(row=5, column=2, value="Tất cả các xã đều có hàng 17 khớp hoàn toàn với tổng các dòng chi tiết hàng 8-16.")

    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_1_m6_test(result, template_path, output_dir):
    """
    Xuất kết quả thử nghiệm biểu 1.M6 (Học sinh học ngoài xã, ngoài tỉnh cấp Tiểu học):
    - Sheet 1: TONG_HOP_THU (Mẫu 1.M6 cấp tỉnh)
    - Sheet 2: CHI_TIET_THEO_XA (Thống kê số lượng từng xã theo độ tuổi)
    - Sheet 3: DANH_SACH_TRUONG_NGOAI (Danh sách các trường học ngoài địa bàn được khai báo)
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_1.M6_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '1.M6':
            del wb[s_name]
            
    ws_th = wb['1.M6']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes_count']
    commune_names_str = ", ".join(result['commune_names'])
    
    ws_th['F2'] = "TỔNG HỢP THỬ - HỌC SINH ĐANG HỌC TẠI XÃ/PHƯỜNG NGOÀI VÀ TỈNH NGOÀI CẤP TIỂU HỌC"
    ws_th['F3'] = f"GHI CHÚ: TỔNG HỢP THỬ NGHIỆM (Số xã thực tế được tính: {total_c} xã - {commune_names_str})"
    font_sub = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    ws_th['F3'].font = font_sub
    
    agg = result['aggregated_cells']
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    COUNT_FORMAT = '#,##0'
    
    # 1. Điền giá trị Hàng 27 (Tổng Phần I), Hàng 48 (Tổng Phần II), Hàng 49 (Tổng I + II)
    all_cols = [get_column_letter(c) for c in range(3, 44)] + ['AR'] # C..AQ và AR
    for r in [27, 48, 49]:
        for col in all_cols:
            coord = f"{col}{r}"
            if coord in agg:
                val = agg[coord]
                cell = ws_th[coord]
                cell.value = val
                cell.font = font_bold
                if isinstance(val, (int, float)):
                    cell.number_format = COUNT_FORMAT

    ws_th.views.sheetView[0].showGridLines = True

    # ==========================================
    # SHEET 2: CHI TIẾT THEO TỪNG XÃ (1.M6)
    # ==========================================
    ws_detail = wb.create_sheet(title='CHI_TIET_THEO_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "BẢNG TỔNG HỢP SỐ LIỆU HỌC SINH HỌC NGOÀI ĐỊA BÀN TỪNG XÃ - BIỂU 1.M6"
    ws_detail['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_detail['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Quy mô: {total_c} xã/phường"
    
    headers_s2 = ['STT', 'Tên Xã / Phường', 'Tên File Nguồn', 'Học ngoài xã (trong tỉnh)', 'Học ngoài tỉnh', 'Tổng cộng học sinh đi học ngoài', '6 Tuổi', '7 Tuổi', '8 Tuổi', '9 Tuổi', '10 Tuổi', '11 Tuổi', '12 Tuổi', '13 Tuổi', '14 Tuổi']
    for ci, h in enumerate(headers_s2, start=1):
        cell_h = ws_detail.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        
    cur_row = 5
    for item in result['commune_summaries']:
        ws_detail.cell(row=cur_row, column=1, value=item['stt'])
        ws_detail.cell(row=cur_row, column=2, value=item['commune_name'])
        ws_detail.cell(row=cur_row, column=3, value=item['filename'])
        
        c4 = ws_detail.cell(row=cur_row, column=4, value=item['part1_sum'])
        c4.number_format = COUNT_FORMAT
        c5 = ws_detail.cell(row=cur_row, column=5, value=item['part2_sum'])
        c5.number_format = COUNT_FORMAT
        c6 = ws_detail.cell(row=cur_row, column=6, value=item['total_sum'])
        c6.number_format = COUNT_FORMAT
        c6.font = font_bold
        
        # Nhóm độ tuổi
        ages = ['6 Tuổi', '7 Tuổi', '8 Tuổi', '9 Tuổi', '10 Tuổi', '11 Tuổi', '12 Tuổi', '13 Tuổi', '14 Tuổi']
        for ai, age_k in enumerate(ages, start=7):
            age_v = item['tot_by_age'].get(age_k, 0.0)
            cell_age = ws_detail.cell(row=cur_row, column=ai, value=age_v)
            cell_age.number_format = COUNT_FORMAT
            
        cur_row += 1
        
    # Hàng tổng cộng toàn tỉnh ở Sheet 2
    ws_detail.cell(row=cur_row, column=2, value="TỔNG CỘNG TOÀN TỈNH").font = font_bold
    c_p1 = ws_detail.cell(row=cur_row, column=4, value=result['total_part1'])
    c_p1.number_format = COUNT_FORMAT
    c_p1.font = font_bold
    c_p2 = ws_detail.cell(row=cur_row, column=5, value=result['total_part2'])
    c_p2.number_format = COUNT_FORMAT
    c_p2.font = font_bold
    c_gt = ws_detail.cell(row=cur_row, column=6, value=result['total_grand'])
    c_gt.number_format = COUNT_FORMAT
    c_gt.font = font_bold
    
    ages = ['6 Tuổi', '7 Tuổi', '8 Tuổi', '9 Tuổi', '10 Tuổi', '11 Tuổi', '12 Tuổi', '13 Tuổi', '14 Tuổi']
    for ai, age_k in enumerate(ages, start=7):
        age_sum = result['province_by_age'][age_k]['total']
        c_age_sum = ws_detail.cell(row=cur_row, column=ai, value=age_sum)
        c_age_sum.number_format = COUNT_FORMAT
        c_age_sum.font = font_bold

    # ==========================================
    # SHEET 3: DANH SÁCH TRƯỜNG HỌC KHAI BÁO
    # ==========================================
    ws_schools = wb.create_sheet(title='DANH_SACH_TRUONG_NGOAI')
    ws_schools.views.sheetView[0].showGridLines = True
    
    ws_schools['A1'] = "DANH SÁCH CÁC TRƯỜNG HỌC NGOÀI ĐỊA BÀN ĐƯỢC 92 XÃ/PHƯỜNG KHAI BÁO - BIỂU 1.M6"
    ws_schools['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_schools['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Tổng số trường/cơ sở ghi nhận: {len(result['school_records'])} lượt khai báo"
    
    headers_s3 = ['STT', 'Xã / Phường khai báo', 'Phân loại', 'Tên trường học ngoài địa bàn', 'Tổng số học sinh', '6 Tuổi', '7 Tuổi', '8 Tuổi', '9 Tuổi', '10 Tuổi', '11 Tuổi', '12 Tuổi', '13 Tuổi', '14 Tuổi']
    for ci, h in enumerate(headers_s3, start=1):
        cell_h = ws_schools.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        
    cur_row_s3 = 5
    for idx, rec in enumerate(result['school_records'], start=1):
        ws_schools.cell(row=cur_row_s3, column=1, value=idx)
        ws_schools.cell(row=cur_row_s3, column=2, value=rec['commune'])
        ws_schools.cell(row=cur_row_s3, column=3, value=rec['part'])
        ws_schools.cell(row=cur_row_s3, column=4, value=rec['school_name'])
        
        c_tot = ws_schools.cell(row=cur_row_s3, column=5, value=rec['total_students'])
        c_tot.number_format = COUNT_FORMAT
        c_tot.font = font_bold
        
        for ai, age_k in enumerate(ages, start=6):
            age_v = rec['by_age'].get(age_k, 0.0)
            cell_age = ws_schools.cell(row=cur_row_s3, column=ai, value=age_v)
            cell_age.number_format = COUNT_FORMAT
            
        cur_row_s3 += 1

    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_1_m7_test(result, template_path="MẪU CHUẨN - PCGD 2026(ok).xlsx", output_dir="outputs"):
    """
    Xuất file Excel tổng hợp thử cho Biểu 1.M7:
    - Sheet 1: TONG_HOP_THU (Cấp tỉnh - Template 1.M7)
    - Sheet 2: CHI_TIET_THEO_XA (Thống kê số lượng từng xã theo độ tuổi)
    - Sheet 3: DANH_SACH_TRUONG (Danh sách các trường học được khai báo)
    - Sheet 4: DOI_CHIEU_SAI_LECH (Bảng ghi nhận sai lệch hàng tổng nguồn)
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_1.M7_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '1.M7':
            del wb[s_name]
            
    ws_th = wb['1.M7']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes_count']
    commune_names_str = ", ".join(result['commune_names'])
    
    ws_th['F2'] = "TỔNG HỢP THỬ - HỌC SINH NGOÀI ĐỊA BÀN ĐANG HỌC TẠI CÁC TRƯỜNG TRÊN ĐỊA BÀN XÃ CẤP TIỂU HỌC"
    ws_th['F3'] = f"GHI CHÚ: TỔNG HỢP THỬ NGHIỆM (Số xã thực tế được tính: {total_c} xã - {commune_names_str})"
    font_sub = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    ws_th['F3'].font = font_sub
    
    agg = result['aggregated_cells']
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    COUNT_FORMAT = '#,##0'
    
    # 1. Điền giá trị Hàng 7 (Tổng I + II), Hàng 34 (Tổng Phần I), Hàng 68 (Tổng Phần II)
    all_cols = [get_column_letter(c) for c in range(3, 44)] + ['AR'] # C..AQ và AR
    for r in [7, 34, 68]:
        for col in all_cols:
            coord = f"{col}{r}"
            if coord in agg:
                val = agg[coord]
                cell = ws_th[coord]
                cell.value = val
                cell.font = font_bold
                if isinstance(val, (int, float)):
                    cell.number_format = COUNT_FORMAT

    ws_th.views.sheetView[0].showGridLines = True

    # ==========================================
    # SHEET 2: CHI TIẾT THEO TỪNG XÃ (1.M7)
    # ==========================================
    ws_detail = wb.create_sheet(title='CHI_TIET_THEO_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "BẢNG TỔNG HỢP SỐ LIỆU HỌC SINH NGOÀI ĐỊA BÀN HỌC TẠI XÃ TỪNG XÃ - BIỂU 1.M7"
    ws_detail['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_detail['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Quy mô: {total_c} xã/phường"
    
    headers_s2 = ['STT', 'Tên Xã / Phường', 'Tên File Nguồn', 'Học ngoài xã (trong tỉnh) đến học', 'Học ngoài tỉnh đến học', 'Tổng cộng ngoại địa bàn', '6 Tuổi', '7 Tuổi', '8 Tuổi', '9 Tuổi', '10 Tuổi', '11 Tuổi', '12 Tuổi', '13 Tuổi', '14 Tuổi']
    for ci, h in enumerate(headers_s2, start=1):
        cell_h = ws_detail.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        
    cur_row = 5
    for item in result['commune_summaries']:
        ws_detail.cell(row=cur_row, column=1, value=item['stt'])
        ws_detail.cell(row=cur_row, column=2, value=item['commune_name'])
        ws_detail.cell(row=cur_row, column=3, value=item['filename'])
        
        c4 = ws_detail.cell(row=cur_row, column=4, value=item['part1_sum'])
        c4.number_format = COUNT_FORMAT
        c5 = ws_detail.cell(row=cur_row, column=5, value=item['part2_sum'])
        c5.number_format = COUNT_FORMAT
        c6 = ws_detail.cell(row=cur_row, column=6, value=item['total_sum'])
        c6.number_format = COUNT_FORMAT
        c6.font = font_bold
        
        # Nhóm độ tuổi
        ages = ['6 Tuổi', '7 Tuổi', '8 Tuổi', '9 Tuổi', '10 Tuổi', '11 Tuổi', '12 Tuổi', '13 Tuổi', '14 Tuổi']
        for ai, age_k in enumerate(ages, start=7):
            age_v = item['tot_by_age'].get(age_k, 0.0)
            cell_age = ws_detail.cell(row=cur_row, column=ai, value=age_v)
            cell_age.number_format = COUNT_FORMAT
            
        cur_row += 1
        
    # Hàng tổng cộng toàn tỉnh ở Sheet 2
    ws_detail.cell(row=cur_row, column=2, value="TỔNG CỘNG TOÀN TỈNH").font = font_bold
    c_p1 = ws_detail.cell(row=cur_row, column=4, value=result['total_part1'])
    c_p1.number_format = COUNT_FORMAT
    c_p1.font = font_bold
    c_p2 = ws_detail.cell(row=cur_row, column=5, value=result['total_part2'])
    c_p2.number_format = COUNT_FORMAT
    c_p2.font = font_bold
    c_gt = ws_detail.cell(row=cur_row, column=6, value=result['total_grand'])
    c_gt.number_format = COUNT_FORMAT
    c_gt.font = font_bold
    
    ages = ['6 Tuổi', '7 Tuổi', '8 Tuổi', '9 Tuổi', '10 Tuổi', '11 Tuổi', '12 Tuổi', '13 Tuổi', '14 Tuổi']
    for ai, age_k in enumerate(ages, start=7):
        age_sum = result['province_by_age'][age_k]['total']
        c_age_sum = ws_detail.cell(row=cur_row, column=ai, value=age_sum)
        c_age_sum.number_format = COUNT_FORMAT
        c_age_sum.font = font_bold

    # ==========================================
    # SHEET 3: DANH SÁCH TRƯỜNG HỌC KHAI BÁO
    # ==========================================
    ws_schools = wb.create_sheet(title='DANH_SACH_TRUONG')
    ws_schools.views.sheetView[0].showGridLines = True
    
    ws_schools['A1'] = "DANH SÁCH CÁC TRƯỜNG HỌC NGOÀI ĐỊA BÀN ĐƯỢC 92 XÃ/PHƯỜNG KHAI BÁO - BIỂU 1.M7"
    ws_schools['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_schools['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Tổng số cơ sở ghi nhận: {len(result['school_records'])} lượt khai báo"
    
    headers_s3 = ['STT', 'Xã / Phường khai báo', 'Phân loại', 'Tên trường học ngoài địa bàn', 'Tổng số học sinh', '6 Tuổi', '7 Tuổi', '8 Tuổi', '9 Tuổi', '10 Tuổi', '11 Tuổi', '12 Tuổi', '13 Tuổi', '14 Tuổi']
    for ci, h in enumerate(headers_s3, start=1):
        cell_h = ws_schools.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        
    cur_row_s3 = 5
    for idx, rec in enumerate(result['school_records'], start=1):
        ws_schools.cell(row=cur_row_s3, column=1, value=idx)
        ws_schools.cell(row=cur_row_s3, column=2, value=rec['commune'])
        ws_schools.cell(row=cur_row_s3, column=3, value=rec['part'])
        ws_schools.cell(row=cur_row_s3, column=4, value=rec['school_name'])
        
        c_tot = ws_schools.cell(row=cur_row_s3, column=5, value=rec['total_students'])
        c_tot.number_format = COUNT_FORMAT
        c_tot.font = font_bold
        
        for ai, age_k in enumerate(ages, start=6):
            age_v = rec['by_age'].get(age_k, 0.0)
            cell_age = ws_schools.cell(row=cur_row_s3, column=ai, value=age_v)
            cell_age.number_format = COUNT_FORMAT
            
        cur_row_s3 += 1

    # ==========================================
    # SHEET 4: ĐỐI CHIẾU SAI LỆCH NGUỒN
    # ==========================================
    ws_disc = wb.create_sheet(title='DOI_CHIEU_SAI_LECH')
    ws_disc.views.sheetView[0].showGridLines = True
    
    ws_disc['A1'] = "BẢNG ĐỐI CHIẾU SAI LỆCH HÀNG TỔNG NGUỒN XÃ ĐÃ ĐƯỢC CLEAN RECALCULATION XỬ LÝ - BIỂU 1.M7"
    ws_disc['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_disc['A2'] = "Ghi chú: So sánh giữa tổng tính lại từ chi tiết Phần I + Phần II và số liệu có sẵn trong hàng 7 nguồn của xã"
    
    headers_s4 = ['STT', 'Tên Xã / Phường', 'Hàng tổng nguồn', 'Tổng AR nguồn', 'Tổng C..AQ nguồn', 'Tổng tính chuẩn (P1+P2)', 'Chênh lệch', 'Chi tiết các ô sai lệch']
    for ci, h in enumerate(headers_s4, start=1):
        cell_h = ws_disc.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        
    cur_row_s4 = 5
    for idx, disc in enumerate(result['discrepancies'], start=1):
        ws_disc.cell(row=cur_row_s4, column=1, value=idx)
        ws_disc.cell(row=cur_row_s4, column=2, value=disc['commune'])
        ws_disc.cell(row=cur_row_s4, column=3, value=disc['row_tot'])
        
        c4 = ws_disc.cell(row=cur_row_s4, column=4, value=disc['src_tot_ar'])
        c4.number_format = COUNT_FORMAT
        c5 = ws_disc.cell(row=cur_row_s4, column=5, value=disc['src_caq_sum'])
        c5.number_format = COUNT_FORMAT
        c6 = ws_disc.cell(row=cur_row_s4, column=6, value=disc['calc_sum'])
        c6.number_format = COUNT_FORMAT
        c6.font = font_bold
        c7 = ws_disc.cell(row=cur_row_s4, column=7, value=disc['diff'])
        c7.number_format = COUNT_FORMAT
        c7.font = font_bold
        
        ws_disc.cell(row=cur_row_s4, column=8, value=disc['diff_cols'])
        cur_row_s4 += 1

    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_2_m2_test(result, template_path="MẪU CHUẨN - PCGD 2026(ok).xlsx", output_dir="outputs"):
    """
    Xuất file Excel tổng hợp thử cho Biểu 2.M2 (Huy động PCGD THCS):
    - Sheet 1: TONG_HOP_THU (Cấp tỉnh - Template 2.M2 với hàng 48 toàn tỉnh)
    - Sheet 2: CHI_TIET_92_XA (Bảng chi tiết 92 xã + dòng Tổng toàn tỉnh)
    - Sheet 3: DOI_CHIEU_SAI_LECH (Bảng ghi nhận đối chiếu sai lệch nguồn)
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_2.M2_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '2.M2':
            del wb[s_name]
            
    ws_th = wb['2.M2']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes_count']
    commune_names_str = ", ".join(result['commune_names'])
    
    ws_th['C2'] = f"GHI CHÚ: TỔNG HỢP THỬ NGHIỆM (Số xã thực tế được tính: {total_c} xã - {commune_names_str})"
    font_sub = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    ws_th['C2'].font = font_sub
    
    agg = result['aggregated_cells']
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    COUNT_FORMAT = '#,##0'
    RATE_FORMAT = '0.00"%"'
    
    count_cols = ['C', 'D', 'F', 'G', 'H', 'I', 'K', 'L', 'N', 'O', 'P', 'R', 'S', 'T', 'U']
    rate_cols = ['E', 'J', 'M', 'Q', 'V']
    
    # 1. Điền giá trị Hàng 48 (Tổng toàn tỉnh)
    for c in count_cols:
        coord = f"{c}48"
        if coord in agg:
            cell = ws_th[coord]
            cell.value = agg[coord]
            cell.font = font_bold
            cell.number_format = COUNT_FORMAT
            
    for c in rate_cols:
        coord = f"{c}48"
        if coord in agg:
            cell = ws_th[coord]
            cell.value = agg[coord]
            cell.font = font_bold
            cell.number_format = RATE_FORMAT
            
    ws_th['W48'].value = agg.get('W48', 'Đạt Mức độ 3')
    ws_th['W48'].font = font_bold
    ws_th.views.sheetView[0].showGridLines = True

    # ==========================================
    # SHEET 2: CHI TIẾT 92 XÃ (2.M2)
    # ==========================================
    ws_detail = wb.create_sheet(title='CHI_TIET_92_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "BẢNG TỔNG HỢP CHI TIẾT KẾT QUẢ HUY ĐỘNG PCGD THCS TỪNG XÃ - BIỂU 2.M2"
    ws_detail['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_detail['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Quy mô: {total_c} xã/phường"
    
    headers_s2 = [
        'STT', 'Tên Xã / Phường', 
        '6t TS (C)', '6t vào L1 (D)', 'Tỷ lệ 6t (E)',
        'HS HTCTTH (F)', 'Vào L6 PT (G)', 'Vào L6 GDTX (H)', 'Cộng vào L6 (I)', 'Tỷ lệ vào L6 (J)',
        '11-14t TS (K)', '11-14t HTCTTH (L)', 'Tỷ lệ 11-14t (M)',
        'HS lớp 9 (N)', 'TN THCS PT (O)', 'TN THCS GDTX (P)', 'Tỷ lệ TN lớp 9 (Q)',
        '15-18t TS (R)', 'Có bằng THCS PT (S)', 'Có bằng GDTX (T)', 'Cộng có bằng (U)', 'Tỷ lệ có bằng (V)',
        'Đạt chuẩn (W)'
    ]
    for ci, h in enumerate(headers_s2, start=1):
        cell_h = ws_detail.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        cell_h.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
    cur_row = 5
    for item in result['commune_summaries']:
        ws_detail.cell(row=cur_row, column=1, value=item['stt'])
        ws_detail.cell(row=cur_row, column=2, value=item['commune_name'])
        
        counts = item['counts']
        rates = item['rates']
        
        # 6t
        c3 = ws_detail.cell(row=cur_row, column=3, value=counts['C'])
        c3.number_format = COUNT_FORMAT
        c4 = ws_detail.cell(row=cur_row, column=4, value=counts['D'])
        c4.number_format = COUNT_FORMAT
        c5 = ws_detail.cell(row=cur_row, column=5, value=round(rates['E'], 2))
        c5.number_format = RATE_FORMAT
        
        # L6
        c6 = ws_detail.cell(row=cur_row, column=6, value=counts['F'])
        c6.number_format = COUNT_FORMAT
        c7 = ws_detail.cell(row=cur_row, column=7, value=counts['G'])
        c7.number_format = COUNT_FORMAT
        c8 = ws_detail.cell(row=cur_row, column=8, value=counts['H'])
        c8.number_format = COUNT_FORMAT
        c9 = ws_detail.cell(row=cur_row, column=9, value=counts['I'])
        c9.number_format = COUNT_FORMAT
        c10 = ws_detail.cell(row=cur_row, column=10, value=round(rates['J'], 2))
        c10.number_format = RATE_FORMAT
        
        # 11-14t
        c11 = ws_detail.cell(row=cur_row, column=11, value=counts['K'])
        c11.number_format = COUNT_FORMAT
        c12 = ws_detail.cell(row=cur_row, column=12, value=counts['L'])
        c12.number_format = COUNT_FORMAT
        c13 = ws_detail.cell(row=cur_row, column=13, value=round(rates['M'], 2))
        c13.number_format = RATE_FORMAT
        
        # L9
        c14 = ws_detail.cell(row=cur_row, column=14, value=counts['N'])
        c14.number_format = COUNT_FORMAT
        c15 = ws_detail.cell(row=cur_row, column=15, value=counts['O'])
        c15.number_format = COUNT_FORMAT
        c16 = ws_detail.cell(row=cur_row, column=16, value=counts['P'])
        c16.number_format = COUNT_FORMAT
        c17 = ws_detail.cell(row=cur_row, column=17, value=round(rates['Q'], 2))
        c17.number_format = RATE_FORMAT
        
        # 15-18t
        c18 = ws_detail.cell(row=cur_row, column=18, value=counts['R'])
        c18.number_format = COUNT_FORMAT
        c19 = ws_detail.cell(row=cur_row, column=19, value=counts['S'])
        c19.number_format = COUNT_FORMAT
        c20 = ws_detail.cell(row=cur_row, column=20, value=counts['T'])
        c20.number_format = COUNT_FORMAT
        c21 = ws_detail.cell(row=cur_row, column=21, value=counts['U'])
        c21.number_format = COUNT_FORMAT
        c22 = ws_detail.cell(row=cur_row, column=22, value=round(rates['V'], 2))
        c22.number_format = RATE_FORMAT
        
        # Đạt chuẩn
        ws_detail.cell(row=cur_row, column=23, value=item['standard'])
        cur_row += 1
        
    # Hàng tổng cộng toàn tỉnh ở Sheet 2
    ws_detail.cell(row=cur_row, column=2, value="TỔNG CỘNG TOÀN TỈNH").font = font_bold
    p_cnt = result['province_counts']
    p_rt = result['province_rates']
    
    col_map = {
        3: ('C', COUNT_FORMAT), 4: ('D', COUNT_FORMAT), 5: ('E', RATE_FORMAT),
        6: ('F', COUNT_FORMAT), 7: ('G', COUNT_FORMAT), 8: ('H', COUNT_FORMAT),
        9: ('I', COUNT_FORMAT), 10: ('J', RATE_FORMAT),
        11: ('K', COUNT_FORMAT), 12: ('L', COUNT_FORMAT), 13: ('M', RATE_FORMAT),
        14: ('N', COUNT_FORMAT), 15: ('O', COUNT_FORMAT), 16: ('P', COUNT_FORMAT),
        17: ('Q', RATE_FORMAT),
        18: ('R', COUNT_FORMAT), 19: ('S', COUNT_FORMAT), 20: ('T', COUNT_FORMAT),
        21: ('U', COUNT_FORMAT), 22: ('V', RATE_FORMAT)
    }
    
    for ci, (col_k, fmt) in col_map.items():
        val = p_cnt[col_k] if col_k in p_cnt else round(p_rt[col_k], 2)
        cell = ws_detail.cell(row=cur_row, column=ci, value=val)
        cell.font = font_bold
        cell.number_format = fmt
        
    ws_detail.cell(row=cur_row, column=23, value="Đạt chuẩn Mức độ 3").font = font_bold

    # ==========================================
    # SHEET 3: ĐỐI CHIẾU SAI LỆCH NGUỒN
    # ==========================================
    ws_disc = wb.create_sheet(title='DOI_CHIEU_SAI_LECH')
    ws_disc.views.sheetView[0].showGridLines = True
    
    ws_disc['A1'] = "BẢNG ĐỐI CHIẾU SAI LỆCH NGUỒN ĐÃ ĐƯỢC CLEAN RECALCULATION XỬ LÝ - BIỂU 2.M2"
    ws_disc['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_disc['A2'] = "Ghi chú: Nhận diện lỗi nhập liệu gõ nhầm số thứ tự vào cột GDTX và xử lý chuẩn hóa số liệu cấp tỉnh"
    
    headers_s3 = ['STT', 'Tên Xã / Phường', 'Hàng dòng tổng', 'Chỉ tiêu sai lệch', 'Giá trị nguồn', 'Giá trị chuẩn hóa', 'Nguyên nhân và cách xử lý']
    for ci, h in enumerate(headers_s3, start=1):
        cell_h = ws_disc.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        
    cur_row_s3 = 5
    for idx, disc in enumerate(result['discrepancies'], start=1):
        ws_disc.cell(row=cur_row_s3, column=1, value=idx)
        ws_disc.cell(row=cur_row_s3, column=2, value=disc['commune'])
        ws_disc.cell(row=cur_row_s3, column=3, value=disc['tot_row'])
        ws_disc.cell(row=cur_row_s3, column=4, value=disc['field'])
        ws_disc.cell(row=cur_row_s3, column=5, value=disc['src_val'])
        ws_disc.cell(row=cur_row_s3, column=6, value=disc['clean_val'])
        ws_disc.cell(row=cur_row_s3, column=7, value=disc['reason'])
        cur_row_s3 += 1

    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_2_m3_test(result, template_path="MẪU CHUẨN - PCGD 2026(ok).xlsx", output_dir="outputs"):
    """
    Xuất file Excel tổng hợp thử cho Biểu 2.M3 (Thống kê PCGD THCS):
    - Sheet 1: TONG_HOP_THU (Cấp tỉnh - Template 2.M3 với hàng 46 toàn tỉnh)
    - Sheet 2: CHI_TIET_92_XA (Bảng chi tiết 92 xã + dòng Tổng toàn tỉnh)
    - Sheet 3: DOI_CHIEU_SAI_LECH (Bảng ghi nhận đối chiếu sai lệch nguồn)
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_2.M3_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '2.M3':
            del wb[s_name]
            
    ws_th = wb['2.M3']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes_count']
    commune_names_str = ", ".join(result['commune_names'])
    
    ws_th['F2'] = f"GHI CHÚ: TỔNG HỢP THỬ NGHIỆM (Số xã thực tế được tính: {total_c} xã - {commune_names_str})"
    font_sub = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    ws_th['F2'].font = font_sub
    
    agg = result['aggregated_cells']
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    COUNT_FORMAT = '#,##0'
    RATE_FORMAT = '0.00"%"'
    
    count_cols = ['C', 'D', 'E', 'H', 'I', 'K', 'M', 'N', 'O']
    rate_cols = ['J', 'L', 'P']
    text_cols = ['F', 'G', 'Q', 'R', 'S']
    
    # 1. Điền giá trị Hàng 46 (Tổng toàn tỉnh)
    for c in count_cols:
        coord = f"{c}46"
        if coord in agg:
            cell = ws_th[coord]
            cell.value = agg[coord]
            cell.font = font_bold
            cell.number_format = COUNT_FORMAT
            
    for c in rate_cols:
        coord = f"{c}46"
        if coord in agg:
            cell = ws_th[coord]
            cell.value = agg[coord]
            cell.font = font_bold
            cell.number_format = RATE_FORMAT
            
    for c in text_cols:
        coord = f"{c}46"
        if coord in agg:
            cell = ws_th[coord]
            cell.value = agg[coord]
            cell.font = font_bold
            
    ws_th.views.sheetView[0].showGridLines = True

    # ==========================================
    # SHEET 2: CHI TIẾT 92 XÃ (2.M3)
    # ==========================================
    ws_detail = wb.create_sheet(title='CHI_TIET_92_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "BẢNG TỔNG HỢP CHI TIẾT KẾT QUẢ PHỔ CẬP GIÁO DỤC THCS TỪNG XÃ - BIỂU 2.M3"
    ws_detail['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_detail['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Quy mô: {total_c} xã/phường"
    
    headers_s2 = [
        'STT', 'Tên Xã / Phường', 
        'Số trường THCS (C)', 'TTN 11-18t TS (D)', 'Khuyết tật (E)',
        'Chuẩn GDTH (F)', 'Chuẩn XMC (G)',
        'Tổng số PPC (H)', 'Tổng TNTHCS (I)', 'Tỷ lệ TN (%) (J)',
        'Học tiếp THPT/GDTX/GDNN (K)', 'Tỷ lệ học tiếp (%) (L)',
        'Trẻ KT Tổng số (M)', 'KT có khả năng HT (N)', 'KT tiếp cận GD (O)', 'Tỷ lệ tiếp cận (%) (P)',
        'Đội ngũ GV-NV (Q)', 'CSVC TBDH (R)', 'Đạt chuẩn PC THCS (S)'
    ]
    for ci, h in enumerate(headers_s2, start=1):
        cell_h = ws_detail.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        cell_h.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
    cur_row = 5
    for item in result['commune_summaries']:
        ws_detail.cell(row=cur_row, column=1, value=item['stt'])
        ws_detail.cell(row=cur_row, column=2, value=item['commune_name'])
        
        counts = item['counts']
        rates = item['rates']
        
        # C, D, E
        c3 = ws_detail.cell(row=cur_row, column=3, value=counts['C'])
        c3.number_format = COUNT_FORMAT
        c4 = ws_detail.cell(row=cur_row, column=4, value=counts['D'])
        c4.number_format = COUNT_FORMAT
        c5 = ws_detail.cell(row=cur_row, column=5, value=counts['E'])
        c5.number_format = COUNT_FORMAT
        
        # F, G
        ws_detail.cell(row=cur_row, column=6, value=item['f_gdth'])
        ws_detail.cell(row=cur_row, column=7, value=item['g_xmc'])
        
        # H, I, J
        c8 = ws_detail.cell(row=cur_row, column=8, value=counts['H'])
        c8.number_format = COUNT_FORMAT
        c9 = ws_detail.cell(row=cur_row, column=9, value=counts['I'])
        c9.number_format = COUNT_FORMAT
        c10 = ws_detail.cell(row=cur_row, column=10, value=round(rates['J'], 2))
        c10.number_format = RATE_FORMAT
        
        # K, L
        c11 = ws_detail.cell(row=cur_row, column=11, value=counts['K'])
        c11.number_format = COUNT_FORMAT
        c12 = ws_detail.cell(row=cur_row, column=12, value=round(rates['L'], 2))
        c12.number_format = RATE_FORMAT
        
        # M, N, O, P
        c13 = ws_detail.cell(row=cur_row, column=13, value=counts['M'])
        c13.number_format = COUNT_FORMAT
        c14 = ws_detail.cell(row=cur_row, column=14, value=counts['N'])
        c14.number_format = COUNT_FORMAT
        c15 = ws_detail.cell(row=cur_row, column=15, value=counts['O'])
        c15.number_format = COUNT_FORMAT
        c16 = ws_detail.cell(row=cur_row, column=16, value=round(rates['P'], 2))
        c16.number_format = RATE_FORMAT
        
        # Q, R, S
        ws_detail.cell(row=cur_row, column=17, value=item['q_gv'])
        ws_detail.cell(row=cur_row, column=18, value=item['r_csvc'])
        ws_detail.cell(row=cur_row, column=19, value=item['s_standard'])
        
        cur_row += 1
        
    # Hàng tổng cộng toàn tỉnh ở Sheet 2
    ws_detail.cell(row=cur_row, column=2, value="TỔNG CỘNG TOÀN TỈNH").font = font_bold
    p_cnt = result['province_counts']
    p_rt = result['province_rates']
    
    col_map = {
        3: ('C', COUNT_FORMAT), 4: ('D', COUNT_FORMAT), 5: ('E', COUNT_FORMAT),
        8: ('H', COUNT_FORMAT), 9: ('I', COUNT_FORMAT), 10: ('J', RATE_FORMAT),
        11: ('K', COUNT_FORMAT), 12: ('L', RATE_FORMAT),
        13: ('M', COUNT_FORMAT), 14: ('N', COUNT_FORMAT), 15: ('O', COUNT_FORMAT), 16: ('P', RATE_FORMAT)
    }
    
    for ci, (col_k, fmt) in col_map.items():
        val = p_cnt[col_k] if col_k in p_cnt else round(p_rt[col_k], 2)
        cell = ws_detail.cell(row=cur_row, column=ci, value=val)
        cell.font = font_bold
        cell.number_format = fmt
        
    ws_detail.cell(row=cur_row, column=6, value="Mức độ 3").font = font_bold
    ws_detail.cell(row=cur_row, column=7, value="Mức độ 2").font = font_bold
    ws_detail.cell(row=cur_row, column=17, value="Đạt").font = font_bold
    ws_detail.cell(row=cur_row, column=18, value="Đạt").font = font_bold
    ws_detail.cell(row=cur_row, column=19, value="Đạt chuẩn Mức độ 3").font = font_bold

    # ==========================================
    # SHEET 3: ĐỐI CHIẾU SAI LỆCH NGUỒN
    # ==========================================
    ws_disc = wb.create_sheet(title='DOI_CHIEU_SAI_LECH')
    ws_disc.views.sheetView[0].showGridLines = True
    
    ws_disc['A1'] = "BẢNG ĐỐI CHIẾU SAI LỆCH NGUỒN ĐÃ ĐƯỢC HỆ THỐNG XỬ LÝ - BIỂU 2.M3"
    ws_disc['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_disc['A2'] = "Ghi chú: So sánh số trường khai báo ở dòng thôn và số trường ở dòng tổng"
    
    headers_s3 = ['STT', 'Tên Xã / Phường', 'Hàng dòng tổng', 'Chỉ tiêu sai lệch', 'Giá trị hàng tổng nguồn', 'Giá trị tổng chi tiết thôn', 'Nguyên nhân và giải trình']
    for ci, h in enumerate(headers_s3, start=1):
        cell_h = ws_disc.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        
    cur_row_s3 = 5
    for idx, disc in enumerate(result['discrepancies'], start=1):
        ws_disc.cell(row=cur_row_s3, column=1, value=idx)
        ws_disc.cell(row=cur_row_s3, column=2, value=disc['commune'])
        ws_disc.cell(row=cur_row_s3, column=3, value=disc['tot_row'])
        ws_disc.cell(row=cur_row_s3, column=4, value=disc['field'])
        ws_disc.cell(row=cur_row_s3, column=5, value=disc['src_val'])
        ws_disc.cell(row=cur_row_s3, column=6, value=disc['clean_val'])
        ws_disc.cell(row=cur_row_s3, column=7, value=disc['reason'])
        cur_row_s3 += 1

    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_2_m4_test(result, template_path="MẪU CHUẨN - PCGD 2026(ok).xlsx", output_dir="outputs"):
    """
    Xuất file Excel tổng hợp thử cho Biểu 2.M4 (Cơ sở vật chất cấp THCS):
    - Sheet 1: TONG_HOP_THU (Cấp tỉnh - Template 2.M4 với Hàng 12 toàn tỉnh)
    - Sheet 2: CHI_TIET_92_XA (Bảng chi tiết 92 xã + dòng Tổng toàn tỉnh)
    - Sheet 3: CHI_TIET_TRUONG (Bảng chi tiết toàn bộ các trường THCS)
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_2.M4_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '2.M4':
            del wb[s_name]
            
    ws_th = wb['2.M4']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes_count']
    commune_names_str = ", ".join(result['commune_names'])
    
    ws_th['E2'] = f"GHI CHÚ: TỔNG HỢP TOÀN TỈNH (Số xã thực tế được tính: {total_c} xã - {commune_names_str})"
    font_sub = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    ws_th['E2'].font = font_sub
    
    agg = result['aggregated_cells']
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    COUNT_FORMAT = '#,##0'
    AREA_FORMAT = '#,##0.0'
    RATE_FORMAT = '0.00"%"'
    
    count_cols = ['C', 'D', 'E', 'F', 'G', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'Q', 'S', 'U', 'W']
    area_cols = ['P', 'R', 'T', 'V', 'X']
    
    # Xóa các dòng mẫu từ 8 đến 11
    for r in range(8, 12):
        for c in range(1, 26):
            col_l = get_column_letter(c)
            ws_th[f"{col_l}{r}"].value = None
            
    ws_th['B8'] = "(Số liệu chi tiết 92 xã và từng trường được tổng hợp tại Sheet 2 và Sheet 3)"
    ws_th['B8'].font = Font(name='Times New Roman', size=10, italic=True, color='555555')
    
    # 1. Điền giá trị Hàng 12 (Tổng toàn tỉnh)
    ws_th['B12'] = "TỔNG CỘNG TOÀN TỈNH"
    ws_th['B12'].font = font_bold
    
    for c in count_cols:
        coord = f"{c}12"
        if coord in agg:
            cell = ws_th[coord]
            cell.value = agg[coord]
            cell.font = font_bold
            cell.number_format = COUNT_FORMAT
            
    for c in area_cols:
        coord = f"{c}12"
        if coord in agg:
            cell = ws_th[coord]
            cell.value = agg[coord]
            cell.font = font_bold
            cell.number_format = AREA_FORMAT
            
    ws_th['H12'].value = agg.get('H12', 0.0)
    ws_th['H12'].font = font_bold
    ws_th['H12'].number_format = RATE_FORMAT
    
    ws_th.views.sheetView[0].showGridLines = True

    # ==========================================
    # SHEET 2: CHI TIẾT 92 XÃ (2.M4)
    # ==========================================
    ws_detail = wb.create_sheet(title='CHI_TIET_92_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "BẢNG TỔNG HỢP CHI TIẾT CƠ SỞ VẬT CHẤT CẤP THCS TỪNG XÃ - BIỂU 2.M4"
    ws_detail['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_detail['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Quy mô: {total_c} xã/phường"
    
    headers_s2 = [
        'STT', 'Tên Xã / Phường', 
        'Số điểm trường (C)', 'Số lớp (D)', 
        'Kiên cố (E)', 'Bán KC (F)', 'Tạm (G)', 'Tổng phòng học', 'Tỉ lệ Ph/Lớp (%) (H)',
        'Hiệu trưởng (I)', 'P.Hiệu trưởng (J)', 'Văn phòng (K)', 'Y tế (L)', 'Phòng họp (M)', 'Thư viện (N)',
        'Thí nghiệm SL (O)', 'Thí nghiệm DT (P)',
        'VS GV SL (Q)', 'VS GV DT (R)',
        'VS HS SL (S)', 'VS HS DT (T)',
        'Sân chơi SL (U)', 'Sân chơi DT (V)',
        'Bãi tập SL (W)', 'Bãi tập DT (X)'
    ]
    for ci, h in enumerate(headers_s2, start=1):
        cell_h = ws_detail.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        cell_h.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
    cur_row = 5
    for item in result['commune_summaries']:
        ws_detail.cell(row=cur_row, column=1, value=item['stt'])
        ws_detail.cell(row=cur_row, column=2, value=item['commune_name'])
        
        counts = item['counts']
        
        # C, D
        ws_detail.cell(row=cur_row, column=3, value=counts['C']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=4, value=counts['D']).number_format = COUNT_FORMAT
        
        # E, F, G, Tổng, H
        ws_detail.cell(row=cur_row, column=5, value=counts['E']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=6, value=counts['F']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=7, value=counts['G']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=8, value=item['total_rooms']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=9, value=round(item['rate_h'], 2)).number_format = RATE_FORMAT
        
        # I..N
        for idx_col, c_col in enumerate(['I', 'J', 'K', 'L', 'M', 'N'], start=10):
            ws_detail.cell(row=cur_row, column=idx_col, value=counts[c_col]).number_format = COUNT_FORMAT
            
        # O, P, Q, R, S, T, U, V, W, X
        pair_cols = [('O', 'P'), ('Q', 'R'), ('S', 'T'), ('U', 'V'), ('W', 'X')]
        col_dest = 16
        for c_sl, c_dt in pair_cols:
            ws_detail.cell(row=cur_row, column=col_dest, value=counts[c_sl]).number_format = COUNT_FORMAT
            ws_detail.cell(row=cur_row, column=col_dest+1, value=counts[c_dt]).number_format = AREA_FORMAT
            col_dest += 2
            
        cur_row += 1
        
    # Hàng tổng cộng toàn tỉnh ở Sheet 2
    ws_detail.cell(row=cur_row, column=2, value="TỔNG CỘNG TOÀN TỈNH").font = font_bold
    p_cnt = result['province_counts']
    
    ws_detail.cell(row=cur_row, column=3, value=p_cnt['C']).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=4, value=p_cnt['D']).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=5, value=p_cnt['E']).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=6, value=p_cnt['F']).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=7, value=p_cnt['G']).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=8, value=result['province_tot_rooms']).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=9, value=round(result['province_rate_h'], 2)).number_format = RATE_FORMAT
    
    for idx_col, c_col in enumerate(['I', 'J', 'K', 'L', 'M', 'N'], start=10):
        ws_detail.cell(row=cur_row, column=idx_col, value=p_cnt[c_col]).number_format = COUNT_FORMAT
        
    pair_cols = [('O', 'P'), ('Q', 'R'), ('S', 'T'), ('U', 'V'), ('W', 'X')]
    col_dest = 16
    for c_sl, c_dt in pair_cols:
        ws_detail.cell(row=cur_row, column=col_dest, value=p_cnt[c_sl]).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=col_dest+1, value=p_cnt[c_dt]).number_format = AREA_FORMAT
        col_dest += 2
        
    for c in range(1, len(headers_s2) + 1):
        ws_detail.cell(row=cur_row, column=c).font = font_bold

    # ==========================================
    # SHEET 3: CHI TIẾT TỪNG TRƯỜNG THCS
    # ==========================================
    ws_schools = wb.create_sheet(title='CHI_TIET_TRUONG')
    ws_schools.views.sheetView[0].showGridLines = True
    
    ws_schools['A1'] = "DANH SÁCH CHI TIẾT CƠ SỞ VẬT CHẤT TỪNG TRƯỜNG CẤP THCS TOÀN TỈNH"
    ws_schools['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_schools['A2'] = f"Tổng số trường THCS/liên cấp: {len(result['schools_details'])} trường"
    
    headers_s3 = ['STT', 'Mã xã', 'Tên Xã / Phường', 'Tên Trường THCS', 'Điểm trường (C)', 'Số lớp (D)', 'Kiên cố (E)', 'Bán KC (F)', 'Tạm (G)', 'Tỉ lệ Ph/Lớp (%) (H)', 'Phòng TN SL (O)', 'DT (P)', 'Sân chơi SL (U)', 'DT (V)', 'Bãi tập SL (W)', 'DT (X)']
    for ci, h in enumerate(headers_s3, start=1):
        cell_h = ws_schools.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid")
        cell_h.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
    cur_row_s3 = 5
    for idx, sch in enumerate(result['schools_details'], start=1):
        ws_schools.cell(row=cur_row_s3, column=1, value=idx)
        ws_schools.cell(row=cur_row_s3, column=2, value=sch['commune_stt'])
        ws_schools.cell(row=cur_row_s3, column=3, value=sch['commune_name'])
        ws_schools.cell(row=cur_row_s3, column=4, value=sch['school_name'])
        
        ws_schools.cell(row=cur_row_s3, column=5, value=sch['C']).number_format = COUNT_FORMAT
        ws_schools.cell(row=cur_row_s3, column=6, value=sch['D']).number_format = COUNT_FORMAT
        ws_schools.cell(row=cur_row_s3, column=7, value=sch['E']).number_format = COUNT_FORMAT
        ws_schools.cell(row=cur_row_s3, column=8, value=sch['F']).number_format = COUNT_FORMAT
        ws_schools.cell(row=cur_row_s3, column=9, value=sch['G']).number_format = COUNT_FORMAT
        ws_schools.cell(row=cur_row_s3, column=10, value=round(sch['rate_h'], 2)).number_format = RATE_FORMAT
        ws_schools.cell(row=cur_row_s3, column=11, value=sch['O']).number_format = COUNT_FORMAT
        ws_schools.cell(row=cur_row_s3, column=12, value=sch['P']).number_format = AREA_FORMAT
        ws_schools.cell(row=cur_row_s3, column=13, value=sch['U']).number_format = COUNT_FORMAT
        ws_schools.cell(row=cur_row_s3, column=14, value=sch['V']).number_format = AREA_FORMAT
        ws_schools.cell(row=cur_row_s3, column=15, value=sch['W']).number_format = COUNT_FORMAT
        ws_schools.cell(row=cur_row_s3, column=16, value=sch['X']).number_format = AREA_FORMAT
        cur_row_s3 += 1

    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_2_m5_test(result, template_path="MẪU CHUẨN - PCGD 2026(ok).xlsx", output_dir="outputs"):
    """
    Xuất file Excel tổng hợp thử cho Biểu 2.M5 (Đội ngũ giáo viên, nhân viên THCS):
    - Sheet 1: TONG_HOP_THU (Cấp tỉnh - Template 2.M5 với Hàng 12 toàn tỉnh & Tiêu chí)
    - Sheet 2: CHI_TIET_92_XA (Bảng chi tiết 92 xã + dòng Tổng toàn tỉnh)
    - Sheet 3: CHI_TIET_TRUONG (Bảng chi tiết toàn bộ các trường THCS)
    - Sheet 4: DOI_CHIEU_SAI_LECH (Bảng đối chiếu sai lệch nguồn)
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_2.M5_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '2.M5':
            del wb[s_name]
            
    ws_th = wb['2.M5']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes_count']
    commune_names_str = ", ".join(result['commune_names'])
    
    ws_th['J2'] = f"GHI CHÚ: TỔNG HỢP TOÀN TỈNH (Số xã thực tế được tính: {total_c} xã - {commune_names_str})"
    font_sub = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    ws_th['J2'].font = font_sub
    
    agg = result['aggregated_cells']
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    COUNT_FORMAT = '#,##0'
    RATIO_FORMAT = '0.00'
    RATE_FORMAT = '0.00"%"'
    
    # Xóa các dòng mẫu từ 8 đến 11
    for r in range(8, 12):
        for c in range(1, 45):
            col_l = get_column_letter(c)
            ws_th[f"{col_l}{r}"].value = None
            
    ws_th['B8'] = "(Số liệu chi tiết 92 xã và từng trường được tổng hợp tại Sheet 2 và Sheet 3)"
    ws_th['B8'].font = Font(name='Times New Roman', size=10, italic=True, color='555555')
    
    # 1. Điền giá trị Hàng 12 (Tổng toàn tỉnh)
    ws_th['B12'] = "TỔNG CỘNG TOÀN TỈNH"
    ws_th['B12'].font = font_bold
    
    cols_int = [
        'F', 'G', 'H', 'I', 'J', 'K', 'L',
        'N', 'O', 'P', 'Q',
        'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH',
        'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ'
    ]
    
    for c in cols_int:
        coord = f"{c}12"
        if coord in agg:
            cell = ws_th[coord]
            cell.value = agg[coord]
            cell.font = font_bold
            cell.number_format = COUNT_FORMAT
            
    ws_th['M12'].value = agg.get('M12', 0.0)
    ws_th['M12'].font = font_bold
    ws_th['M12'].number_format = RATIO_FORMAT
    
    # 2. Điền Bảng tiêu chí Hàng 15..17
    crit = result['criteria']
    # Đạt chuẩn
    ws_th['I15'].value = crit['dat_chuan']['sl']
    ws_th['I15'].font = font_bold
    ws_th['I15'].number_format = COUNT_FORMAT
    ws_th['J15'].value = crit['dat_chuan']['tl']
    ws_th['J15'].font = font_bold
    ws_th['J15'].number_format = RATE_FORMAT
    
    # Trên chuẩn
    ws_th['I16'].value = crit['tren_chuan']['sl']
    ws_th['I16'].font = font_bold
    ws_th['I16'].number_format = COUNT_FORMAT
    ws_th['J16'].value = crit['tren_chuan']['tl']
    ws_th['J16'].font = font_bold
    ws_th['J16'].number_format = RATE_FORMAT
    
    # Chuẩn nghề nghiệp
    ws_th['I17'].value = crit['chuan_nghe_nghiep']['sl']
    ws_th['I17'].font = font_bold
    ws_th['I17'].number_format = COUNT_FORMAT
    ws_th['J17'].value = crit['chuan_nghe_nghiep']['tl']
    ws_th['J17'].font = font_bold
    ws_th['J17'].number_format = RATE_FORMAT
    
    ws_th.views.sheetView[0].showGridLines = True

    # ==========================================
    # SHEET 2: CHI TIẾT 92 XÃ (2.M5)
    # ==========================================
    ws_detail = wb.create_sheet(title='CHI_TIET_92_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "BẢNG TỔNG HỢP CHI TIẾT ĐỘI NGŨ GIÁO VIÊN CẤP THCS TỪNG XÃ - BIỂU 2.M5"
    ws_detail['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_detail['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Quy mô: {total_c} xã/phường"
    
    headers_s2 = [
        'STT', 'Tên Xã / Phường', 
        'Hiệu trưởng (F)', 'P.Hiệu trưởng (G)',
        'Tổng số GV (H)', 'Biên chế (I)', 'Hợp đồng (J)', 'Nữ (K)', 'Dân tộc (L)',
        'Trên ĐH (N)', 'Đại học (O)', 'Cao đẳng (P)', 'THSP (Q)',
        'Toán (R)', 'Văn (S)', 'KHTN-Lí (T)', 'Hóa (U)', 'Sinh (V)', 'KHXH-Sử (W)', 'Địa (X)',
        'Nhạc (Y)', 'Mỹ thuật (Z)', 'Thể dục (AA)', 'GDCD (AB)', 'Công nghệ (AC)', 'Tin học (AD)', 'Tiếng Anh (AE)',
        'TPT Đội (AI)', 'NN Tốt (AJ)', 'Khá (AK)', 'Đạt (AL)', 'Chưa đạt (AM)',
        'Thư viện (AN)', 'Thiết bị (AO)', 'Văn phòng (AP)', 'Y tế (AQ)'
    ]
    for ci, h in enumerate(headers_s2, start=1):
        cell_h = ws_detail.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        cell_h.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
    cols_detail_order = [
        'F', 'G', 'H', 'I', 'J', 'K', 'L',
        'N', 'O', 'P', 'Q',
        'R', 'S', 'T', 'U', 'V', 'W', 'X',
        'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE',
        'AI', 'AJ', 'AK', 'AL', 'AM',
        'AN', 'AO', 'AP', 'AQ'
    ]
    
    cur_row = 5
    for item in result['commune_summaries']:
        ws_detail.cell(row=cur_row, column=1, value=item['stt'])
        ws_detail.cell(row=cur_row, column=2, value=item['commune_name'])
        
        counts = item['counts']
        for ci, col_k in enumerate(cols_detail_order, start=3):
            cell_c = ws_detail.cell(row=cur_row, column=ci, value=counts[col_k])
            cell_c.number_format = COUNT_FORMAT
            
        cur_row += 1
        
    # Hàng tổng cộng toàn tỉnh ở Sheet 2
    ws_detail.cell(row=cur_row, column=2, value="TỔNG CỘNG TOÀN TỈNH").font = font_bold
    p_cnt = result['province_counts']
    for ci, col_k in enumerate(cols_detail_order, start=3):
        cell_c = ws_detail.cell(row=cur_row, column=ci, value=p_cnt[col_k])
        cell_c.font = font_bold
        cell_c.number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=1).font = font_bold

    # ==========================================
    # SHEET 3: CHI TIẾT TỪNG TRƯỜNG THCS
    # ==========================================
    ws_schools = wb.create_sheet(title='CHI_TIET_TRUONG')
    ws_schools.views.sheetView[0].showGridLines = True
    
    ws_schools['A1'] = "DANH SÁCH CHI TIẾT ĐỘI NGŨ GIÁO VIÊN TỪNG TRƯỜNG CẤP THCS TOÀN TỈNH"
    ws_schools['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_schools['A2'] = f"Tổng số trường THCS/liên cấp: {len(result['schools_details'])} trường"
    
    headers_s3 = ['STT', 'Mã xã', 'Tên Xã / Phường', 'Tên Trường THCS', 'Tổng số GV (H)', 'Biên chế (I)', 'Hợp đồng (J)', 'Nữ (K)', 'Dân tộc (L)', 'Trên ĐH (N)', 'Đại học (O)', 'Cao đẳng (P)', 'THSP (Q)', 'TPT Đội (AI)']
    for ci, h in enumerate(headers_s3, start=1):
        cell_h = ws_schools.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid")
        cell_h.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
    cur_row_s3 = 5
    for idx, sch in enumerate(result['schools_details'], start=1):
        ws_schools.cell(row=cur_row_s3, column=1, value=idx)
        ws_schools.cell(row=cur_row_s3, column=2, value=sch['commune_stt'])
        ws_schools.cell(row=cur_row_s3, column=3, value=sch['commune_name'])
        ws_schools.cell(row=cur_row_s3, column=4, value=sch['school_name'])
        
        for ci, col_k in enumerate(['H', 'I', 'J', 'K', 'L', 'N', 'O', 'P', 'Q', 'AI'], start=5):
            ws_schools.cell(row=cur_row_s3, column=ci, value=sch[col_k]).number_format = COUNT_FORMAT
        cur_row_s3 += 1

    # ==========================================
    # SHEET 4: ĐỐI CHIẾU SAI LỆCH NGUỒN
    # ==========================================
    ws_disc = wb.create_sheet(title='DOI_CHIEU_SAI_LECH')
    ws_disc.views.sheetView[0].showGridLines = True
    
    ws_disc['A1'] = "BẢNG ĐỐI CHIẾU SAI LỆCH NGUỒN GIỮA TỔNG SỐ GV VÀ BIÊN CHẾ + HỢP ĐỒNG - BIỂU 2.M5"
    ws_disc['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_disc['A2'] = "Ghi chú: So sánh Tổng số GV nguồn (H) so với Biên chế (I) + Hợp đồng (J) và giải trình theo trình độ đào tạo"
    
    headers_s4 = ['STT', 'Mã xã', 'Tên Xã / Phường', 'Hàng dòng tổng', 'Tổng GV nguồn (H)', 'Biên chế (I)', 'Hợp đồng (J)', 'Tổng BC+HĐ (I+J)', 'Chênh lệch', 'Nguyên nhân và giải trình']
    for ci, h in enumerate(headers_s4, start=1):
        cell_h = ws_disc.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        cell_h.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
    cur_row_s4 = 5
    for idx, disc in enumerate(result['discrepancies'], start=1):
        ws_disc.cell(row=cur_row_s4, column=1, value=idx)
        ws_disc.cell(row=cur_row_s4, column=2, value=disc['stt'])
        ws_disc.cell(row=cur_row_s4, column=3, value=disc['commune'])
        ws_disc.cell(row=cur_row_s4, column=4, value=disc['tot_row'])
        
        c5 = ws_disc.cell(row=cur_row_s4, column=5, value=disc['h_src'])
        c5.number_format = COUNT_FORMAT
        c6 = ws_disc.cell(row=cur_row_s4, column=6, value=disc['i_bc'])
        c6.number_format = COUNT_FORMAT
        c7 = ws_disc.cell(row=cur_row_s4, column=7, value=disc['j_hd'])
        c7.number_format = COUNT_FORMAT
        c8 = ws_disc.cell(row=cur_row_s4, column=8, value=disc['calc_h'])
        c8.number_format = COUNT_FORMAT
        c8.font = font_bold
        c9 = ws_disc.cell(row=cur_row_s4, column=9, value=disc['diff'])
        c9.number_format = COUNT_FORMAT
        c9.font = font_bold
        
        ws_disc.cell(row=cur_row_s4, column=10, value=disc['reason'])
        cur_row_s4 += 1

    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_2_m6_test(result, template_path="MẪU CHUẨN - PCGD 2026(ok).xlsx", output_dir="outputs"):
    """
    Xuất file Excel tổng hợp thử cho Biểu 2.M6 (Đối tượng khuyết tật THCS):
    - Sheet 1: TONG_HOP_THU (Cấp tỉnh - Template 2.M6 với Hàng 6..14 toàn tỉnh)
    - Sheet 2: CHI_TIET_92_XA (Bảng chi tiết 92 xã theo từng độ tuổi + Dòng Tổng toàn tỉnh)
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_2.M6_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '2.M6':
            del wb[s_name]
            
    ws_th = wb['2.M6']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes_count']
    commune_names_str = ", ".join(result['commune_names'])
    
    ws_th['D3'] = f"GHI CHÚ: TỔNG HỢP TOÀN TỈNH (Số xã thực tế được tính: {total_c} xã - {commune_names_str})"
    font_sub = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    ws_th['D3'].font = font_sub
    
    agg = result['aggregated_cells']
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    COUNT_FORMAT = '#,##0'
    RATE_FORMAT = '0.00"%"'
    
    cols_count = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    
    # 1. Điền giá trị Hàng 6 đến 14 (Tổng toàn tỉnh)
    for r in range(6, 15):
        is_total_row = (r == 14)
        for c in cols_count:
            coord = f"{c}{r}"
            if coord in agg:
                cell = ws_th[coord]
                cell.value = agg[coord]
                if is_total_row:
                    cell.font = font_bold
                cell.number_format = COUNT_FORMAT
                
        # Cột M: Tỷ lệ %
        coord_m = f"M{r}"
        if coord_m in agg:
            cell_m = ws_th[coord_m]
            cell_m.value = agg[coord_m]
            if is_total_row:
                cell_m.font = font_bold
            cell_m.number_format = RATE_FORMAT
            
    ws_th.views.sheetView[0].showGridLines = True

    # ==========================================
    # SHEET 2: CHI TIẾT 92 XÃ (2.M6)
    # ==========================================
    ws_detail = wb.create_sheet(title='CHI_TIET_92_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "BẢNG TỔNG HỢP CHI TIẾT ĐỐI TƯỢNG KHUYẾT TẬT CẤP THCS TỪNG XÃ - BIỂU 2.M6"
    ws_detail['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_detail['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Quy mô: {total_c} xã/phường"
    
    headers_s2 = [
        'STT', 'Tên Xã / Phường', 'Năm sinh / Độ tuổi',
        'Tổng số đối tượng KT (C)', 'Vận động (D)', 'Nghe, nói (E)', 'Nhìn (F)', 'Thần kinh tâm thần (G)',
        'Trí tuệ (H)', 'Tự kỷ (I)', 'Học tập (J)', 'Khác (K)',
        'Tiếp cận GD (L)', 'Tỷ lệ tiếp cận (%) (M)'
    ]
    for ci, h in enumerate(headers_s2, start=1):
        cell_h = ws_detail.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        cell_h.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
    birth_years = {
        6: "2015 (11 tuổi)", 7: "2014 (12 tuổi)", 8: "2013 (13 tuổi)", 9: "2012 (14 tuổi)",
        10: "2011 (15 tuổi)", 11: "2010 (16 tuổi)", 12: "2009 (17 tuổi)", 13: "2008 (18 tuổi)",
        14: "TỔNG CỘNG"
    }
    
    cur_row = 5
    for cname, cdata in result['commune_details'].items():
        cstt = cdata['stt']
        cmatrix = cdata['matrix']
        for r in range(6, 15):
            ws_detail.cell(row=cur_row, column=1, value=cstt)
            ws_detail.cell(row=cur_row, column=2, value=cname)
            ws_detail.cell(row=cur_row, column=3, value=birth_years[r])
            
            c_val = cmatrix.get(f"C{r}", 0.0)
            l_val = cmatrix.get(f"L{r}", 0.0)
            r_val = (l_val / c_val * 100) if c_val > 0 else 0.0
            
            for ci, col_k in enumerate(cols_count, start=4):
                val = cmatrix.get(f"{col_k}{r}", 0.0)
                cell_c = ws_detail.cell(row=cur_row, column=ci, value=val)
                cell_c.number_format = COUNT_FORMAT
                if r == 14:
                    cell_c.font = font_bold
                    
            cell_m = ws_detail.cell(row=cur_row, column=14, value=round(r_val, 2))
            cell_m.number_format = RATE_FORMAT
            if r == 14:
                cell_m.font = font_bold
                ws_detail.cell(row=cur_row, column=3).font = font_bold
                
            cur_row += 1
            
    # Hàng tổng cộng toàn tỉnh ở Sheet 2
    for r in range(6, 15):
        ws_detail.cell(row=cur_row, column=1, value="")
        ws_detail.cell(row=cur_row, column=2, value="TOÀN TỈNH")
        ws_detail.cell(row=cur_row, column=3, value=birth_years[r])
        
        for ci, col_k in enumerate(cols_count, start=4):
            cell_c = ws_detail.cell(row=cur_row, column=ci, value=agg[f"{col_k}{r}"])
            cell_c.font = font_bold
            cell_c.number_format = COUNT_FORMAT
            
        cell_m = ws_detail.cell(row=cur_row, column=14, value=agg[f"M{r}"])
        cell_m.font = font_bold
        cell_m.number_format = RATE_FORMAT
        cur_row += 1

    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_2_m7_test(result, template_path="MẪU CHUẨN - PCGD 2026(ok).xlsx", output_dir="outputs"):
    """
    Xuất kết quả thử nghiệm biểu 2.M7 (Học sinh học ngoài xã, ngoài tỉnh cấp THCS):
    - Sheet 1: TONG_HOP_THU (Mẫu 2.M7 cấp tỉnh: Hàng 21, 39, 40)
    - Sheet 2: CHI_TIET_THEO_XA (Thống kê số lượng từng xã theo độ tuổi 10..18)
    - Sheet 3: DANH_SACH_TRUONG_NGOAI (Danh sách các trường ngoài địa bàn được khai báo)
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_2.M7_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '2.M7':
            del wb[s_name]
            
    ws_th = wb['2.M7']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes_count']
    commune_names_str = ", ".join(result['commune_names'])
    
    ws_th['C1'] = "TỔNG HỢP THỬ - HỌC SINH ĐANG HỌC TẠI XÃ/PHƯỜNG NGOÀI VÀ TỈNH NGOÀI CẤP THCS"
    ws_th['C2'] = f"GHI CHÚ: TỔNG HỢP THỬ NGHIỆM (Số xã thực tế được tính: {total_c} xã - {commune_names_str})"
    font_sub = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    ws_th['C2'].font = font_sub
    
    agg = result['aggregated_cells']
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    COUNT_FORMAT = '#,##0'
    
    # 1. Điền giá trị Hàng 21 (Tổng Phần I), Hàng 39 (Tổng Phần II), Hàng 40 (Tổng I + II)
    cols_2m7 = [get_column_letter(c) for c in range(3, 40)] # C..AM
    for r in [21, 39, 40]:
        for col in cols_2m7:
            coord = f"{col}{r}"
            if coord in agg:
                val = agg[coord]
                cell = ws_th[coord]
                cell.value = val
                cell.font = font_bold
                if isinstance(val, (int, float)):
                    cell.number_format = COUNT_FORMAT

    ws_th.views.sheetView[0].showGridLines = True

    # ==========================================
    # SHEET 2: CHI TIẾT THEO TỪNG XÃ (2.M7)
    # ==========================================
    ws_detail = wb.create_sheet(title='CHI_TIET_THEO_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "BẢNG TỔNG HỢP SỐ LIỆU HỌC SINH HỌC NGOÀI ĐỊA BÀN TỪNG XÃ - BIỂU 2.M7"
    ws_detail['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_detail['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Quy mô: {total_c} xã/phường"
    
    headers_s2 = ['STT', 'Tên Xã / Phường', 'Tên File Nguồn', 'Học ngoài xã (trong tỉnh)', 'Học ngoài tỉnh', 'Tổng cộng học sinh đi học ngoài']
    for ci, h in enumerate(headers_s2, start=1):
        cell_h = ws_detail.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        
    cur_row = 5
    for item in result['commune_summaries']:
        ws_detail.cell(row=cur_row, column=1, value=item['stt'])
        ws_detail.cell(row=cur_row, column=2, value=item['commune_name'])
        ws_detail.cell(row=cur_row, column=3, value=item['filename'])
        
        c4 = ws_detail.cell(row=cur_row, column=4, value=item['part1_sum'])
        c4.number_format = COUNT_FORMAT
        c5 = ws_detail.cell(row=cur_row, column=5, value=item['part2_sum'])
        c5.number_format = COUNT_FORMAT
        c6 = ws_detail.cell(row=cur_row, column=6, value=item['total_sum'])
        c6.number_format = COUNT_FORMAT
        c6.font = font_bold
        cur_row += 1
        
    # Hàng tổng cộng toàn tỉnh ở Sheet 2
    ws_detail.cell(row=cur_row, column=2, value="TỔNG CỘNG TOÀN TỈNH").font = font_bold
    c_p1 = ws_detail.cell(row=cur_row, column=4, value=result['total_part1'])
    c_p1.number_format = COUNT_FORMAT
    c_p1.font = font_bold
    c_p2 = ws_detail.cell(row=cur_row, column=5, value=result['total_part2'])
    c_p2.number_format = COUNT_FORMAT
    c_p2.font = font_bold
    c_gt = ws_detail.cell(row=cur_row, column=6, value=result['total_grand'])
    c_gt.number_format = COUNT_FORMAT
    c_gt.font = font_bold

    # ==========================================
    # SHEET 3: DANH SÁCH TRƯỜNG HỌC KHAI BÁO
    # ==========================================
    ws_schools = wb.create_sheet(title='DANH_SACH_TRUONG_NGOAI')
    ws_schools.views.sheetView[0].showGridLines = True
    
    ws_schools['A1'] = "DANH SÁCH CÁC TRƯỜNG HỌC NGOÀI ĐỊA BÀN ĐƯỢC 92 XÃ/PHƯỜNG KHAI BÁO - BIỂU 2.M7"
    ws_schools['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_schools['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Tổng số trường/cơ sở ghi nhận: {len(result['school_records'])} lượt khai báo"
    
    headers_s3 = ['STT', 'Mã/STT Xã', 'Xã / Phường khai báo', 'Phân loại', 'Tên trường học ngoài địa bàn', 'Tổng số học sinh']
    for ci, h in enumerate(headers_s3, start=1):
        cell_h = ws_schools.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        
    cur_row_s3 = 5
    for idx, rec in enumerate(result['school_records'], start=1):
        ws_schools.cell(row=cur_row_s3, column=1, value=idx)
        ws_schools.cell(row=cur_row_s3, column=2, value=rec['commune_stt'])
        ws_schools.cell(row=cur_row_s3, column=3, value=rec['commune_name'])
        ws_schools.cell(row=cur_row_s3, column=4, value=rec['part'])
        ws_schools.cell(row=cur_row_s3, column=5, value=rec['school_name'])
        
        c_tot = ws_schools.cell(row=cur_row_s3, column=6, value=rec['total_students'])
        c_tot.number_format = COUNT_FORMAT
        c_tot.font = font_bold
        cur_row_s3 += 1

    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_2_m8_test(result, template_path="MẪU CHUẨN - PCGD 2026(ok).xlsx", output_dir="outputs"):
    """
    Xuất kết quả thử nghiệm biểu 2.M8 (Học sinh ngoại địa bàn học tại xã cấp THCS):
    - Sheet 1: TONG_HOP_THU (Mẫu 2.M8 cấp tỉnh: Hàng 24, 38, 39)
    - Sheet 2: CHI_TIET_THEO_XA (Thống kê số lượng từng xã theo độ tuổi 10..18)
    - Sheet 3: DANH_SACH_TRUONG_NGOAI (Danh sách các trường ngoài địa bàn được khai báo)
    - Sheet 4: DOI_CHIEU_SAI_LECH (Bảng đối soát các xã lệch nguồn giữa Cột AM và Tổng C..AL)
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_2.M8_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '2.M8':
            del wb[s_name]
            
    ws_th = wb['2.M8']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes_count']
    commune_names_str = ", ".join(result['commune_names'])
    
    ws_th['C2'] = "TỔNG HỢP THỬ - HỌC SINH XÃ/PHƯỜNG NGOÀI VÀ TỈNH NGOÀI ĐANG HỌC TẠI ĐỊA PHƯƠNG CẤP THCS"
    ws_th['C3'] = f"GHI CHÚ: TỔNG HỢP THỬ NGHIỆM (Số xã thực tế được tính: {total_c} xã - {commune_names_str})"
    font_sub = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    ws_th['C3'].font = font_sub
    
    agg = result['aggregated_cells']
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    COUNT_FORMAT = '#,##0'
    
    # 1. Điền giá trị Hàng 24 (Tổng Phần I), Hàng 38 (Tổng Phần II), Hàng 39 (Tổng số I + II)
    cols_2m8 = [get_column_letter(c) for c in range(3, 40)] # C..AM
    for r in [24, 38, 39]:
        for col in cols_2m8:
            coord = f"{col}{r}"
            if coord in agg:
                val = agg[coord]
                cell = ws_th[coord]
                cell.value = val
                cell.font = font_bold
                if isinstance(val, (int, float)):
                    cell.number_format = COUNT_FORMAT

    ws_th.views.sheetView[0].showGridLines = True

    # ==========================================
    # SHEET 2: CHI TIẾT THEO TỪNG XÃ (2.M8)
    # ==========================================
    ws_detail = wb.create_sheet(title='CHI_TIET_THEO_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "BẢNG TỔNG HỢP SỐ LIỆU HỌC SINH NGOẠI ĐỊA BÀN ĐẾN HỌC TỪNG XÃ - BIỂU 2.M8"
    ws_detail['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_detail['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Quy mô: {total_c} xã/phường"
    
    headers_s2 = ['STT', 'Tên Xã / Phường', 'Tên File Nguồn', 'Học ngoài xã (trong tỉnh) đến học', 'Ngoài tỉnh đến học', 'Tổng cộng học sinh ngoại địa bàn']
    for ci, h in enumerate(headers_s2, start=1):
        cell_h = ws_detail.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        
    cur_row = 5
    for item in result['commune_summaries']:
        ws_detail.cell(row=cur_row, column=1, value=item['stt'])
        ws_detail.cell(row=cur_row, column=2, value=item['commune_name'])
        ws_detail.cell(row=cur_row, column=3, value=item['filename'])
        
        c4 = ws_detail.cell(row=cur_row, column=4, value=item['part1_sum'])
        c4.number_format = COUNT_FORMAT
        c5 = ws_detail.cell(row=cur_row, column=5, value=item['part2_sum'])
        c5.number_format = COUNT_FORMAT
        c6 = ws_detail.cell(row=cur_row, column=6, value=item['total_sum'])
        c6.number_format = COUNT_FORMAT
        c6.font = font_bold
        cur_row += 1
        
    # Hàng tổng cộng toàn tỉnh ở Sheet 2
    ws_detail.cell(row=cur_row, column=2, value="TỔNG CỘNG TOÀN TỈNH").font = font_bold
    c_p1 = ws_detail.cell(row=cur_row, column=4, value=result['total_part1'])
    c_p1.number_format = COUNT_FORMAT
    c_p1.font = font_bold
    c_p2 = ws_detail.cell(row=cur_row, column=5, value=result['total_part2'])
    c_p2.number_format = COUNT_FORMAT
    c_p2.font = font_bold
    c_gt = ws_detail.cell(row=cur_row, column=6, value=result['total_grand'])
    c_gt.number_format = COUNT_FORMAT
    c_gt.font = font_bold

    # ==========================================
    # SHEET 3: DANH SÁCH TRƯỜNG HỌC KHAI BÁO
    # ==========================================
    ws_schools = wb.create_sheet(title='DANH_SACH_TRUONG_NGOAI')
    ws_schools.views.sheetView[0].showGridLines = True
    
    ws_schools['A1'] = "DANH SÁCH CÁC TRƯỜNG HỌC ĐƯỢC 92 XÃ/PHƯỜNG KHAI BÁO - BIỂU 2.M8"
    ws_schools['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_schools['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Tổng số trường/cơ sở ghi nhận: {len(result['school_records'])} lượt khai báo"
    
    headers_s3 = ['STT', 'Mã/STT Xã', 'Xã / Phường khai báo', 'Phân loại', 'Tên trường học khai báo', 'Tổng số học sinh']
    for ci, h in enumerate(headers_s3, start=1):
        cell_h = ws_schools.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        
    cur_row_s3 = 5
    for idx, rec in enumerate(result['school_records'], start=1):
        ws_schools.cell(row=cur_row_s3, column=1, value=idx)
        ws_schools.cell(row=cur_row_s3, column=2, value=rec['commune_stt'])
        ws_schools.cell(row=cur_row_s3, column=3, value=rec['commune_name'])
        ws_schools.cell(row=cur_row_s3, column=4, value=rec['part'])
        ws_schools.cell(row=cur_row_s3, column=5, value=rec['school_name'])
        
        c_tot = ws_schools.cell(row=cur_row_s3, column=6, value=rec['total_students'])
        c_tot.number_format = COUNT_FORMAT
        c_tot.font = font_bold
        cur_row_s3 += 1

    # ==========================================
    # SHEET 4: ĐỐI SOÁT SAI LỆCH NGUỒN CỘT AM
    # ==========================================
    ws_disc = wb.create_sheet(title='DOI_CHIEU_SAI_LECH')
    ws_disc.views.sheetView[0].showGridLines = True
    
    ws_disc['A1'] = "BẢNG ĐỐI SOÁT CÁC XÃ CÓ SAI LỆCH NGUỒN GIỮA CỘT AM VÀ TỔNG CÁC CỘT TUỔI C..AL - BIỂU 2.M8"
    ws_disc['A1'].font = Font(name='Times New Roman', size=12, bold=True)
    ws_disc['A2'] = "Lưu ý: Hệ thống tính lại chuẩn Cột AM = sum(C..AL) để bảo toàn tuyệt đối số lượng học sinh thực tế."
    
    headers_s4 = ['STT', 'Mã/STT Xã', 'Tên Xã / Phường', 'Phân loại', 'Số tính chuẩn từ C..AL', 'Số tĩnh nguồn ô AM', 'Chênh lệch', 'Lý do giải trình']
    for ci, h in enumerate(headers_s4, start=1):
        cell_h = ws_disc.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        
    cur_row_s4 = 5
    for idx, d in enumerate(result['discrepancies'], start=1):
        ws_disc.cell(row=cur_row_s4, column=1, value=idx)
        ws_disc.cell(row=cur_row_s4, column=2, value=d['commune_stt'])
        ws_disc.cell(row=cur_row_s4, column=3, value=d['commune_name'])
        ws_disc.cell(row=cur_row_s4, column=4, value=d['part'])
        
        c5 = ws_disc.cell(row=cur_row_s4, column=5, value=d['calc_am'])
        c5.number_format = COUNT_FORMAT
        c6 = ws_disc.cell(row=cur_row_s4, column=6, value=d['src_am'])
        c6.number_format = COUNT_FORMAT
        c7 = ws_disc.cell(row=cur_row_s4, column=7, value=d['diff'])
        c7.number_format = COUNT_FORMAT
        c7.font = font_bold
        
        ws_disc.cell(row=cur_row_s4, column=8, value=d['reason'])
        cur_row_s4 += 1

    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_2_m9_test(result, template_path="MẪU CHUẨN - PCGD 2026(ok).xlsx", output_dir="outputs"):
    """
    Xuất kết quả thử nghiệm biểu 2.M9 (Huy động phổ cập giáo dục trung học):
    - Sheet 1: TONG_HOP_THU (Mẫu 2.M9 cấp tỉnh: Hàng 48)
    - Sheet 2: CHI_TIET_92_XA (Bảng chi tiết số liệu từng xã)
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_2.M9_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '2.M9':
            del wb[s_name]
            
    ws_th = wb['2.M9']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes_count']
    commune_names_str = ", ".join(result['commune_names'])
    
    ws_th['C3'] = "TỔNG HỢP THỬ - BIỂU TỔNG HỢP KẾT QUẢ HUY ĐỘNG PHỔ CẬP GIÁO DỤC TRUNG HỌC"
    ws_th['C4'] = f"GHI CHÚ: TỔNG HỢP THỬ NGHIỆM (Số xã thực tế được tính: {total_c} xã - {commune_names_str})"
    font_sub = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    ws_th['C4'].font = font_sub
    
    agg = result['aggregated_cells']
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    COUNT_FORMAT = '#,##0'
    RATE_FORMAT = '0.00"%"'
    
    count_cols = ['C', 'D', 'E', 'G', 'H', 'I', 'J', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S']
    rate_cols = ['F', 'K', 'T', 'U', 'V']
    
    # 1. Điền giá trị Hàng 48 cấp tỉnh
    for c in count_cols:
        coord = f"{c}48"
        if coord in agg:
            cell = ws_th[coord]
            cell.value = agg[coord]
            cell.font = font_bold
            cell.number_format = COUNT_FORMAT
            
    for c in rate_cols:
        coord = f"{c}48"
        if coord in agg:
            cell = ws_th[coord]
            cell.value = agg[coord]
            cell.font = font_bold
            cell.number_format = RATE_FORMAT

    ws_th.views.sheetView[0].showGridLines = True

    # ==========================================
    # SHEET 2: CHI TIẾT 92 XÃ (2.M9)
    # ==========================================
    ws_detail = wb.create_sheet(title='CHI_TIET_92_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "BẢNG TỔNG HỢP KẾT QUẢ HUY ĐỘNG PCGD TRUNG HỌC TỪNG XÃ - BIỂU 2.M9"
    ws_detail['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_detail['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Quy mô: {total_c} xã/phường"
    
    headers_s2 = [
        'STT', 'Tên Xã / Phường', 'Dòng nguồn',
        'Tổng số 15-21 (C)', 'Phải PC 15-18 (D)', 'TN THCS 15-18 (E)', 'Tỷ lệ % (F)',
        'Còn học THCS (G)', 'Bỏ học THCS (H)',
        'Lớp 9 năm qua (I)', 'TN THCS năm qua (J)', 'Tỷ lệ % (K)',
        'Vào lớp 10 THPT (L)', 'Học GDTX/Nghề (M)', 'Bỏ học THPT (N)',
        'Tổng phải PC 18-21 (O)', 'Học THPT (P)', 'Học GDTX (Q)', 'Học TCCN (R)', 'Học Nghề (S)',
        'Tỷ lệ TCCN (T)', 'Tỷ lệ Nghề (U)', 'Tỷ lệ TN THPT/Nghề (V)'
    ]
    for ci, h in enumerate(headers_s2, start=1):
        cell_h = ws_detail.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        cell_h.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
    cur_row = 5
    for item in result['commune_summaries']:
        ws_detail.cell(row=cur_row, column=1, value=item['stt'])
        ws_detail.cell(row=cur_row, column=2, value=item['commune_name'])
        ws_detail.cell(row=cur_row, column=3, value=item['row_total'])
        
        c_cnt = item['counts']
        c_rt = item['rates']
        
        # C, D, E
        ws_detail.cell(row=cur_row, column=4, value=c_cnt['C']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=5, value=c_cnt['D']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=6, value=c_cnt['E']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=7, value=c_rt['F']).number_format = RATE_FORMAT
        
        # G, H
        ws_detail.cell(row=cur_row, column=8, value=c_cnt['G']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=9, value=c_cnt['H']).number_format = COUNT_FORMAT
        
        # I, J, K
        ws_detail.cell(row=cur_row, column=10, value=c_cnt['I']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=11, value=c_cnt['J']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=12, value=c_rt['K']).number_format = RATE_FORMAT
        
        # L, M, N
        ws_detail.cell(row=cur_row, column=13, value=c_cnt['L']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=14, value=c_cnt['M']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=15, value=c_cnt['N']).number_format = COUNT_FORMAT
        
        # O, P, Q, R, S, T, U, V
        ws_detail.cell(row=cur_row, column=16, value=c_cnt['O']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=17, value=c_cnt['P']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=18, value=c_cnt['Q']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=19, value=c_cnt['R']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=20, value=c_cnt['S']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=21, value=c_rt['T']).number_format = RATE_FORMAT
        ws_detail.cell(row=cur_row, column=22, value=c_rt['U']).number_format = RATE_FORMAT
        ws_detail.cell(row=cur_row, column=23, value=c_rt['V']).number_format = RATE_FORMAT
        
        cur_row += 1
        
    # Hàng tổng cộng toàn tỉnh ở Sheet 2
    ws_detail.cell(row=cur_row, column=2, value="TỔNG CỘNG TOÀN TỈNH").font = font_bold
    p_cnt = result['province_counts']
    p_rt = result['province_rates']
    
    ws_detail.cell(row=cur_row, column=4, value=p_cnt['C']).font = font_bold
    ws_detail.cell(row=cur_row, column=4).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=5, value=p_cnt['D']).font = font_bold
    ws_detail.cell(row=cur_row, column=5).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=6, value=p_cnt['E']).font = font_bold
    ws_detail.cell(row=cur_row, column=6).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=7, value=p_rt['F']).font = font_bold
    ws_detail.cell(row=cur_row, column=7).number_format = RATE_FORMAT
    
    ws_detail.cell(row=cur_row, column=8, value=p_cnt['G']).font = font_bold
    ws_detail.cell(row=cur_row, column=8).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=9, value=p_cnt['H']).font = font_bold
    ws_detail.cell(row=cur_row, column=9).number_format = COUNT_FORMAT
    
    ws_detail.cell(row=cur_row, column=10, value=p_cnt['I']).font = font_bold
    ws_detail.cell(row=cur_row, column=10).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=11, value=p_cnt['J']).font = font_bold
    ws_detail.cell(row=cur_row, column=11).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=12, value=p_rt['K']).font = font_bold
    ws_detail.cell(row=cur_row, column=12).number_format = RATE_FORMAT
    
    ws_detail.cell(row=cur_row, column=13, value=p_cnt['L']).font = font_bold
    ws_detail.cell(row=cur_row, column=13).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=14, value=p_cnt['M']).font = font_bold
    ws_detail.cell(row=cur_row, column=14).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=15, value=p_cnt['N']).font = font_bold
    ws_detail.cell(row=cur_row, column=15).number_format = COUNT_FORMAT
    
    ws_detail.cell(row=cur_row, column=16, value=p_cnt['O']).font = font_bold
    ws_detail.cell(row=cur_row, column=16).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=17, value=p_cnt['P']).font = font_bold
    ws_detail.cell(row=cur_row, column=17).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=18, value=p_cnt['Q']).font = font_bold
    ws_detail.cell(row=cur_row, column=18).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=19, value=p_cnt['R']).font = font_bold
    ws_detail.cell(row=cur_row, column=19).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=20, value=p_cnt['S']).font = font_bold
    ws_detail.cell(row=cur_row, column=20).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=21, value=p_rt['T']).font = font_bold
    ws_detail.cell(row=cur_row, column=21).number_format = RATE_FORMAT
    ws_detail.cell(row=cur_row, column=22, value=p_rt['U']).font = font_bold
    ws_detail.cell(row=cur_row, column=22).number_format = RATE_FORMAT
    ws_detail.cell(row=cur_row, column=23, value=p_rt['V']).font = font_bold
    ws_detail.cell(row=cur_row, column=23).number_format = RATE_FORMAT

    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_2_m10_test(result, template_path="MẪU CHUẨN - PCGD 2026(ok).xlsx", output_dir="outputs"):
    """
    Xuất kết quả thử nghiệm biểu 2.M10 (Hiệu quả phổ cập giáo dục trung học):
    - Sheet 1: TONG_HOP_THU (Mẫu 2.M10 cấp tỉnh: Hàng 47)
    - Sheet 2: CHI_TIET_92_XA (Bảng chi tiết số liệu từng xã)
    """
    os.makedirs(output_dir, exist_ok=True)
    out_filename = "TONG_HOP_THU_2.M10_PCGD_2026.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)
    
    wb = openpyxl.load_workbook(template_path)
    for s_name in list(wb.sheetnames):
        if s_name != '2.M10':
            del wb[s_name]
            
    ws_th = wb['2.M10']
    ws_th.title = 'TONG_HOP_THU'
    
    total_c = result['total_communes_count']
    commune_names_str = ", ".join(result['commune_names'])
    
    ws_th['C3'] = "TỔNG HỢP THỬ - BIỂU TỔNG HỢP HIỆU QUẢ PHỔ CẬP GIÁO DỤC TRUNG HỌC"
    ws_th['C4'] = f"GHI CHÚ: TỔNG HỢP THỬ NGHIỆM (Số xã thực tế được tính: {total_c} xã - {commune_names_str})"
    font_sub = Font(name='Times New Roman', size=11, italic=True, bold=True, color='002060')
    ws_th['C4'].font = font_sub
    
    agg = result['aggregated_cells']
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    COUNT_FORMAT = '#,##0'
    RATE_FORMAT = '0.00"%"'
    
    count_cols = ['C', 'D', 'E', 'F', 'G', 'I', 'J', 'L', 'M', 'O', 'P', 'Q', 'R', 'S']
    rate_cols = ['H', 'K', 'N', 'T']
    
    # 1. Điền giá trị Hàng 47 cấp tỉnh
    for c in count_cols:
        coord = f"{c}47"
        if coord in agg:
            cell = ws_th[coord]
            cell.value = agg[coord]
            cell.font = font_bold
            cell.number_format = COUNT_FORMAT
            
    for c in rate_cols:
        coord = f"{c}47"
        if coord in agg:
            cell = ws_th[coord]
            cell.value = agg[coord]
            cell.font = font_bold
            cell.number_format = RATE_FORMAT

    ws_th.views.sheetView[0].showGridLines = True

    # ==========================================
    # SHEET 2: CHI TIẾT 92 XÃ (2.M10)
    # ==========================================
    ws_detail = wb.create_sheet(title='CHI_TIET_92_XA')
    ws_detail.views.sheetView[0].showGridLines = True
    
    ws_detail['A1'] = "BẢNG TỔNG HỢP HIỆU QUẢ PCGD TRUNG HỌC TỪNG XÃ - BIỂU 2.M10"
    ws_detail['A1'].font = Font(name='Times New Roman', size=13, bold=True)
    ws_detail['A2'] = f"Thời điểm tổng hợp: Năm 2026 | Quy mô: {total_c} xã/phường"
    
    headers_s2 = [
        'STT', 'Tên Xã / Phường', 'Dòng nguồn',
        'Tổng số 18-21 (C)', 'Phải PC 18-21 (D)',
        'TN THPT (E)', 'TN BT THPT (F)', 'TN THCN, TDN (G)', 'Tỷ lệ có bằng (%) (H)',
        'Lớp 12 năm qua (I)', 'TN THPT năm qua (J)', 'Tỷ lệ % (K)',
        'Cuối cấp GDTX/TCCN (L)', 'TN GDTX/TCCN (M)', 'Tỷ lệ % (N)',
        'Đại học (O)', 'Cao đẳng (P)', 'Trung cấp (Q)',
        'Học nghề >= 3 năm (R)', 'Có bằng nghề (S)', 'Tỷ lệ bằng nghề (%) (T)'
    ]
    for ci, h in enumerate(headers_s2, start=1):
        cell_h = ws_detail.cell(row=4, column=ci, value=h)
        cell_h.font = Font(name='Times New Roman', size=10, bold=True)
        cell_h.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        cell_h.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
    cur_row = 5
    for item in result['commune_summaries']:
        ws_detail.cell(row=cur_row, column=1, value=item['stt'])
        ws_detail.cell(row=cur_row, column=2, value=item['commune_name'])
        ws_detail.cell(row=cur_row, column=3, value=item['row_total'])
        
        c_cnt = item['counts']
        c_rt = item['rates']
        
        # C, D
        ws_detail.cell(row=cur_row, column=4, value=c_cnt['C']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=5, value=c_cnt['D']).number_format = COUNT_FORMAT
        
        # E, F, G, H
        ws_detail.cell(row=cur_row, column=6, value=c_cnt['E']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=7, value=c_cnt['F']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=8, value=c_cnt['G']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=9, value=c_rt['H']).number_format = RATE_FORMAT
        
        # I, J, K
        ws_detail.cell(row=cur_row, column=10, value=c_cnt['I']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=11, value=c_cnt['J']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=12, value=c_rt['K']).number_format = RATE_FORMAT
        
        # L, M, N
        ws_detail.cell(row=cur_row, column=13, value=c_cnt['L']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=14, value=c_cnt['M']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=15, value=c_rt['N']).number_format = RATE_FORMAT
        
        # O, P, Q
        ws_detail.cell(row=cur_row, column=16, value=c_cnt['O']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=17, value=c_cnt['P']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=18, value=c_cnt['Q']).number_format = COUNT_FORMAT
        
        # R, S, T
        ws_detail.cell(row=cur_row, column=19, value=c_cnt['R']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=20, value=c_cnt['S']).number_format = COUNT_FORMAT
        ws_detail.cell(row=cur_row, column=21, value=c_rt['T']).number_format = RATE_FORMAT
        
        cur_row += 1
        
    # Hàng tổng cộng toàn tỉnh ở Sheet 2
    ws_detail.cell(row=cur_row, column=2, value="TỔNG CỘNG TOÀN TỈNH").font = font_bold
    p_cnt = result['province_counts']
    p_rt = result['province_rates']
    
    ws_detail.cell(row=cur_row, column=4, value=p_cnt['C']).font = font_bold
    ws_detail.cell(row=cur_row, column=4).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=5, value=p_cnt['D']).font = font_bold
    ws_detail.cell(row=cur_row, column=5).number_format = COUNT_FORMAT
    
    ws_detail.cell(row=cur_row, column=6, value=p_cnt['E']).font = font_bold
    ws_detail.cell(row=cur_row, column=6).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=7, value=p_cnt['F']).font = font_bold
    ws_detail.cell(row=cur_row, column=7).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=8, value=p_cnt['G']).font = font_bold
    ws_detail.cell(row=cur_row, column=8).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=9, value=p_rt['H']).font = font_bold
    ws_detail.cell(row=cur_row, column=9).number_format = RATE_FORMAT
    
    ws_detail.cell(row=cur_row, column=10, value=p_cnt['I']).font = font_bold
    ws_detail.cell(row=cur_row, column=10).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=11, value=p_cnt['J']).font = font_bold
    ws_detail.cell(row=cur_row, column=11).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=12, value=p_rt['K']).font = font_bold
    ws_detail.cell(row=cur_row, column=12).number_format = RATE_FORMAT
    
    ws_detail.cell(row=cur_row, column=13, value=p_cnt['L']).font = font_bold
    ws_detail.cell(row=cur_row, column=13).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=14, value=p_cnt['M']).font = font_bold
    ws_detail.cell(row=cur_row, column=14).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=15, value=p_rt['N']).font = font_bold
    ws_detail.cell(row=cur_row, column=15).number_format = RATE_FORMAT
    
    ws_detail.cell(row=cur_row, column=16, value=p_cnt['O']).font = font_bold
    ws_detail.cell(row=cur_row, column=16).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=17, value=p_cnt['P']).font = font_bold
    ws_detail.cell(row=cur_row, column=17).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=18, value=p_cnt['Q']).font = font_bold
    ws_detail.cell(row=cur_row, column=18).number_format = COUNT_FORMAT
    
    ws_detail.cell(row=cur_row, column=19, value=p_cnt['R']).font = font_bold
    ws_detail.cell(row=cur_row, column=19).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=20, value=p_cnt['S']).font = font_bold
    ws_detail.cell(row=cur_row, column=20).number_format = COUNT_FORMAT
    ws_detail.cell(row=cur_row, column=21, value=p_rt['T']).font = font_bold
    ws_detail.cell(row=cur_row, column=21).number_format = RATE_FORMAT

    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_3_m1_test(result: Dict[str, Any], template_path: str, out_filepath: str) -> str:
    """
    Xuất kết quả tổng hợp Biểu 3.M1 (Chuẩn XMC Mức 1 - Độ tuổi 15-25) ra Excel.
    Gồm:
    - Sheet 1: TONG_HOP_THU (Template 3.M1 với số liệu toàn tỉnh ở dòng @ Cộng)
    - Sheet 2: CHI_TIET_92_XA (Bảng chi tiết 92 xã)
    - Sheet 3: DOI_CHIEU_SAI_LECH (Đối soát tỷ lệ nếu có lệch)
    """
    COUNT_FORMAT = '#,##0'
    RATE_FORMAT = '0.00"%"'
    
    font_bold = Font(name='Times New Roman', size=11, bold=True)
    font_data = Font(name='Times New Roman', size=11, bold=False)
    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')
    
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name='Times New Roman', size=11, bold=True, color="FFFFFF")
    
    wb = openpyxl.load_workbook(template_path)
    
    # Giữ lại sheet 3.M1, đổi tên thành TONG_HOP_THU hoặc giữ 3.M1
    for sname in list(wb.sheetnames):
        if sname != '3.M1':
            wb.remove(wb[sname])
            
    ws_th = wb['3.M1']
    ws_th.title = 'TONG_HOP_THU'
    
    # 1. Điền thông tin cấp tỉnh vào sheet TONG_HOP_THU
    # Ghi chú tiêu đề
    ws_th['A1'] = "TỔNG HỢP TOÀN TỈNH THÁI NGUYÊN - BIỂU 3.M1: THỐNG KÊ ĐẠT CHUẨN XÓA MÙ CHỮ MỨC ĐỘ 1 (15-25 TUỔI)"
    ws_th['A1'].font = Font(name='Times New Roman', size=13, bold=True, color='1F497D')
    
    # Tìm dòng @ Cộng trong template (thường là dòng 46)
    total_r = None
    for r in range(1, ws_th.max_row + 1):
        v0 = str(ws_th.cell(r, 1).value or '').strip()
        v1 = str(ws_th.cell(r, 2).value or '').strip().lower()
        if v0 == '@' or 'cộng' in v1:
            total_r = r
            break
            
    if not total_r:
        total_r = 46
        
    # Xóa các dòng mẫu xóm ở giữa (từ dòng 7 đến total_r - 1)
    for r in range(7, total_r):
        for c in range(1, 10):
            ws_th.cell(r, c).value = None
    ws_th.cell(7, 2, value="(Số liệu chi tiết 92 xã được tổng hợp tại sheet CHI_TIET_92_XA)").font = Font(name='Times New Roman', size=10, italic=True, color='555555')
    
    # Điền dòng tổng tỉnh
    ws_th.cell(total_r, 1, value="@").font = font_bold
    ws_th.cell(total_r, 2, value="Cộng toàn tỉnh (92 xã)").font = font_bold
    
    c_cell = ws_th.cell(total_r, 3, value=result['sum_c_dan_so'])
    c_cell.font = font_bold; c_cell.number_format = COUNT_FORMAT; c_cell.alignment = align_right
    
    d_cell = ws_th.cell(total_r, 4, value=result['sum_d_biet_chu_m1'])
    d_cell.font = font_bold; d_cell.number_format = COUNT_FORMAT; d_cell.alignment = align_right
    
    e_cell = ws_th.cell(total_r, 5, value=result['rate_m1_prov'])
    e_cell.font = font_bold; e_cell.number_format = RATE_FORMAT; e_cell.alignment = align_right
    
    f_cell = ws_th.cell(total_r, 6, value=result['sum_f_biet_chu_m2'])
    f_cell.font = font_bold; f_cell.number_format = COUNT_FORMAT; f_cell.alignment = align_right
    
    g_cell = ws_th.cell(total_r, 7, value=result['rate_m2_prov'])
    g_cell.font = font_bold; g_cell.number_format = RATE_FORMAT; g_cell.alignment = align_right
    
    # Các dòng đạt chuẩn
    ws_th.cell(total_r + 1, 1, value="Số xã, phường đạt chuẩn XMC mức độ 1:").font = font_bold
    ws_th.cell(total_r + 1, 4, value=f"{result['count_dat_m1']}/{result['total_communes']} xã").font = font_bold
    ws_th.cell(total_r + 1, 4).alignment = align_center
    
    ws_th.cell(total_r + 2, 1, value="Số xã, phường đạt chuẩn XMC mức độ 2:").font = font_bold
    ws_th.cell(total_r + 2, 4, value=f"{result['count_dat_m2']}/{result['total_communes']} xã").font = font_bold
    ws_th.cell(total_r + 2, 4).alignment = align_center
    
    # 2. Sheet 2: CHI_TIET_92_XA
    ws_detail = wb.create_sheet(title='CHI_TIET_92_XA')
    headers_dt = [
        "STT", "Tên Xã / Phường", "Tổng số dân số 15-25t (C)", 
        "Biết chữ Mức 1 (D)", "Tỷ lệ M1 nguồn (%)", "Tỷ lệ M1 chuẩn (%) (E)",
        "Biết chữ Mức 2 (F)", "Tỷ lệ M2 nguồn (%)", "Tỷ lệ M2 chuẩn (%) (G)",
        "Đạt chuẩn M1", "Đạt chuẩn M2"
    ]
    ws_detail.append(headers_dt)
    for col_idx in range(1, len(headers_dt) + 1):
        cell = ws_detail.cell(1, col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        
    for idx, c_data in enumerate(result['communes_detail'], start=1):
        row_vals = [
            idx,
            c_data['commune_name'],
            c_data['c_dan_so'],
            c_data['d_biet_chu_m1'],
            c_data['e_rate_m1_src'],
            c_data['e_rate_m1_calc'],
            c_data['f_biet_chu_m2'],
            c_data['g_rate_m2_src'],
            c_data['g_rate_m2_calc'],
            "Đạt" if c_data['dat_m1'] else "Chưa đạt",
            "Đạt" if c_data['dat_m2'] else "Chưa đạt"
        ]
        ws_detail.append(row_vals)
        r_idx = idx + 1
        
        ws_detail.cell(r_idx, 1).alignment = align_center
        ws_detail.cell(r_idx, 2).alignment = align_left
        
        ws_detail.cell(r_idx, 3).number_format = COUNT_FORMAT; ws_detail.cell(r_idx, 3).alignment = align_right
        ws_detail.cell(r_idx, 4).number_format = COUNT_FORMAT; ws_detail.cell(r_idx, 4).alignment = align_right
        ws_detail.cell(r_idx, 5).number_format = RATE_FORMAT; ws_detail.cell(r_idx, 5).alignment = align_right
        ws_detail.cell(r_idx, 6).number_format = RATE_FORMAT; ws_detail.cell(r_idx, 6).alignment = align_right
        ws_detail.cell(r_idx, 7).number_format = COUNT_FORMAT; ws_detail.cell(r_idx, 7).alignment = align_right
        ws_detail.cell(r_idx, 8).number_format = RATE_FORMAT; ws_detail.cell(r_idx, 8).alignment = align_right
        ws_detail.cell(r_idx, 9).number_format = RATE_FORMAT; ws_detail.cell(r_idx, 9).alignment = align_right
        ws_detail.cell(r_idx, 10).alignment = align_center
        ws_detail.cell(r_idx, 11).alignment = align_center
        
    # Hàng tổng cộng ở Sheet 2
    tot_row = len(result['communes_detail']) + 2
    ws_detail.cell(tot_row, 1, value="").font = font_bold
    ws_detail.cell(tot_row, 2, value="TỔNG CỘNG TOÀN TỈNH").font = font_bold
    
    ws_detail.cell(tot_row, 3, value=result['sum_c_dan_so']).font = font_bold
    ws_detail.cell(tot_row, 3).number_format = COUNT_FORMAT; ws_detail.cell(tot_row, 3).alignment = align_right
    
    ws_detail.cell(tot_row, 4, value=result['sum_d_biet_chu_m1']).font = font_bold
    ws_detail.cell(tot_row, 4).number_format = COUNT_FORMAT; ws_detail.cell(tot_row, 4).alignment = align_right
    
    ws_detail.cell(tot_row, 5, value="").font = font_bold
    ws_detail.cell(tot_row, 6, value=result['rate_m1_prov']).font = font_bold
    ws_detail.cell(tot_row, 6).number_format = RATE_FORMAT; ws_detail.cell(tot_row, 6).alignment = align_right
    
    ws_detail.cell(tot_row, 7, value=result['sum_f_biet_chu_m2']).font = font_bold
    ws_detail.cell(tot_row, 7).number_format = COUNT_FORMAT; ws_detail.cell(tot_row, 7).alignment = align_right
    
    ws_detail.cell(tot_row, 8, value="").font = font_bold
    ws_detail.cell(tot_row, 9, value=result['rate_m2_prov']).font = font_bold
    ws_detail.cell(tot_row, 9).number_format = RATE_FORMAT; ws_detail.cell(tot_row, 9).alignment = align_right
    
    ws_detail.cell(tot_row, 10, value=f"{result['count_dat_m1']}/{result['total_communes']}").font = font_bold
    ws_detail.cell(tot_row, 10).alignment = align_center
    ws_detail.cell(tot_row, 11, value=f"{result['count_dat_m2']}/{result['total_communes']}").font = font_bold
    ws_detail.cell(tot_row, 11).alignment = align_center
    
    # 3. Sheet 3: DOI_CHIEU_SAI_LECH (nếu có)
    if result.get('discrepancies'):
        ws_dc = wb.create_sheet(title='DOI_CHIEU_SAI_LECH')
        headers_dc = ["STT", "Tên Xã / Phường", "Tỷ lệ M1 nguồn", "Tỷ lệ M1 chuẩn", "Tỷ lệ M2 nguồn", "Tỷ lệ M2 chuẩn", "Ghi chú"]
        ws_dc.append(headers_dc)
        fill_dc = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
        for col_idx in range(1, len(headers_dc) + 1):
            cell = ws_dc.cell(1, col_idx)
            cell.fill = fill_dc
            cell.font = header_font
            cell.alignment = align_center
            
        for d_idx, d in enumerate(result['discrepancies'], start=1):
            ws_dc.append([
                d_idx, d['commune_name'], 
                d['e_rate_m1_src'], d['e_rate_m1_calc'],
                d['g_rate_m2_src'], d['g_rate_m2_calc'],
                d['note']
            ])
            r_idx = d_idx + 1
            ws_dc.cell(r_idx, 3).number_format = RATE_FORMAT
            ws_dc.cell(r_idx, 4).number_format = RATE_FORMAT
            ws_dc.cell(r_idx, 5).number_format = RATE_FORMAT
            ws_dc.cell(r_idx, 6).number_format = RATE_FORMAT
            
    # Hiển thị đường lưới cho tất cả sheets
    for ws in wb.worksheets:
        ws.views.sheetView[0].showGridLines = True
        
    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_3_m2_test(result: Dict[str, Any], template_path: str, out_filepath: str) -> str:
    """
    Xuất kết quả tổng hợp Biểu 3.M2 (Chuẩn XMC Mức 2 - Độ tuổi 15-35) ra Excel.
    Gồm:
    - Sheet 1: TONG_HOP_THU (Template 3.M2 với số liệu toàn tỉnh ở dòng @ Cộng)
    - Sheet 2: CHI_TIET_92_XA (Bảng chi tiết 92 xã)
    - Sheet 3: DOI_CHIEU_SAI_LECH (Đối soát tỷ lệ nếu có lệch)
    """
    COUNT_FORMAT = '#,##0'
    RATE_FORMAT = '0.00"%"'
    
    font_bold = Font(name='Times New Roman', size=11, bold=True)
    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')
    
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name='Times New Roman', size=11, bold=True, color="FFFFFF")
    
    wb = openpyxl.load_workbook(template_path)
    
    # Giữ lại sheet 3.M2, đổi tên thành TONG_HOP_THU
    for sname in list(wb.sheetnames):
        if sname != '3.M2':
            wb.remove(wb[sname])
            
    ws_th = wb['3.M2']
    ws_th.title = 'TONG_HOP_THU'
    
    # 1. Điền thông tin cấp tỉnh vào sheet TONG_HOP_THU
    ws_th['A1'] = "TỔNG HỢP TOÀN TỈNH THÁI NGUYÊN - BIỂU 3.M2: THỐNG KÊ ĐẠT CHUẨN XÓA MÙ CHỮ MỨC ĐỘ 2 (15-35 TUỔI)"
    ws_th['A1'].font = Font(name='Times New Roman', size=13, bold=True, color='1F497D')
    
    # Tìm dòng @ Cộng trong template (thường là dòng 47 trong 3.M2)
    total_r = None
    for r in range(1, ws_th.max_row + 1):
        v0 = str(ws_th.cell(r, 1).value or '').strip()
        v1 = str(ws_th.cell(r, 2).value or '').strip().lower()
        if v0 == '@' or 'cộng' in v1:
            total_r = r
            break
            
    if not total_r:
        total_r = 47
        
    # Xóa các dòng mẫu xóm ở giữa (từ dòng 8 đến total_r - 1)
    for r in range(8, total_r):
        for c in range(1, 10):
            ws_th.cell(r, c).value = None
    ws_th.cell(8, 2, value="(Số liệu chi tiết 92 xã được tổng hợp tại sheet CHI_TIET_92_XA)").font = Font(name='Times New Roman', size=10, italic=True, color='555555')
    
    # Điền dòng tổng tỉnh
    ws_th.cell(total_r, 1, value="@").font = font_bold
    ws_th.cell(total_r, 2, value="Cộng toàn tỉnh (92 xã)").font = font_bold
    
    c_cell = ws_th.cell(total_r, 3, value=result['sum_c_dan_so'])
    c_cell.font = font_bold; c_cell.number_format = COUNT_FORMAT; c_cell.alignment = align_right
    
    d_cell = ws_th.cell(total_r, 4, value=result['sum_d_biet_chu_m1'])
    d_cell.font = font_bold; d_cell.number_format = COUNT_FORMAT; d_cell.alignment = align_right
    
    e_cell = ws_th.cell(total_r, 5, value=result['rate_m1_prov'])
    e_cell.font = font_bold; e_cell.number_format = RATE_FORMAT; e_cell.alignment = align_right
    
    f_cell = ws_th.cell(total_r, 6, value=result['sum_f_biet_chu_m2'])
    f_cell.font = font_bold; f_cell.number_format = COUNT_FORMAT; f_cell.alignment = align_right
    
    g_cell = ws_th.cell(total_r, 7, value=result['rate_m2_prov'])
    g_cell.font = font_bold; g_cell.number_format = RATE_FORMAT; g_cell.alignment = align_right
    
    # Các dòng đạt chuẩn
    ws_th.cell(total_r + 1, 1, value="Số xã, phường đạt chuẩn XMC mức độ 1:").font = font_bold
    ws_th.cell(total_r + 1, 4, value=f"{result['count_dat_m1']}/{result['total_communes']} xã").font = font_bold
    ws_th.cell(total_r + 1, 4).alignment = align_center
    
    ws_th.cell(total_r + 2, 1, value="Số xã, phường đạt chuẩn XMC mức độ 2:").font = font_bold
    ws_th.cell(total_r + 2, 4, value=f"{result['count_dat_m2']}/{result['total_communes']} xã").font = font_bold
    ws_th.cell(total_r + 2, 4).alignment = align_center
    
    # 2. Sheet 2: CHI_TIET_92_XA
    ws_detail = wb.create_sheet(title='CHI_TIET_92_XA')
    headers_dt = [
        "STT", "Tên Xã / Phường", "Tổng số dân số 15-35t (C)", 
        "Biết chữ Mức 1 (D)", "Tỷ lệ M1 nguồn (%)", "Tỷ lệ M1 chuẩn (%) (E)",
        "Biết chữ Mức 2 (F)", "Tỷ lệ M2 nguồn (%)", "Tỷ lệ M2 chuẩn (%) (G)",
        "Đạt chuẩn M1", "Đạt chuẩn M2"
    ]
    ws_detail.append(headers_dt)
    for col_idx in range(1, len(headers_dt) + 1):
        cell = ws_detail.cell(1, col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        
    for idx, c_data in enumerate(result['communes_detail'], start=1):
        row_vals = [
            idx,
            c_data['commune_name'],
            c_data['c_dan_so'],
            c_data['d_biet_chu_m1'],
            c_data['e_rate_m1_src'],
            c_data['e_rate_m1_calc'],
            c_data['f_biet_chu_m2'],
            c_data['g_rate_m2_src'],
            c_data['g_rate_m2_calc'],
            "Đạt" if c_data['dat_m1'] else "Chưa đạt",
            "Đạt" if c_data['dat_m2'] else "Chưa đạt"
        ]
        ws_detail.append(row_vals)
        r_idx = idx + 1
        
        ws_detail.cell(r_idx, 1).alignment = align_center
        ws_detail.cell(r_idx, 2).alignment = align_left
        
        ws_detail.cell(r_idx, 3).number_format = COUNT_FORMAT; ws_detail.cell(r_idx, 3).alignment = align_right
        ws_detail.cell(r_idx, 4).number_format = COUNT_FORMAT; ws_detail.cell(r_idx, 4).alignment = align_right
        ws_detail.cell(r_idx, 5).number_format = RATE_FORMAT; ws_detail.cell(r_idx, 5).alignment = align_right
        ws_detail.cell(r_idx, 6).number_format = RATE_FORMAT; ws_detail.cell(r_idx, 6).alignment = align_right
        ws_detail.cell(r_idx, 7).number_format = COUNT_FORMAT; ws_detail.cell(r_idx, 7).alignment = align_right
        ws_detail.cell(r_idx, 8).number_format = RATE_FORMAT; ws_detail.cell(r_idx, 8).alignment = align_right
        ws_detail.cell(r_idx, 9).number_format = RATE_FORMAT; ws_detail.cell(r_idx, 9).alignment = align_right
        ws_detail.cell(r_idx, 10).alignment = align_center
        ws_detail.cell(r_idx, 11).alignment = align_center
        
    # Hàng tổng cộng ở Sheet 2
    tot_row = len(result['communes_detail']) + 2
    ws_detail.cell(tot_row, 1, value="").font = font_bold
    ws_detail.cell(tot_row, 2, value="TỔNG CỘNG TOÀN TỈNH").font = font_bold
    
    ws_detail.cell(tot_row, 3, value=result['sum_c_dan_so']).font = font_bold
    ws_detail.cell(tot_row, 3).number_format = COUNT_FORMAT; ws_detail.cell(tot_row, 3).alignment = align_right
    
    ws_detail.cell(tot_row, 4, value=result['sum_d_biet_chu_m1']).font = font_bold
    ws_detail.cell(tot_row, 4).number_format = COUNT_FORMAT; ws_detail.cell(tot_row, 4).alignment = align_right
    
    ws_detail.cell(tot_row, 5, value="").font = font_bold
    ws_detail.cell(tot_row, 6, value=result['rate_m1_prov']).font = font_bold
    ws_detail.cell(tot_row, 6).number_format = RATE_FORMAT; ws_detail.cell(tot_row, 6).alignment = align_right
    
    ws_detail.cell(tot_row, 7, value=result['sum_f_biet_chu_m2']).font = font_bold
    ws_detail.cell(tot_row, 7).number_format = COUNT_FORMAT; ws_detail.cell(tot_row, 7).alignment = align_right
    
    ws_detail.cell(tot_row, 8, value="").font = font_bold
    ws_detail.cell(tot_row, 9, value=result['rate_m2_prov']).font = font_bold
    ws_detail.cell(tot_row, 9).number_format = RATE_FORMAT; ws_detail.cell(tot_row, 9).alignment = align_right
    
    ws_detail.cell(tot_row, 10, value=f"{result['count_dat_m1']}/{result['total_communes']}").font = font_bold
    ws_detail.cell(tot_row, 10).alignment = align_center
    ws_detail.cell(tot_row, 11, value=f"{result['count_dat_m2']}/{result['total_communes']}").font = font_bold
    ws_detail.cell(tot_row, 11).alignment = align_center
    
    # 3. Sheet 3: DOI_CHIEU_SAI_LECH (nếu có)
    if result.get('discrepancies'):
        ws_dc = wb.create_sheet(title='DOI_CHIEU_SAI_LECH')
        headers_dc = ["STT", "Tên Xã / Phường", "Tỷ lệ M1 nguồn", "Tỷ lệ M1 chuẩn", "Tỷ lệ M2 nguồn", "Tỷ lệ M2 chuẩn", "Ghi chú"]
        ws_dc.append(headers_dc)
        fill_dc = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
        for col_idx in range(1, len(headers_dc) + 1):
            cell = ws_dc.cell(1, col_idx)
            cell.fill = fill_dc
            cell.font = header_font
            cell.alignment = align_center
            
        for d_idx, d in enumerate(result['discrepancies'], start=1):
            ws_dc.append([
                d_idx, d['commune_name'], 
                d['e_rate_m1_src'], d['e_rate_m1_calc'],
                d['g_rate_m2_src'], d['g_rate_m2_calc'],
                d['note']
            ])
            r_idx = d_idx + 1
            ws_dc.cell(r_idx, 3).number_format = RATE_FORMAT
            ws_dc.cell(r_idx, 4).number_format = RATE_FORMAT
            ws_dc.cell(r_idx, 5).number_format = RATE_FORMAT
            ws_dc.cell(r_idx, 6).number_format = RATE_FORMAT
            
    # Hiển thị đường lưới cho tất cả sheets
    for ws in wb.worksheets:
        ws.views.sheetView[0].showGridLines = True
        
    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_3_m3_test(result: Dict[str, Any], template_path: str, out_filepath: str) -> str:
    """
    Xuất kết quả tổng hợp Biểu 3.M3 (Chuẩn XMC Chung - Độ tuổi 15-60) ra Excel.
    Gồm:
    - Sheet 1: TONG_HOP_THU (Template 3.M3 với số liệu toàn tỉnh ở dòng @ Cộng)
    - Sheet 2: CHI_TIET_92_XA (Bảng chi tiết 92 xã)
    - Sheet 3: DOI_CHIEU_SAI_LECH (Đối soát tỷ lệ nếu có lệch)
    """
    COUNT_FORMAT = '#,##0'
    RATE_FORMAT = '0.00"%"'
    
    font_bold = Font(name='Times New Roman', size=11, bold=True)
    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')
    
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name='Times New Roman', size=11, bold=True, color="FFFFFF")
    
    wb = openpyxl.load_workbook(template_path)
    
    # Giữ lại sheet 3.M3, đổi tên thành TONG_HOP_THU
    for sname in list(wb.sheetnames):
        if sname != '3.M3':
            wb.remove(wb[sname])
            
    ws_th = wb['3.M3']
    ws_th.title = 'TONG_HOP_THU'
    
    # 1. Điền thông tin cấp tỉnh vào sheet TONG_HOP_THU
    ws_th['A1'] = "TỔNG HỢP TOÀN TỈNH THÁI NGUYÊN - BIỂU 3.M3: THỐNG KÊ ĐẠT CHUẨN XÓA MÙ CHỮ TOÀN XÃ (15-60 TUỔI)"
    ws_th['A1'].font = Font(name='Times New Roman', size=13, bold=True, color='1F497D')
    
    # Tìm dòng @ Cộng trong template (thường là dòng 47 trong 3.M3)
    total_r = None
    for r in range(1, ws_th.max_row + 1):
        v0 = str(ws_th.cell(r, 1).value or '').strip()
        v1 = str(ws_th.cell(r, 2).value or '').strip().lower()
        if v0 == '@' or 'cộng' in v1:
            total_r = r
            break
            
    if not total_r:
        total_r = 47
        
    # Xóa các dòng mẫu xóm ở giữa (từ dòng 8 đến total_r - 1)
    for r in range(8, total_r):
        for c in range(1, 10):
            ws_th.cell(r, c).value = None
    ws_th.cell(8, 2, value="(Số liệu chi tiết 92 xã được tổng hợp tại sheet CHI_TIET_92_XA)").font = Font(name='Times New Roman', size=10, italic=True, color='555555')
    
    # Điền dòng tổng tỉnh
    ws_th.cell(total_r, 1, value="@").font = font_bold
    ws_th.cell(total_r, 2, value="Cộng toàn tỉnh (92 xã)").font = font_bold
    
    c_cell = ws_th.cell(total_r, 3, value=result['sum_c_dan_so'])
    c_cell.font = font_bold; c_cell.number_format = COUNT_FORMAT; c_cell.alignment = align_right
    
    d_cell = ws_th.cell(total_r, 4, value=result['sum_d_biet_chu_m1'])
    d_cell.font = font_bold; d_cell.number_format = COUNT_FORMAT; d_cell.alignment = align_right
    
    e_cell = ws_th.cell(total_r, 5, value=result['rate_m1_prov'])
    e_cell.font = font_bold; e_cell.number_format = RATE_FORMAT; e_cell.alignment = align_right
    
    f_cell = ws_th.cell(total_r, 6, value=result['sum_f_biet_chu_m2'])
    f_cell.font = font_bold; f_cell.number_format = COUNT_FORMAT; f_cell.alignment = align_right
    
    g_cell = ws_th.cell(total_r, 7, value=result['rate_m2_prov'])
    g_cell.font = font_bold; g_cell.number_format = RATE_FORMAT; g_cell.alignment = align_right
    
    # Các dòng đạt chuẩn
    ws_th.cell(total_r + 1, 1, value="Số xã, phường đạt chuẩn XMC mức độ 1:").font = font_bold
    ws_th.cell(total_r + 1, 4, value=f"{result['count_dat_m1']}/{result['total_communes']} xã").font = font_bold
    ws_th.cell(total_r + 1, 4).alignment = align_center
    
    ws_th.cell(total_r + 2, 1, value="Số xã, phường đạt chuẩn XMC mức độ 2:").font = font_bold
    ws_th.cell(total_r + 2, 4, value=f"{result['count_dat_m2']}/{result['total_communes']} xã").font = font_bold
    ws_th.cell(total_r + 2, 4).alignment = align_center
    
    # 2. Sheet 2: CHI_TIET_92_XA
    ws_detail = wb.create_sheet(title='CHI_TIET_92_XA')
    headers_dt = [
        "STT", "Tên Xã / Phường", "Tổng số dân số 15-60t (C)", 
        "Biết chữ Mức 1 (D)", "Tỷ lệ M1 nguồn (%)", "Tỷ lệ M1 chuẩn (%) (E)",
        "Biết chữ Mức 2 (F)", "Tỷ lệ M2 nguồn (%)", "Tỷ lệ M2 chuẩn (%) (G)",
        "Đạt chuẩn M1", "Đạt chuẩn M2"
    ]
    ws_detail.append(headers_dt)
    for col_idx in range(1, len(headers_dt) + 1):
        cell = ws_detail.cell(1, col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        
    for idx, c_data in enumerate(result['communes_detail'], start=1):
        row_vals = [
            idx,
            c_data['commune_name'],
            c_data['c_dan_so'],
            c_data['d_biet_chu_m1'],
            c_data['e_rate_m1_src'],
            c_data['e_rate_m1_calc'],
            c_data['f_biet_chu_m2'],
            c_data['g_rate_m2_src'],
            c_data['g_rate_m2_calc'],
            "Đạt" if c_data['dat_m1'] else "Chưa đạt",
            "Đạt" if c_data['dat_m2'] else "Chưa đạt"
        ]
        ws_detail.append(row_vals)
        r_idx = idx + 1
        
        ws_detail.cell(r_idx, 1).alignment = align_center
        ws_detail.cell(r_idx, 2).alignment = align_left
        
        ws_detail.cell(r_idx, 3).number_format = COUNT_FORMAT; ws_detail.cell(r_idx, 3).alignment = align_right
        ws_detail.cell(r_idx, 4).number_format = COUNT_FORMAT; ws_detail.cell(r_idx, 4).alignment = align_right
        ws_detail.cell(r_idx, 5).number_format = RATE_FORMAT; ws_detail.cell(r_idx, 5).alignment = align_right
        ws_detail.cell(r_idx, 6).number_format = RATE_FORMAT; ws_detail.cell(r_idx, 6).alignment = align_right
        ws_detail.cell(r_idx, 7).number_format = COUNT_FORMAT; ws_detail.cell(r_idx, 7).alignment = align_right
        ws_detail.cell(r_idx, 8).number_format = RATE_FORMAT; ws_detail.cell(r_idx, 8).alignment = align_right
        ws_detail.cell(r_idx, 9).number_format = RATE_FORMAT; ws_detail.cell(r_idx, 9).alignment = align_right
        ws_detail.cell(r_idx, 10).alignment = align_center
        ws_detail.cell(r_idx, 11).alignment = align_center
        
    # Hàng tổng cộng ở Sheet 2
    tot_row = len(result['communes_detail']) + 2
    ws_detail.cell(tot_row, 1, value="").font = font_bold
    ws_detail.cell(tot_row, 2, value="TỔNG CỘNG TOÀN TỈNH").font = font_bold
    
    ws_detail.cell(tot_row, 3, value=result['sum_c_dan_so']).font = font_bold
    ws_detail.cell(tot_row, 3).number_format = COUNT_FORMAT; ws_detail.cell(tot_row, 3).alignment = align_right
    
    ws_detail.cell(tot_row, 4, value=result['sum_d_biet_chu_m1']).font = font_bold
    ws_detail.cell(tot_row, 4).number_format = COUNT_FORMAT; ws_detail.cell(tot_row, 4).alignment = align_right
    
    ws_detail.cell(tot_row, 5, value="").font = font_bold
    ws_detail.cell(tot_row, 6, value=result['rate_m1_prov']).font = font_bold
    ws_detail.cell(tot_row, 6).number_format = RATE_FORMAT; ws_detail.cell(tot_row, 6).alignment = align_right
    
    ws_detail.cell(tot_row, 7, value=result['sum_f_biet_chu_m2']).font = font_bold
    ws_detail.cell(tot_row, 7).number_format = COUNT_FORMAT; ws_detail.cell(tot_row, 7).alignment = align_right
    
    ws_detail.cell(tot_row, 8, value="").font = font_bold
    ws_detail.cell(tot_row, 9, value=result['rate_m2_prov']).font = font_bold
    ws_detail.cell(tot_row, 9).number_format = RATE_FORMAT; ws_detail.cell(tot_row, 9).alignment = align_right
    
    ws_detail.cell(tot_row, 10, value=f"{result['count_dat_m1']}/{result['total_communes']}").font = font_bold
    ws_detail.cell(tot_row, 10).alignment = align_center
    ws_detail.cell(tot_row, 11, value=f"{result['count_dat_m2']}/{result['total_communes']}").font = font_bold
    ws_detail.cell(tot_row, 11).alignment = align_center
    
    # 3. Sheet 3: DOI_CHIEU_SAI_LECH (nếu có)
    if result.get('discrepancies'):
        ws_dc = wb.create_sheet(title='DOI_CHIEU_SAI_LECH')
        headers_dc = ["STT", "Tên Xã / Phường", "Tỷ lệ M1 nguồn", "Tỷ lệ M1 chuẩn", "Tỷ lệ M2 nguồn", "Tỷ lệ M2 chuẩn", "Ghi chú"]
        ws_dc.append(headers_dc)
        fill_dc = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
        for col_idx in range(1, len(headers_dc) + 1):
            cell = ws_dc.cell(1, col_idx)
            cell.fill = fill_dc
            cell.font = header_font
            cell.alignment = align_center
            
        for d_idx, d in enumerate(result['discrepancies'], start=1):
            ws_dc.append([
                d_idx, d['commune_name'], 
                d['e_rate_m1_src'], d['e_rate_m1_calc'],
                d['g_rate_m2_src'], d['g_rate_m2_calc'],
                d['note']
            ])
            r_idx = d_idx + 1
            ws_dc.cell(r_idx, 3).number_format = RATE_FORMAT
            ws_dc.cell(r_idx, 4).number_format = RATE_FORMAT
            ws_dc.cell(r_idx, 5).number_format = RATE_FORMAT
            ws_dc.cell(r_idx, 6).number_format = RATE_FORMAT
            
    # Hiển thị đường lưới cho tất cả sheets
    for ws in wb.worksheets:
        ws.views.sheetView[0].showGridLines = True
        
    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_3_mu_chu_test(result: Dict[str, Any], template_path: str, out_filepath: str) -> str:
    """
    Xuất kết quả tổng hợp cho các Biểu Hiện trạng mù chữ (3.M4 đến 3.M9) ra Excel.
    """
    COUNT_FORMAT = '#,##0'
    RATE_FORMAT = '0.00"%"'
    
    sheet_code = result['sheet_code']
    font_bold = Font(name='Times New Roman', size=11, bold=True)
    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')
    
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name='Times New Roman', size=11, bold=True, color="FFFFFF")
    
    wb = openpyxl.load_workbook(template_path)
    
    for sname in list(wb.sheetnames):
        if sname != sheet_code:
            wb.remove(wb[sname])
            
    ws_th = wb[sheet_code]
    ws_th.title = 'TONG_HOP_THU'
    
    ws_th['A1'] = f"TỔNG HỢP TOÀN TỈNH THÁI NGUYÊN - BIỂU {sheet_code}: THỐNG KÊ HIỆN TRẠNG MÙ CHỮ {result['muc_do'].upper()} ({result['age_label'].upper()})"
    ws_th['A1'].font = Font(name='Times New Roman', size=13, bold=True, color='1F497D')
    
    total_r = result['def_sum_row']
    # Tìm chính xác dòng tổng trong template nếu có
    for r in range(1, ws_th.max_row + 1):
        v0 = str(ws_th.cell(r, 1).value or '').strip()
        v1 = str(ws_th.cell(r, 2).value or '').strip().lower()
        if v0 == '@' or any(k in v1 for k in ['tổng', 'cộng']):
            total_r = r
            break
            
    # Xóa các dòng xóm mẫu
    start_r = 9 if sheet_code in ['3.M4', '3.M7'] else 8
    for r in range(start_r, total_r):
        for c in range(1, 16):
            ws_th.cell(r, c).value = None
    ws_th.cell(start_r, 2, value="(Số liệu chi tiết 92 xã được tổng hợp tại sheet CHI_TIET_92_XA)").font = Font(name='Times New Roman', size=10, italic=True, color='555555')
    
    # Điền dòng tổng tỉnh
    ws_th.cell(total_r, 1, value="@").font = font_bold
    ws_th.cell(total_r, 2, value="Cộng toàn tỉnh (92 xã)").font = font_bold
    
    cols_vals = [
        (3, result['sum_c_dan_so'], COUNT_FORMAT),
        (4, result['sum_d_nu'], COUNT_FORMAT),
        (5, result['sum_e_dan_toc'], COUNT_FORMAT),
        (6, result['sum_f_nu_dan_toc'], COUNT_FORMAT),
        (7, result['sum_g_mu_chu'], COUNT_FORMAT),
        (8, result['rate_h_prov'], RATE_FORMAT),
        (9, result['sum_i_mu_chu_nu'], COUNT_FORMAT),
        (10, result['rate_j_prov'], RATE_FORMAT),
        (11, result['sum_k_mu_chu_dt'], COUNT_FORMAT),
        (12, result['rate_l_prov'], RATE_FORMAT),
        (13, result['sum_m_mu_chu_nu_dt'], COUNT_FORMAT),
        (14, result['rate_n_prov'], RATE_FORMAT),
    ]
    for c_idx, val, fmt in cols_vals:
        cell = ws_th.cell(total_r, c_idx, value=val)
        cell.font = font_bold
        cell.number_format = fmt
        cell.alignment = align_right
        
    # Sheet 2: CHI_TIET_92_XA
    ws_detail = wb.create_sheet(title='CHI_TIET_92_XA')
    headers_dt = [
        "STT", "Tên Xã / Phường", "Tổng DS", "Nữ", "Dân tộc", "Nữ DT",
        "Mù chữ (G)", "Tỷ lệ (%) (H)", "Mù chữ Nữ (I)", "Tỷ lệ Nữ (%) (J)",
        "Mù chữ DT (K)", "Tỷ lệ DT (%) (L)", "Mù chữ Nữ DT (M)", "Tỷ lệ Nữ DT (%) (N)"
    ]
    ws_detail.append(headers_dt)
    for col_idx in range(1, len(headers_dt) + 1):
        cell = ws_detail.cell(1, col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        
    for idx, c_data in enumerate(result['communes_detail'], start=1):
        row_vals = [
            idx,
            c_data['commune_name'],
            c_data['c_dan_so'],
            c_data['d_nu'],
            c_data['e_dan_toc'],
            c_data['f_nu_dan_toc'],
            c_data['g_mu_chu'],
            c_data['h_rate_calc'],
            c_data['i_mu_chu_nu'],
            c_data['j_rate_nu_calc'],
            c_data['k_mu_chu_dt'],
            c_data['l_rate_dt_calc'],
            c_data['m_mu_chu_nu_dt'],
            c_data['n_rate_nu_dt_calc']
        ]
        ws_detail.append(row_vals)
        r_idx = idx + 1
        ws_detail.cell(r_idx, 1).alignment = align_center
        ws_detail.cell(r_idx, 2).alignment = align_left
        for c in range(3, 15):
            cell = ws_detail.cell(r_idx, c)
            cell.alignment = align_right
            if c in [8, 10, 12, 14]:
                cell.number_format = RATE_FORMAT
            else:
                cell.number_format = COUNT_FORMAT
                
    # Dòng tổng tỉnh ở Sheet 2
    tot_row = len(result['communes_detail']) + 2
    ws_detail.cell(tot_row, 1, value="").font = font_bold
    ws_detail.cell(tot_row, 2, value="TỔNG CỘNG TOÀN TỈNH").font = font_bold
    for c_idx, val, fmt in cols_vals:
        cell = ws_detail.cell(tot_row, c_idx, value=val)
        cell.font = font_bold
        cell.number_format = fmt
        cell.alignment = align_right
        
    # Sheet 3: DOI_CHIEU_SAI_LECH (nếu có)
    if result.get('discrepancies'):
        ws_dc = wb.create_sheet(title='DOI_CHIEU_SAI_LECH')
        headers_dc = ["STT", "Tên Xã / Phường", "Tỷ lệ nguồn", "Tỷ lệ tính lại", "Ghi chú"]
        ws_dc.append(headers_dc)
        fill_dc = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
        for col_idx in range(1, len(headers_dc) + 1):
            cell = ws_dc.cell(1, col_idx)
            cell.fill = fill_dc
            cell.font = header_font
            cell.alignment = align_center
            
        for d_idx, d in enumerate(result['discrepancies'], start=1):
            ws_dc.append([d_idx, d['commune_name'], d['h_rate_src'], d['h_rate_calc'], d['note']])
            r_idx = d_idx + 1
            ws_dc.cell(r_idx, 3).number_format = RATE_FORMAT
            ws_dc.cell(r_idx, 4).number_format = RATE_FORMAT
            
    for ws in wb.worksheets:
        ws.views.sheetView[0].showGridLines = True
        
    try:
        wb.save(out_filepath)
    except PermissionError:
        base, ext = os.path.splitext(out_filepath)
        alt_filepath = f"{base}_capnhat{ext}"
        wb.save(alt_filepath)
        out_filepath = alt_filepath
        
    wb.close()
    return out_filepath


def export_3_m4_test(result: Dict[str, Any], template_path: str, out_filepath: str) -> str:
    return export_3_mu_chu_test(result, template_path, out_filepath)

def export_3_m5_test(result: Dict[str, Any], template_path: str, out_filepath: str) -> str:
    return export_3_mu_chu_test(result, template_path, out_filepath)

def export_3_m6_test(result: Dict[str, Any], template_path: str, out_filepath: str) -> str:
    return export_3_mu_chu_test(result, template_path, out_filepath)

def export_3_m7_test(result: Dict[str, Any], template_path: str, out_filepath: str) -> str:
    return export_3_mu_chu_test(result, template_path, out_filepath)

def export_3_m8_test(result: Dict[str, Any], template_path: str, out_filepath: str) -> str:
    return export_3_mu_chu_test(result, template_path, out_filepath)

def export_3_m9_test(result: Dict[str, Any], template_path: str, out_filepath: str) -> str:
    return export_3_mu_chu_test(result, template_path, out_filepath)

























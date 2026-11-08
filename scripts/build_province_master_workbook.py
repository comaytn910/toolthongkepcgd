"""
Script đóng gói toàn bộ 33 biểu nghiệp vụ PCGD - XMC năm 2026
vào 01 file Sổ cái Excel hợp nhất duy nhất cấp tỉnh:
outputs/SO_CAI_TONG_HOP_PCGD_2026_THAI_NGUYEN.xlsx
"""

import os
import sys
import glob
import time
import openpyxl
from copy import copy

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATH = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
OUTPUTS_DIR = os.path.join(BASE_DIR, 'outputs')
MASTER_OUTPUT_PATH = os.path.join(OUTPUTS_DIR, 'SO_CAI_TONG_HOP_PCGD_2026_THAI_NGUYEN.xlsx')

SHEET_CODES = [
    '0.M1', '0.M2', '0.M3', '0.M4', '0.M5', '0.M6', '0.M7',
    '1.M1', '1.M2', '1.M3', '1.M4', '1.M5', '1.M6', '1.M7',
    '2.M1', '2.M2', '2.M3', '2.M4', '2.M5', '2.M6', '2.M7', '2.M8', '2.M9', '2.M10',
    '3.M1', '3.M2', '3.M3', '3.M4', '3.M5', '3.M6', '3.M7', '3.M8', '3.M9'
]

def copy_sheet_data(src_ws, tgt_ws, sheet_code):
    """
    Sao chép chuẩn xác toàn bộ giá trị, công thức, định dạng số
    từ sheet tổng hợp cấp tỉnh sang sheet tương ứng trong sổ cái master.
    """
    max_r = src_ws.max_row
    max_c = src_ws.max_column

    # Sao chép kích thước cột nếu có
    for col_letter, col_dim in src_ws.column_dimensions.items():
        if col_dim.width:
            tgt_ws.column_dimensions[col_letter].width = col_dim.width

    for r in range(1, max_r + 1):
        # Chiều cao hàng
        if r in src_ws.row_dimensions and src_ws.row_dimensions[r].height:
            tgt_ws.row_dimensions[r].height = src_ws.row_dimensions[r].height

        for c in range(1, max_c + 1):
            src_cell = src_ws.cell(r, c)
            val = src_cell.value
            
            # Nếu ô trống trong src mà template có giá trị thì cân nhắc
            tgt_cell = tgt_ws.cell(r, c)
            
            if val is not None:
                tgt_cell.value = val

            # Sao chép định dạng số (number_format)
            if src_cell.number_format:
                tgt_cell.number_format = src_cell.number_format

            # Chuẩn hóa tỷ lệ nếu có định dạng %
            if src_cell.number_format and '%' in src_cell.number_format:
                tgt_cell.number_format = '0.00"%"'

    # Cập nhật thông tin tiêu đề cấp tỉnh nếu là hàng tiêu đề 1-3
    for r in range(1, 4):
        for c in range(1, 5):
            cell_val = str(tgt_ws.cell(r, c).value or '')
            if 'UBND' in cell_val or 'XÃ' in cell_val or 'Xã' in cell_val or 'Trường' in cell_val:
                tgt_ws.cell(r, c).value = "SỞ GIÁO DỤC VÀ ĐÀO TẠO TỈNH THÁI NGUYÊN"
                break

def build_master_workbook():
    print("==========================================================================")
    print("BẮT ĐẦU ĐÓNG GÓI TOÀN BỘ 33 BIỂU VÀO SỔ CÁI HỢP NHẤT DUY NHẤT CẤP TỈNH")
    print("==========================================================================")
    t0 = time.time()

    print(f"1. Nạp workbook mẫu chuẩn: {TEMPLATE_PATH}")
    wb_master = openpyxl.load_workbook(TEMPLATE_PATH)
    print(f"   Đã nạp template gốc với {len(wb_master.sheetnames)} sheets.")

    # Cập nhật từng sheet trong 33 biểu
    for idx, sc in enumerate(SHEET_CODES, 1):
        t_s = time.time()
        out_file = os.path.join(OUTPUTS_DIR, f"TONG_HOP_THU_{sc}_PCGD_2026.xlsx")
        if not os.path.exists(out_file):
            print(f"   [CẢNH BÁO] Không tìm thấy tệp kết quả: {out_file}")
            continue

        if sc not in wb_master.sheetnames:
            print(f"   [CẢNH BÁO] Sheet {sc} không có trong template master!")
            continue

        tgt_ws = wb_master[sc]

        # Nạp tệp nguồn tương ứng
        # Lưu ý: Với 0.M7 (sổ theo dõi danh sách 107k dòng), trong sổ cái cấp tỉnh ta nạp sheet tổng hợp
        wb_src = openpyxl.load_workbook(out_file, data_only=False)
        src_sheet_name = 'TONG_HOP_THU'
        if src_sheet_name not in wb_src.sheetnames:
            if 'TONG_HOP_TOAN_TINH' in wb_src.sheetnames:
                src_sheet_name = 'TONG_HOP_TOAN_TINH'
            elif 'TONG_HOP_1.M4' in wb_src.sheetnames:
                src_sheet_name = 'TONG_HOP_1.M4'
            else:
                src_sheet_name = wb_src.sheetnames[0]

        src_ws = wb_src[src_sheet_name]
        copy_sheet_data(src_ws, tgt_ws, sc)
        wb_src.close()

        dur = time.time() - t_s
        print(f"   [{idx:02d}/33] Đã đóng gói biểu {sc:<5} (từ {src_sheet_name}) vào Sổ cái ({dur:.2f}s)")

    # Cập nhật sheet TỔNG_QUAN
    if 'TỔNG_QUAN' in wb_master.sheetnames:
        ws_tq = wb_master['TỔNG_QUAN']
        # Đặt tên tỉnh
        ws_tq['A1'] = "TỔNG HỢP TOÀN TỈNH THÁI NGUYÊN - NĂM 2026"
        ws_tq['A2'] = "Đơn vị: SỞ GIÁO DỤC VÀ ĐÀO TẠO TỈNH THÁI NGUYÊN (92 XÃ/PHƯỜNG/THỊ TRẤN)"
        print("   Đã cập nhật thông tin đơn vị cấp tỉnh trên sheet TỔNG_QUAN.")

    # Lưu Sổ cái tổng hợp
    print(f"\n2. Ghi Sổ cái hợp nhất ra tệp: {MASTER_OUTPUT_PATH}")
    try:
        wb_master.save(MASTER_OUTPUT_PATH)
        wb_master.close()
        saved_file = MASTER_OUTPUT_PATH
    except PermissionError:
        alt_path = MASTER_OUTPUT_PATH.replace('.xlsx', '_capnhat.xlsx')
        print(f"   [CHÚ Ý KHÓA FILE] Đang mở tệp chính, lưu dự phòng ra: {alt_path}")
        wb_master.save(alt_path)
        wb_master.close()
        saved_file = alt_path

    sz = os.path.getsize(saved_file) / 1024
    dur_all = time.time() - t0
    print("==========================================================================")
    print(f"HOÀN THÀNH ĐÓNG GÓI SỔ CÁI HỢP NHẤT TOÀN TỈNH TRONG {dur_all:.2f} GIÂY!")
    print(f"Đường dẫn tệp: {saved_file} (Kích thước: {sz:.1f} KB)")
    print("==========================================================================\n")
    return saved_file

if __name__ == '__main__':
    build_master_workbook()

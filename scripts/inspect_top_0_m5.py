import os
import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)
DATA_DIR = os.path.join(BASE_DIR, 'data')

from core.excel_reader import scan_data_directory
scan_res = scan_data_directory(DATA_DIR)
files_map = {f['commune_name']: f['filepath'] for f in scan_res['files'] if f['is_valid']}

target_communes = [
    '32. Xã Phúc Lộc',
    '78. Phường Vạn Xuân',
    '44. Xã Ngân Sơn',
    '22. Xã Thần Sa',
    '30. Xã Quân Chu',
    '66. Xã Chợ Rã',
    '01. Xã Yên Trạch',
    '17. Xã Thượng Quan',
    '39. Xã Phú Bình',
    '91. Phường Bắc Kạn'
]

print("=== SOI CHI TIẾT CÁC XÃ ĐÓNG GÓP LỚN NHẤT BIỂU 0.M5 ===")

for c_name in target_communes:
    f_path = files_map.get(c_name)
    if not f_path:
        print(f"Không tìm thấy file của {c_name}")
        continue
    print(f"\n=======================================================")
    print(f"XÃ: {c_name} | File: {os.path.basename(f_path)}")
    print(f"=======================================================")
    
    wb_val = openpyxl.load_workbook(f_path, data_only=True)
    wb_raw = openpyxl.load_workbook(f_path, data_only=False)
    ws_v = wb_val['0.M5']
    ws_r = wb_raw['0.M5']
    
    # In tiêu đề hàng 5, 6
    print("Tiêu đề cột hàng 5-6:")
    for r in [5, 6]:
        row_str = [f"{openpyxl.utils.get_column_letter(c)}{r}: {ws_v.cell(r, c).value}" for c in range(1, 11) if ws_v.cell(r, c).value is not None]
        print("  " + " | ".join(row_str))
        
    print("\nChi tiết các hàng tài chính then chốt (Hàng 8 đến 15):")
    print(f"{'Hàng':<5} | {'Nội dung':<35} | {'ĐVT':<8} | {'D (Tổng)':<18} | {'E (2026)':<18} | {'F (2027)':<15} | {'Công thức E':<25}")
    print("-" * 135)
    for r in range(8, 16):
        b_val = str(ws_v.cell(r, 2).value or '').strip()[:34]
        c_val = str(ws_v.cell(r, 3).value or '').strip()
        d_val = ws_v.cell(r, 4).value
        e_val = ws_v.cell(r, 5).value
        f_val = ws_v.cell(r, 6).value
        e_raw = str(ws_r.cell(r, 5).value or '')
        
        d_str = f"{d_val:,.2f}" if isinstance(d_val, (int, float)) else str(d_val)
        e_str = f"{e_val:,.2f}" if isinstance(e_val, (int, float)) else str(e_val)
        f_str = f"{f_val:,.2f}" if isinstance(f_val, (int, float)) else str(f_val)
        
        print(f"{r:<5} | {b_val:<35} | {c_val:<8} | {d_str:<18} | {e_str:<18} | {f_str:<15} | {e_raw:<25}")
        
    wb_val.close()
    wb_raw.close()

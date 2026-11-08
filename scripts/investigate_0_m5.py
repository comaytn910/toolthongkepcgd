import os
import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)
TEMPLATE_PATH = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
DATA_DIR = os.path.join(BASE_DIR, 'data')

def investigate_template():
    print("=== 1. KIỂM TRA MẪU CHUẨN BIỂU 0.M5 ===")
    wb = openpyxl.load_workbook(TEMPLATE_PATH, data_only=False)
    ws = wb['0.M5']
    
    print("\n--- Tiêu đề hàng 1-10 ---")
    for r in range(1, 11):
        row_str = []
        for c in range(1, 12):
            col_letter = openpyxl.utils.get_column_letter(c)
            v = ws.cell(r, c).value
            if v is not None:
                row_str.append(f"{col_letter}{r}: {str(v).strip()}")
        if row_str:
            print(" | ".join(row_str))
            
    print("\n--- Hàng 8 (Chỉ tiêu Tổng chi) ---")
    for c in range(1, 12):
        col_letter = openpyxl.utils.get_column_letter(c)
        cell = ws.cell(8, c)
        print(f"  {col_letter}8: value={cell.value!r}, formula={cell.value if str(cell.value).startswith('=') else None}, number_format={cell.number_format}")
        
    print("\n--- Hàng 9 (Ngân sách thường xuyên) ---")
    for c in range(1, 12):
        col_letter = openpyxl.utils.get_column_letter(c)
        cell = ws.cell(9, c)
        print(f"  {col_letter}9: value={cell.value!r}, formula={cell.value if str(cell.value).startswith('=') else None}, number_format={cell.number_format}")

def investigate_communes():
    print("\n=== 2. QUÉT GIÁ TRỊ HÀNG 8 VÀ CỘT D:I Ở 92 XÃ ===")
    wb_tpl = openpyxl.load_workbook(TEMPLATE_PATH, data_only=True)
    ws_tpl = wb_tpl['0.M5']
    
    from core.excel_reader import scan_data_directory
    scan_res = scan_data_directory(DATA_DIR)
    valid_files = [f for f in scan_res['files'] if f['is_valid']]
    
    commune_rows = []
    
    for f in valid_files:
        c_name = f['commune_name']
        f_path = f['filepath']
        
        try:
            wb_data = openpyxl.load_workbook(f_path, data_only=True)
            wb_form = openpyxl.load_workbook(f_path, data_only=False)
            ws_d = wb_data['0.M5']
            ws_f = wb_form['0.M5']
            
            # Đơn vị tính ghi trên file của xã (thường ở hàng 5-7)
            unit_str = ""
            for r in range(4, 8):
                for c in range(1, 12):
                    v = str(ws_d.cell(r, c).value or '')
                    if 'Đơn vị' in v or 'đơn vị' in v or 'triệu' in v.lower() or 'đồng' in v.lower():
                        unit_str += f"[{openpyxl.utils.get_column_letter(c)}{r}: {v}] "
                        
            # Lấy các giá trị D8, E8, F8, G8, H8, I8
            vals_d = {}
            forms_d = {}
            for col in ['D', 'E', 'F', 'G', 'H', 'I']:
                c_idx = openpyxl.utils.column_index_from_string(col)
                vals_d[col] = ws_d.cell(8, c_idx).value
                forms_d[col] = ws_f.cell(8, c_idx).value
                
            wb_data.close()
            wb_form.close()
            
            commune_rows.append({
                'commune': c_name,
                'file': os.path.basename(f_path),
                'unit_str': unit_str.strip(),
                'vals': vals_d,
                'forms': forms_d
            })
        except Exception as e:
            print(f"Lỗi đọc {c_name}: {e}")
            
    # Sắp xếp theo giá trị E8 giảm dần
    def get_num(v):
        if isinstance(v, (int, float)): return v
        try: return float(str(v).replace(',', ''))
        except: return 0.0

    print(f"\n--- TOP 15 XÃ CÓ GIÁ TRỊ E8 LỚN NHẤT ---")
    sorted_by_e8 = sorted(commune_rows, key=lambda x: get_num(x['vals']['E']), reverse=True)
    
    print(f"{'STT':<4} | {'Xã/Phường':<30} | {'Đơn vị ghi':<25} | {'D8 (Tổng)':<18} | {'E8':<18} | {'F8':<14} | {'G8':<14}")
    print("-" * 135)
    for idx, item in enumerate(sorted_by_e8[:20], 1):
        v = item['vals']
        d_val = f"{get_num(v['D']):,.2f}" if v['D'] is not None else "None"
        e_val = f"{get_num(v['E']):,.2f}" if v['E'] is not None else "None"
        f_val = f"{get_num(v['F']):,.2f}" if v['F'] is not None else "None"
        g_val = f"{get_num(v['G']):,.2f}" if v['G'] is not None else "None"
        u_str = item['unit_str'][:24]
        print(f"{idx:<4} | {item['commune']:<30} | {u_str:<25} | {d_val:<18} | {e_val:<18} | {f_val:<14} | {g_val:<14}")

    print(f"\n--- TỔNG CỘNG E8 VÀ D8 CỦA 92 XÃ (THEO SỐ THỰC NGUYÊN BẢN CƠ SỞ NHẬP) ---")
    sum_e8 = sum(get_num(x['vals']['E']) for x in commune_rows)
    sum_d8 = sum(get_num(x['vals']['D']) for x in commune_rows)
    print(f"Tổng E8 nguồn trực tiếp = {sum_e8:,.2f}")
    print(f"Tổng D8 nguồn trực tiếp = {sum_d8:,.2f}")
    
    # Kiểm tra phân bố độ lớn của E8: ai nhập triệu đồng, ai nhập đồng?
    print("\n--- PHÂN LOẠI ĐỘ LỚN CỦA E8 ---")
    cat_ty = [x for x in commune_rows if get_num(x['vals']['E']) >= 1_000_000_000] # >= 1 tỷ (nếu đơn vị là triệu đồng thì là 1 triệu tỷ? Hay là đơn vị ĐỒNG?)
    cat_trieu = [x for x in commune_rows if 1_000_000 <= get_num(x['vals']['E']) < 1_000_000_000] # 1 triệu đến 1 tỷ
    cat_nghin = [x for x in commune_rows if 1_000 <= get_num(x['vals']['E']) < 1_000_000] # 1 nghìn đến 1 triệu (vài tỷ đồng nếu là triệu)
    cat_nho = [x for x in commune_rows if 0 < get_num(x['vals']['E']) < 1_000] # < 1000 triệu (tức dưới 1 tỷ đồng)
    cat_zero = [x for x in commune_rows if get_num(x['vals']['E']) == 0]
    
    print(f"  Nhóm >= 1,000,000,000 (hàng tỷ trở lên): {len(cat_ty)} xã")
    for x in cat_ty:
        print(f"    + {x['commune']}: E8 = {get_num(x['vals']['E']):,.2f} | D8 = {get_num(x['vals']['D']):,.2f} | Đơn vị ghi: {x['unit_str']} | Form E8: {x['forms']['E']}")
        
    print(f"  Nhóm 1,000,000 đến < 1,000,000,000 (hàng triệu đến hàng tỷ): {len(cat_trieu)} xã")
    for x in cat_trieu[:10]:
        print(f"    + {x['commune']}: E8 = {get_num(x['vals']['E']):,.2f} | D8 = {get_num(x['vals']['D']):,.2f} | Đơn vị ghi: {x['unit_str']}")
        
    print(f"  Nhóm 1,000 đến < 1,000,000: {len(cat_nghin)} xã")
    for x in cat_nghin[:10]:
        print(f"    + {x['commune']}: E8 = {get_num(x['vals']['E']):,.2f} | D8 = {get_num(x['vals']['D']):,.2f} | Đơn vị ghi: {x['unit_str']}")
        
    print(f"  Nhóm < 1,000: {len(cat_nho)} xã")
    for x in cat_nho[:5]:
        print(f"    + {x['commune']}: E8 = {get_num(x['vals']['E']):,.2f} | D8 = {get_num(x['vals']['D']):,.2f} | Đơn vị ghi: {x['unit_str']}")
        
    print(f"  Nhóm = 0 hoặc trống: {len(cat_zero)} xã")

if __name__ == '__main__':
    investigate_template()
    investigate_communes()

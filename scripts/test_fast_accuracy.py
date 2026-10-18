import os
import sys
import openpyxl

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from core.fast_excel import fast_read_sheet

def test_fast_vs_openpyxl():
    f1 = os.path.join(BASE_DIR, 'data', '01_01. Xã Yên Trạch - Nhập hoàn chỉnh PCGD 2026.xlsx')
    
    for s_name in ['0.M1', '0.M5', '0.M6', '1.M1', '2.M1', '0.M2']:
        print(f"\nĐối chiếu sheet {s_name}...")
        fast_cells = fast_read_sheet(f1, s_name)
        
        wb = openpyxl.load_workbook(f1, data_only=True)
        ws = wb[s_name]
        
        diff_count = 0
        total_checked = 0
        
        for r in range(1, min(ws.max_row + 1, 65)):
            for c in range(1, min(ws.max_column + 1, 30)):
                coord = f"{openpyxl.utils.get_column_letter(c)}{r}"
                op_val = ws.cell(row=r, column=c).value
                fast_info = fast_cells.get(coord, {'val': None, 'type': 'EMPTY'})
                fast_val = fast_info['val']
                
                # So sánh giá trị số
                if isinstance(op_val, (int, float)) and isinstance(fast_val, (int, float)):
                    if abs(op_val - fast_val) > 0.001:
                        print(f"  LỆCH tại {coord}: openpyxl={op_val} != fast={fast_val}")
                        diff_count += 1
                elif str(op_val or '').strip() != str(fast_val or '').strip():
                    # Chỉ kiểm tra nếu không phải khoảng trắng
                    if str(op_val or '').strip() != '':
                        print(f"  KHÁC CHUỖI tại {coord}: openpyxl={op_val} != fast={fast_val}")
                        diff_count += 1
                total_checked += 1
        wb.close()
        
        if diff_count == 0:
            print(f"  -> KHỚP 100% trên {total_checked} ô!")
        else:
            print(f"  -> Có {diff_count} ô khác biệt!")

if __name__ == '__main__':
    test_fast_vs_openpyxl()

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from core.excel_reader import scan_data_directory
from processors.m0_m1 import process_0_m1
from exporters.excel_exporter import export_0_m1_test

def main():
    sys.stdout.reconfigure(encoding='utf-8')
    base_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(base_dir, 'data')
    outputs_dir = os.path.join(base_dir, 'outputs')
    template_path = os.path.join(base_dir, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
    
    print("=== CHẠY KIỂM THỬ ĐỘC LẬP BIỂU 0.M1 ===\n")
    
    # Lần 1: Quét và tổng hợp
    scan_res = scan_data_directory(data_dir)
    print(f"Quét thư mục data: {scan_res['valid_count']} file hợp lệ.")
    
    valid_files = [f for f in scan_res['files'] if f['is_valid']]
    
    proc_res = process_0_m1(valid_files)
    print(f"Tổng hợp biểu 0.M1 hoàn thành cho: {len(proc_res['commune_names'])} xã.")
    print(f"Danh sách xã: {', '.join(proc_res['commune_names'])}")
    
    # Xuất file kết quả
    out_file = export_0_m1_test(proc_res, template_path, outputs_dir)
    print(f"Xuất file kết quả thành công: {out_file} ({os.path.getsize(out_file):,} bytes)")
    
    # Kiểm tra sai lệch cột K nguồn
    k_diffs = proc_res['source_k_discrepancies']
    print(f"\nĐối chiếu cột K nguồn: Phát hiện {len(k_diffs)} sai lệch:")
    for d in k_diffs:
        print(f"  - [{d['commune']}] {d['coord']} (Hàng {d['row']}): K nguồn = {d['source_k_val']} | K tính lại D:I/H:J = {d['expected_k_val']} (Lệch: {d['diff']:+g}) | Công thức: {d['raw_formula']}")
        
    # In một số chỉ tiêu mẫu
    agg = proc_res['aggregated_cells']
    print("\n--- MỘT SỐ CHỈ TIÊU TỔNG HỢP (2 XÃ) ---")
    print(f"1. Tổng số trẻ trong độ tuổi (Hàng 5):")
    for c in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
        print(f"   {c}5 = {agg[f'{c}5']}", end=' | ')
    print()
    print(f"2. Trẻ đến trường (Hàng 12):")
    for c in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
        print(f"   {c}12 = {agg[f'{c}12']}", end=' | ')
    print()
    print(f"3. Tỉ lệ huy động (Hàng 16, tính lại tử/mẫu):")
    for c in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
        print(f"   {c}16 = {agg[f'{c}16']}%", end=' | ')
    print()
    print(f"4. Tỉ lệ học 2b/ngày (Hàng 22, tính lại tử/mẫu):")
    for c in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
        print(f"   {c}22 = {agg[f'{c}22']}%", end=' | ')
    print()
    print(f"5. Tỉ lệ hoàn thành CT GDMN (Hàng 28, tính lại tử/mẫu):")
    for c in ['H', 'I', 'J', 'K']:
        print(f"   {c}28 = {agg[f'{c}28']}%", end=' | ')
    print("\n")
    
    # 6. Kiểm tra trực tiếp file Excel kết quả vừa xuất
    import openpyxl
    wb_out = openpyxl.load_workbook(out_file, data_only=False)
    ws_th = wb_out['TONG_HOP_THU']
    
    print("\n--- KIỂM TRA ĐỊNH DẠNG TỶ LỆ PHẦN TRĂM BIỂU 0.M1 (TONG_HOP_THU) ---")
    
    # Kiểm tra các ô người dùng yêu cầu:
    check_cells = {
        'E16': (3.45, '0.00"%"'),
        'F16': (79.13, '0.00"%"'),
        'E22': (100.0, '0.00"%"'),
        'K16': (74.78, '0.00"%"'),
        'K22': (100.0, '0.00"%"'),
        'H28': (100.0, '0.00"%"'),
        'I28': (100.0, '0.00"%"'),
        'J28': (100.0, '0.00"%"'),
        'K28': (100.0, '0.00"%"'),
    }
    
    all_ok = True
    for coord, (exp_val, exp_fmt) in check_cells.items():
        cell = ws_th[coord]
        val = cell.value
        fmt = cell.number_format
        val_match = abs(val - exp_val) < 0.01 if isinstance(val, (int, float)) else False
        fmt_match = (fmt == exp_fmt)
        status_str = "[ĐẠT]" if (val_match and fmt_match) else "[LỖI]"
        if not (val_match and fmt_match): all_ok = False
        print(f"  {status_str} Ô {coord}: Giá trị = {val} (Kỳ vọng ~{exp_val}) | Định dạng = {fmt!r} (Kỳ vọng {exp_fmt!r})")
        
    # Kiểm tra ô 'x' tại D28:G28
    print("\n--- KIỂM TRA BẢO TOÀN Ô 'x' VÀ TRẠNG THÁI CHƯA CHỐT ---")
    for col in ['D', 'E', 'F', 'G']:
        cell = ws_th[f"{col}28"]
        status_x = "[ĐẠT]" if cell.value == 'x' and cell.number_format != '0.00"%"' else "[LỖI]"
        print(f"  {status_x} Ô {col}28: Giá trị = {cell.value!r} | Định dạng = {cell.number_format!r}")
        
    # Kiểm tra C33:C40
    print("\n--- KIỂM TRA PHẦN 2: C33:C40 ---")
    for r in range(33, 41):
        cell_c = ws_th[f"C{r}"]
        val = cell_c.value
        fmt = cell_c.number_format
        if isinstance(val, (int, float)):
            stt = "[ĐẠT - Số]" if fmt == '0.00"%"' else "[LỖI]"
        elif val == "Chưa chốt quy tắc":
            stt = "[ĐẠT - Chưa chốt]"
        else:
            stt = "[TIÊU ĐỀ/KHÁC]"
        print(f"  {stt} C{r}: Giá trị = {val!r} | Định dạng = {fmt!r}")
        
    # Kiểm tra các tỷ lệ vượt 100%
    print("\n--- CẢNH BÁO KIỂM TRA PHẠM VI TỬ SỐ / MẪU SỐ (> 100%) ---")
    for col in ['G', 'H', 'I', 'J']:
        cell = ws_th[f"{col}16"]
        print(f"  [BẢO TOÀN >100%] {col}16: Giá trị = {cell.value}% | Định dạng = {cell.number_format!r}")
        
    wb_out.close()
    
    # Kiểm tra tính lại từ đầu (Chạy lần 2 để đảm bảo KHÔNG cộng dồn)
    print("\n--- KIỂM TRA CHẠY LẠI KHÔNG CỘNG DỒN ---")
    proc_res_2 = process_0_m1(valid_files)
    if proc_res_2['aggregated_cells']['K5'] == agg['K5']:
        print("[ĐẠT] Kết quả lần 2 giống hệt lần 1, không hề bị cộng dồn (K5 = 1.620).")
    else:
        print("[LỖI] Bị cộng dồn kết quả!")

if __name__ == '__main__':
    main()

import os
import sys
import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from core.excel_reader import scan_data_directory
from processors.m0_m1 import process_0_m1
from processors.m0_m6 import process_0_m6, DISABILITY_COLS, DATA_ROWS
from exporters.excel_exporter import export_0_m6_test

def test_0_m6_pipeline():
    sys.stdout.reconfigure(encoding='utf-8')
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = os.path.join(base_dir, 'data')
    outputs_dir = os.path.join(base_dir, 'outputs')
    template_path = os.path.join(base_dir, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
    
    print("=== BẮT ĐẦU KIỂM THỬ ĐỘC LẬP BIỂU 0.M6 (ĐỐI TƯỢNG KHUYẾT TẬT MN) ===\n")
    
    # 1. Quét dữ liệu
    scan = scan_data_directory(data_dir)
    valid_files = [f for f in scan['files'] if f['is_valid']]
    assert len(valid_files) >= 1, "Phải có ít nhất 1 file hợp lệ trong data"
    print(f"Quét được {len(valid_files)} file hợp lệ: {[f['commune_name'] for f in valid_files]}")
    
    # 2. Xử lý tổng hợp 0.M6
    res_m6 = process_0_m6(valid_files)
    agg = res_m6['aggregated_cells']
    
    # 3. Kiểm tra tính đúng đắn các dạng tật và đối chiếu độc lập:
    # Hàng 10 (3t): 1 trẻ khuyết tật khác (K10=1), tiếp cận GD L10=1 -> M10=100.00%
    assert agg['C10'] == 1.0, f"Lỗi C10: {agg['C10']} != 1.0"
    assert agg['K10'] == 1.0, f"Lỗi K10: {agg['K10']} != 1.0"
    assert agg['L10'] == 1.0, f"Lỗi L10: {agg['L10']} != 1.0"
    assert agg['M10'] == 100.0, f"Lỗi M10: {agg['M10']} != 100.0"
    print(f"[ĐẠT] Hàng 10 (3 tuổi): C10 = {agg['C10']} trẻ (K10={agg['K10']}), L10 = {agg['L10']}, Tỉ lệ M10 = {agg['M10']}%.")
    
    # Hàng 11 (4t): 4 trẻ (D=1, E=1, I=2), L=3 -> M=75.00%
    assert agg['C11'] == 4.0, f"Lỗi C11: {agg['C11']} != 4.0"
    assert agg['D11'] == 1.0 and agg['E11'] == 1.0 and agg['I11'] == 2.0
    assert agg['L11'] == 3.0
    assert agg['M11'] == 75.0, f"Lỗi M11: {agg['M11']} != 75.0"
    print(f"[ĐẠT] Hàng 11 (4 tuổi): C11 = {agg['C11']} trẻ, L11 = {agg['L11']}, Tỉ lệ M11 = {agg['M11']}%.")
    
    # Hàng 12 (5t): 1 trẻ (D=1), L=1 -> M=100.00%
    assert agg['C12'] == 1.0
    assert agg['D12'] == 1.0
    assert agg['L12'] == 1.0
    assert agg['M12'] == 100.0
    print(f"[ĐẠT] Hàng 12 (5 tuổi): C12 = {agg['C12']} trẻ, L12 = {agg['L12']}, Tỉ lệ M12 = {agg['M12']}%.")
    
    # Hàng 13 (Tổng cộng): C13 = 6, L13 = 5 -> M13 = 83.33%
    assert agg['C13'] == 6.0, f"Lỗi C13: {agg['C13']} != 6.0"
    assert agg['L13'] == 5.0, f"Lỗi L13: {agg['L13']} != 5.0"
    assert abs(agg['M13'] - 83.33) < 0.01, f"Lỗi M13: {agg['M13']} != 83.33"
    print(f"[ĐẠT] Hàng 13 (Tổng cộng): C13 = {agg['C13']} trẻ, L13 = {agg['L13']}, Tỉ lệ M13 = {agg['M13']}%.")
    
    # 4. Đối soát logic liên biểu với 0.M1
    res_m1 = process_0_m1(valid_files)
    agg_m1 = res_m1['aggregated_cells']
    assert agg['C12'] == agg_m1['I8'], f"Không khớp liên biểu: 0.M6 C12 ({agg['C12']}) != 0.M1 I8 ({agg_m1['I8']})"
    assert agg['L12'] == agg_m1['I10'], f"Không khớp liên biểu: 0.M6 L12 ({agg['L12']}) != 0.M1 I10 ({agg_m1['I10']})"
    print(f"[ĐẠT] Đối soát liên biểu 0.M6 và 0.M1 hoàn toàn trùng khớp (Trẻ 5t khuyết tật = {agg['C12']}, Tiếp cận GD = {agg['L12']}).")
    
    # 5. Xuất file kết quả
    out_file = export_0_m6_test(res_m6, template_path, outputs_dir)
    print(f"Xuất file kết quả thành công: {out_file} ({os.path.getsize(out_file):,} bytes)")
    
    # 6. Kiểm tra trực tiếp file Excel kết quả
    wb_out = openpyxl.load_workbook(out_file, data_only=False)
    ws_th = wb_out['TONG_HOP_THU']
    
    # Cột M phải có định dạng '0.00"%"'
    for r in range(7, 14):
        cell_m = ws_th[f"M{r}"]
        assert cell_m.number_format == '0.00"%"', f"M{r} format sai: {cell_m.number_format}"
    print(f"[ĐẠT] Cột M (Tỉ lệ tiếp cận GD) nhận đúng định dạng '0.00\"%\"' (M10={ws_th['M10'].value}%, M11={ws_th['M11'].value}%, M13={ws_th['M13'].value}%).")
    
    # Cột C..L là số lượng: '#,##0'
    for r in range(7, 14):
        for c in ['C', 'D', 'L']:
            cell = ws_th[f"{c}{r}"]
            assert '%' not in cell.number_format, f"{c}{r} bị dán nhầm %: {cell.number_format}"
            assert cell.number_format == '#,##0'
    print("[ĐẠT] Các cột số lượng người khuyết tật (C..L) nhận đúng định dạng '#,##0', không có ký hiệu %.")
    
    wb_out.close()
    
    # 7. Kiểm tra chạy lần 2 không cộng dồn
    res_m6_2 = process_0_m6(valid_files)
    assert res_m6_2['aggregated_cells']['C13'] == agg['C13'], "Bị lỗi cộng dồn kết quả lần 2!"
    print("[ĐẠT] Chạy lại lần 2 độc lập không hề bị cộng dồn kết quả.")
    
    print("\n=== TẤT CẢ KIỂM THỬ BIỂU 0.M6 HOÀN TOÀN ĐẠT CHUẨN! ===")

if __name__ == '__main__':
    test_0_m6_pipeline()

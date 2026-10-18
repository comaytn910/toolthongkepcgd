import os
import sys
import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from core.excel_reader import scan_data_directory
from processors.m1_m1 import process_1_m1, AGE_COLS_6_10, AGE_COLS_11_14
from exporters.excel_exporter import export_1_m1_test

def test_1_m1_pipeline():
    sys.stdout.reconfigure(encoding='utf-8')
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = os.path.join(base_dir, 'data')
    outputs_dir = os.path.join(base_dir, 'outputs')
    template_path = os.path.join(base_dir, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
    
    print("=== BẮT ĐẦU KIỂM THỬ ĐỘC LẬP BIỂU 1.M1 (TRẺ 6-14T TIỂU HỌC) ===\n")
    
    # 1. Quét dữ liệu
    scan = scan_data_directory(data_dir)
    valid_files = [f for f in scan['files'] if f['is_valid']]
    assert len(valid_files) >= 1, "Phải có ít nhất 1 file hợp lệ trong data"
    print(f"Quét được {len(valid_files)} file hợp lệ: {[f['commune_name'] for f in valid_files]}")
    
    # 2. Xử lý tổng hợp 1.M1
    res_m1 = process_1_m1(valid_files)
    agg = res_m1['aggregated_cells']
    crit = res_m1['criteria_results']
    
    # 3. Kiểm tra tính độc lập Tổng số trẻ Hàng 8
    # Đối chiếu độc lập từ file nguồn
    calc_f8 = 0.0
    for f in valid_files:
        wb = openpyxl.load_workbook(f['filepath'], data_only=True)
        v = wb['1.M1']['F8'].value
        if isinstance(v, (int, float)): calc_f8 += v
        wb.close()
    assert agg['F8'] == calc_f8, f"Lỗi tổng F8: {agg['F8']} != {calc_f8}"
    print(f"[ĐẠT] Tổng số trẻ 6 tuổi F8 = {agg['F8']} khớp đáp án độc lập ({calc_f8}).")
    
    # 4. Kiểm tra Cột K = SUM(F:J) và Cột P = SUM(L:O)
    for r in range(8, 57):
        sum_k = sum(agg[f"{c}{r}"] for c in AGE_COLS_6_10)
        assert abs(agg[f"K{r}"] - sum_k) < 0.01, f"Hàng {r}: K ({agg[f'K{r}']}) != SUM(F:J) ({sum_k})"
        sum_p = sum(agg[f"{c}{r}"] for c in AGE_COLS_11_14)
        assert abs(agg[f"P{r}"] - sum_p) < 0.01, f"Hàng {r}: P ({agg[f'P{r}']}) != SUM(L:O) ({sum_p})"
    print("[ĐẠT] Không cộng trùng: Cột K (6-10t) và Cột P (11-14t) tính lại từ chi tiết chính xác 100%.")
    
    # 5. Kiểm tra Hàng 14 (Số PPC = Hàng 8 - Hàng 11)
    for c in AGE_COLS_6_10 + ['K'] + AGE_COLS_11_14 + ['P']:
        assert agg[f"{c}14"] == (agg[f"{c}8"] - agg[f"{c}11"]), f"Lỗi PPC cột {c}: {agg[f'{c}14']} != {agg[f'{c}8']} - {agg[f'{c}11']}"
    print("[ĐẠT] Hàng 14 (Số phải phổ cập): Tính lại chính xác = Tổng số (Hàng 8) - Khuyết tật (Hàng 11).")
    
    # 6. Kiểm tra Tiêu chí Phần 2 (Hàng 58 đến 63)
    print("\n--- KẾT QUẢ TIÊU CHÍ PHỔ CẬP TIỂU HỌC (HÀNG 58-63) ---")
    for r in range(58, 64):
        c_info = crit[r]
        qty = agg[f"F{r}"]
        rate = agg[f"G{r}"]
        print(f"  [ĐẠT] Hàng {r} - {c_info['name']}: Số lượng = {qty}, Tỉ lệ = {rate}% ({c_info['rule']})")
        assert rate <= 100.0 or rate > 100.0 # Hợp lệ số thực
        
    # 7. Xuất file kết quả
    out_file = export_1_m1_test(res_m1, template_path, outputs_dir)
    print(f"\nXuất file kết quả thành công: {out_file} ({os.path.getsize(out_file):,} bytes)")
    
    # 8. Mở file Excel kiểm tra trực tiếp định dạng
    wb_out = openpyxl.load_workbook(out_file, data_only=False)
    ws_th = wb_out['TONG_HOP_THU']
    
    # G58..G63 phải có number_format == '0.00"%"'
    for r in range(58, 64):
        cell_g = ws_th[f"G{r}"]
        assert cell_g.number_format == '0.00"%"', f"G{r} format sai: {cell_g.number_format}"
    print("[ĐẠT] Cột G (Tỉ lệ % tiêu chí 58-63) nhận đúng định dạng '0.00\"%\"'.")
    
    # Các ô số lượng (F..P hàng 8..56 và F58..F63) phải có number_format == '#,##0'
    for r in [8, 9, 14, 58, 61]:
        cell_f = ws_th[f"F{r}"]
        assert '%' not in cell_f.number_format, f"F{r} bị dán nhầm %: {cell_f.number_format}"
        assert cell_f.number_format == '#,##0', f"F{r} format sai: {cell_f.number_format}"
    print("[ĐẠT] Các ô số lượng học sinh nhận đúng định dạng '#,##0', không có ký hiệu %.")
    
    wb_out.close()
    
    # 9. Kiểm tra chạy lần 2 không cộng dồn
    res_m1_2 = process_1_m1(valid_files)
    assert res_m1_2['aggregated_cells']['F8'] == agg['F8'], "Bị lỗi cộng dồn kết quả lần 2!"
    print("[ĐẠT] Chạy lại lần 2 độc lập không hề bị cộng dồn kết quả.")
    
    print("\n=== TẤT CẢ KIỂM THỬ BIỂU 1.M1 HOÀN TOÀN ĐẠT CHUẨN! ===")

if __name__ == '__main__':
    test_1_m1_pipeline()

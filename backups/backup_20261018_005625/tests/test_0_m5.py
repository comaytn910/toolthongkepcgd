import os
import sys
import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from core.excel_reader import scan_data_directory
from processors.m0_m1 import process_0_m1
from processors.m0_m5 import process_0_m5, MONEY_ROWS
from exporters.excel_exporter import export_0_m5_test

def test_0_m5_pipeline():
    sys.stdout.reconfigure(encoding='utf-8')
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = os.path.join(base_dir, 'data')
    outputs_dir = os.path.join(base_dir, 'outputs')
    template_path = os.path.join(base_dir, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
    
    print("=== BẮT ĐẦU KIỂM THỬ ĐỘC LẬP BIỂU 0.M5 (TÀI CHÍNH GDMN) ===\n")
    
    # 1. Quét dữ liệu
    scan = scan_data_directory(data_dir)
    valid_files = [f for f in scan['files'] if f['is_valid']]
    assert len(valid_files) >= 1, "Phải có ít nhất 1 file hợp lệ trong data"
    print(f"Quét được {len(valid_files)} file hợp lệ: {[f['commune_name'] for f in valid_files]}")
    
    # Lấy thêm kết quả 0.M1 để có trọng số số trẻ cho Hàng 14
    res_m1 = process_0_m1(valid_files)
    
    # 2. Xử lý tổng hợp 0.M5
    res_m5 = process_0_m5(valid_files, m0_m1_result=res_m1)
    agg = res_m5['aggregated_cells']
    status = res_m5['status_cells']
    
    # 3. Đối chiếu độc lập số tiền Hàng 8 (Tổng chi GDMN)
    # Lấy độc lập từ 2 file
    calc_e8 = 0.0
    for f in valid_files:
        wb = openpyxl.load_workbook(f['filepath'], data_only=True)
        v = wb['0.M5']['E8'].value
        if isinstance(v, (int, float)):
            calc_e8 += v
        wb.close()
    calc_e8 = round(calc_e8, 2)
    assert abs(agg['E8'] - calc_e8) < 0.01, f"Lỗi tổng E8: {agg['E8']} != {calc_e8}"
    print(f"[ĐẠT] Tổng chi E8 (2026) = {agg['E8']} Tr.đg khớp đáp án độc lập ({calc_e8} Tr.đg).")
    
    # 4. Kiểm tra xử lý lỗi kiểu dữ liệu tại Hàng 36 (Yên Thịnh có ô Date 04/05/2026)
    # Đảm bảo KHÔNG bị cộng 46146.0
    print(f"[ĐẠT] Ô D36 tổng = {agg['D36']} Tr.đg (Không bị cộng sai serial date 46146.0).")
    print(f"[ĐẠT] Trạng thái D36: {status['D36']}")
    assert agg['D36'] < 1000, f"D36 bị cộng nhầm serial date: {agg['D36']}"
    
    # 5. Kiểm tra Cột D = SUM(E:I)
    for r in MONEY_ROWS:
        sum_ei = round(sum(agg[f"{c}{r}"] for c in ['E', 'F', 'G', 'H', 'I']), 2)
        assert abs(agg[f"D{r}"] - sum_ei) < 0.01, f"Hàng {r}: D ({agg[f'D{r}']}) != SUM E:I ({sum_ei})"
    print("[ĐẠT] Toàn bộ các dòng số tiền: Cột D cấp tỉnh = SUM(E:I) chính xác 100%.")
    
    # 6. Xuất file kết quả
    out_file = export_0_m5_test(res_m5, template_path, outputs_dir)
    print(f"Xuất file kết quả thành công: {out_file} ({os.path.getsize(out_file):,} bytes)")
    
    # 7. Kiểm tra trực tiếp file Excel kết quả
    wb_out = openpyxl.load_workbook(out_file, data_only=False)
    ws_th = wb_out['TONG_HOP_THU']
    
    # Kiểm tra định dạng:
    # Hàng 13 phải là 0.00"%"
    for col in ['D', 'E', 'F', 'G', 'H', 'I']:
        c13 = ws_th[f"{col}13"]
        assert c13.number_format == '0.00"%"', f"{col}13 number_format sai: {c13.number_format}"
    print(f"[ĐẠT] Hàng 13 (Tỷ lệ chi HĐ chuyên môn) nhận đúng định dạng '0.00\"%\"' (Giá trị D13={ws_th['D13'].value}%).")
    
    # Hàng 14 & Hàng 8..12, 15..38 KHÔNG được dùng định dạng %
    for r in [8, 9, 14, 15, 26, 33]:
        for col in ['D', 'E']:
            cell = ws_th[f"{col}{r}"]
            assert '%' not in cell.number_format, f"Ô số tiền {col}{r} bị dán nhầm định dạng %: {cell.number_format}"
            assert cell.number_format == '#,##0.00', f"Ô {col}{r} định dạng sai: {cell.number_format}"
    print("[ĐẠT] Các ô số tiền và mức chi bình quân nhận đúng định dạng '#,##0.00', không có ký hiệu %.")
    
    wb_out.close()
    
    # 8. Kiểm tra tính lại lần 2 không bị cộng dồn
    res_m5_2 = process_0_m5(valid_files, m0_m1_result=res_m1)
    assert res_m5_2['aggregated_cells']['E8'] == agg['E8'], "Bị lỗi cộng dồn kết quả lần 2!"
    print("[ĐẠT] Chạy lại lần 2 độc lập không hề bị cộng dồn kết quả.")
    
    print("\n=== TẤT CẢ KIỂM THỬ BIỂU 0.M5 HOÀN TOÀN ĐẠT CHUẨN! ===")

if __name__ == '__main__':
    test_0_m5_pipeline()

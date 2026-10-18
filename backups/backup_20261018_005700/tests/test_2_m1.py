import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.stdout.reconfigure(encoding='utf-8')

import openpyxl
from processors.m2_m1 import process_2_m1, DETAIL_COLS, POP_COLS, COLS_15_18, POP_COLS_15_18
from exporters.excel_exporter import export_2_m1_test
from core.excel_reader import scan_data_directory

DATA_DIR = "data"
TEMPLATE_PATH = "MẪU CHUẨN - PCGD 2026(ok).xlsx"
OUTPUT_DIR = "outputs"

def test_deduplication_and_loading_2_m1(files_info):
    """Kiểm tra nhận diện đúng xã và loại trừ file trùng"""
    assert len(files_info) >= 2
    commune_names = [f['commune_name'] for f in files_info]
    assert len(commune_names) == len(set(commune_names)), "Trùng lặp xã không được cho phép"

def test_real_data_independent_reconciliation_2_m1(valid_files):
    """
    Kiểm thử độc lập với 2 file thật trong data:
    1. Đọc độc lập từ từng file bằng openpyxl không thông qua process_2_m1.
    2. Tự tính tổng độc lập, tự tính Z, tự tính các tiêu chí.
    3. So sánh 100% với output của process_2_m1.
    """
    assert len(valid_files) == 2
    
    # 1. Tính toán độc lập
    indep_counts = {}
    indep_pop = {}
    
    for f in valid_files:
        wb = openpyxl.load_workbook(f['filepath'], data_only=True)
        ws = wb['2.M1']
        
        # Đọc dân số 7-12
        for r in range(7, 13):
            for col in POP_COLS:
                val = ws[f"{col}{r}"].value
                num = float(val) if isinstance(val, (int, float)) else 0.0
                indep_pop[(col, r)] = indep_pop.get((col, r), 0.0) + num
                
        # Đọc chi tiết 15-49
        for r in range(15, 50):
            for col in DETAIL_COLS:
                val = ws[f"{col}{r}"].value
                num = float(val) if isinstance(val, (int, float)) else 0.0
                indep_counts[(col, r)] = indep_counts.get((col, r), 0.0) + num
        wb.close()
                
    # Tính độc lập hàng 13
    for col in POP_COLS:
        indep_pop[(col, 13)] = indep_pop[(col, 7)] - indep_pop[(col, 10)]
        
    # Tính độc lập Z cấp tỉnh
    indep_z = {}
    for r in range(7, 13):
        indep_z[r] = sum(indep_pop[(col, r)] for col in POP_COLS_15_18)
    indep_z[13] = indep_z[7] - indep_z[10]
    for r in range(15, 50):
        indep_z[r] = sum(indep_counts[(col, r)] for col in COLS_15_18)
        
    # Tính độc lập tiêu chí
    # G55: sum(N35:Y37)
    indep_g55 = sum(indep_counts[(col, r)] for r in range(35, 38) for col in COLS_15_18)
    indep_h55 = round(indep_g55 / indep_z[13] * 100, 2)
    
    # G56: sum(N40:Y42)
    indep_g56 = sum(indep_counts[(col, r)] for r in range(40, 43) for col in COLS_15_18)
    indep_h56 = round(indep_g56 / indep_z[13] * 100, 2)
    
    # G57: sum(F12:Y12)
    indep_g57 = sum(indep_pop[(col, 12)] for col in POP_COLS)
    indep_denom57 = sum(indep_pop[(col, 11)] for col in POP_COLS)
    indep_h57 = round(indep_g57 / indep_denom57 * 100, 2)
    
    # G58: sum(F15:Y17, F20:Y22, F25:Y27, F30:Y32, F40:Y42)
    rows_g58 = list(range(15, 18)) + list(range(20, 23)) + list(range(25, 28)) + list(range(30, 33)) + list(range(40, 43))
    indep_g58 = sum(indep_counts[(col, r)] for r in rows_g58 for col in DETAIL_COLS)
    indep_denom58 = sum(indep_pop[(col, 13)] for col in POP_COLS)
    indep_h58 = round(indep_g58 / indep_denom58 * 100, 2)
    
    # G59: sum(W35:Y37)
    cols_18 = ['W', 'X', 'Y']
    indep_g59 = sum(indep_counts[(col, r)] for r in range(35, 38) for col in cols_18)
    indep_w13 = indep_pop[('W', 13)]
    indep_h59 = round(indep_g59 / indep_w13 * 100, 2) # Chuẩn hóa thang 0-100!
    
    # 2. Chạy hàm kiểm thử
    res1 = process_2_m1(valid_files)
    
    # Chạy lại lần 2 để kiểm tra không bị cộng dồn
    res2 = process_2_m1(valid_files)
    assert res1['aggregated_cells'] == res2['aggregated_cells'], "Chạy lại bị sai khác kết quả"
    
    agg = res1['aggregated_cells']
    
    # Đối chiếu dân số 7-12
    for r in range(7, 13):
        for col in POP_COLS:
            assert agg[f"{col}{r}"] == indep_pop[(col, r)], f"Lệch dân số ô {col}{r}"
        assert agg[f"Z{r}"] == indep_z[r], f"Lệch Z{r}"
        
    # Đối chiếu hàng 13
    for col in POP_COLS:
        assert agg[f"{col}13"] == indep_pop[(col, 13)], f"Lệch PPC {col}13"
    assert agg["Z13"] == indep_z[13], "Lệch PPC Z13"
    
    # Đối chiếu chi tiết 15-49
    for r in range(15, 50):
        for col in DETAIL_COLS:
            assert agg[f"{col}{r}"] == indep_counts[(col, r)], f"Lệch chi tiết {col}{r}"
        assert agg[f"Z{r}"] == indep_z[r], f"Lệch Z{r}"
        
    # Đối chiếu tiêu chí
    assert agg['G55'] == indep_g55, f"Lệch G55: {agg['G55']} vs {indep_g55}"
    assert agg['H55'] == indep_h55, f"Lệch H55: {agg['H55']} vs {indep_h55}"
    
    assert agg['G56'] == indep_g56, f"Lệch G56: {agg['G56']} vs {indep_g56}"
    assert agg['H56'] == indep_h56, f"Lệch H56: {agg['H56']} vs {indep_h56}"
    
    assert agg['G57'] == indep_g57, f"Lệch G57: {agg['G57']} vs {indep_g57}"
    assert agg['H57'] == indep_h57, f"Lệch H57: {agg['H57']} vs {indep_h57}"
    
    assert agg['G58'] == indep_g58, f"Lệch G58: {agg['G58']} vs {indep_g58}"
    assert agg['H58'] == indep_h58, f"Lệch H58: {agg['H58']} vs {indep_h58}"
    
    assert agg['G59'] == indep_g59, f"Lệch G59: {agg['G59']} vs {indep_g59}"
    assert agg['H59'] == indep_h59, f"Lệch H59: {agg['H59']} vs {indep_h59}"
    
    # Kiểm tra riêng H59 thang 0-100
    assert agg['H59'] >= 90.0 and agg['H59'] <= 100.0, f"H59 chưa chuẩn hóa thang 0-100: {agg['H59']}"

def test_export_and_formatting_2_m1(valid_files):
    """Kiểm tra file Excel xuất ra: định dạng 0.00% cho tỷ lệ, #,##0 cho số lượng"""
    result = process_2_m1(valid_files)
    
    out_path = export_2_m1_test(result, TEMPLATE_PATH, OUTPUT_DIR)
    assert os.path.exists(out_path)
    
    wb = openpyxl.load_workbook(out_path, data_only=False)
    assert 'TONG_HOP_THU' in wb.sheetnames
    assert 'CHI_TIET_THEO_XA' in wb.sheetnames
    assert 'DOI_CHIEU_Z_NGUON' in wb.sheetnames
    
    ws = wb['TONG_HOP_THU']
    
    # 1. Kiểm tra định dạng tỷ lệ H55..H59
    RATE_FMT = '0.00"%"'
    for r in [55, 56, 57, 58, 59]:
        cell = ws[f"H{r}"]
        val = cell.value
        fmt = cell.number_format
        assert fmt == RATE_FMT, f"H{r} có định dạng {fmt}, kỳ vọng {RATE_FMT}"
        assert isinstance(val, (int, float)), f"H{r} phải là số: {val}"
        assert val >= 0.0 and val <= 105.0, f"H{r} không nằm trong thang 0-100: {val}"
        
    # 2. Kiểm tra định dạng số lượng G55..G59 và các ô chi tiết
    COUNT_FMT = '#,##0'
    for r in [55, 56, 57, 58, 59]:
        cell = ws[f"G{r}"]
        fmt = cell.number_format
        assert fmt == COUNT_FMT, f"G{r} có định dạng {fmt}, kỳ vọng {COUNT_FMT}"
        
    for r in range(7, 14):
        for col in POP_COLS + ['Z']:
            fmt = ws[f"{col}{r}"].number_format
            assert fmt == COUNT_FMT, f"Dân số {col}{r} có định dạng {fmt}"
            
    for r in range(15, 50):
        for col in DETAIL_COLS + ['Z']:
            fmt = ws[f"{col}{r}"].number_format
            assert fmt == COUNT_FMT, f"Chi tiết {col}{r} có định dạng {fmt}"
            
    wb.close()

def test_regression_0_m1(valid_files):
    """Kiểm tra không làm ảnh hưởng đến biểu 0.M1 đã hoàn thành"""
    from processors.m0_m1 import process_0_m1
    from exporters.excel_exporter import export_0_m1_test
    
    res_0_m1 = process_0_m1(valid_files)
    out_0_m1 = export_0_m1_test(res_0_m1, TEMPLATE_PATH, OUTPUT_DIR)
    
    assert os.path.exists(out_0_m1)
    wb = openpyxl.load_workbook(out_0_m1, data_only=False)
    ws = wb['TONG_HOP_THU']
    assert ws['E16'].number_format == '0.00"%"'
    assert ws['E16'].value == 3.45
    assert ws['E22'].number_format == '0.00"%"'
    assert ws['E22'].value == 100.00
    wb.close()

if __name__ == '__main__':
    print("=== BẮT ĐẦU CHẠY KIỂM THỬ TỰ ĐỘNG BIỂU 2.M1 ===", flush=True)
    scan = scan_data_directory(DATA_DIR)
    files_info = scan['files']
    valid_files = [f for f in files_info if f['is_valid']]
    
    test_deduplication_and_loading_2_m1(files_info)
    print("[ĐẠT] 1. Kiểm tra nhận diện xã và chống trùng lặp file.", flush=True)
    
    test_real_data_independent_reconciliation_2_m1(valid_files)
    print("[ĐẠT] 2. Đối chiếu số liệu độc lập 100% khớp nguồn, cột Z tính lại, tiêu chí H55-H59 chuẩn xác.", flush=True)
    
    test_export_and_formatting_2_m1(valid_files)
    print("[ĐẠT] 3. Xuất file Excel thành công: Định dạng 0.00\"%\" cho H55-H59, #,##0 cho số lượng.", flush=True)
    
    test_regression_0_m1(valid_files)
    print("[ĐẠT] 4. Hồi quy biểu 0.M1: Không bị ảnh hưởng khi bổ sung biểu 2.M1.", flush=True)
    
    print("=== TẤT CẢ KIỂM THỬ BIỂU 2.M1 ĐÃ ĐẠT 100%! ===", flush=True)

"""
Kiểm thử độc lập Biểu 3.M2: Thống kê đạt chuẩn XMC Mức 2 (15-35 tuổi)
Hệ thống Tổng hợp số liệu PCGD - XMC tỉnh Thái Nguyên năm 2026
"""

import os
import sys
import glob
import time
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from processors.m3_m2 import process_3_m2_single_file, aggregate_3_m2
from exporters.excel_exporter import export_3_m2_test


def test_3_m2():
    print("=" * 60)
    print("BẮT ĐẦU KIỂM THỬ ĐỘC LẬP BIỂU 3.M2 (CHUẨN XMC MỨC 2 15-35T)")
    print("=" * 60)
    
    start_time = time.time()
    files = sorted(glob.glob('data/*.xlsx'))
    assert len(files) == 92, f"Kỳ vọng 92 files, thực tế: {len(files)}"
    
    # 1. Đọc và xử lý từng xã
    communes_data = []
    for f in files:
        c_name = os.path.basename(f).split(' - ')[0]
        data = process_3_m2_single_file(f, c_name)
        assert data is not None, f"Không đọc được sheet 3.M2 từ {c_name}"
        communes_data.append(data)
        
    assert len(communes_data) == 92, f"Kỳ vọng 92 xã, thực tế: {len(communes_data)}"
    print(f"-> Test 1 PASS: Đọc thành công 92/92 xã.")
    
    # 2. Tổng hợp cấp tỉnh
    result = aggregate_3_m2(communes_data)
    assert result['total_communes'] == 92
    assert result['sum_c_dan_so'] == 479435.0, f"Tổng C dân số sai: {result['sum_c_dan_so']}"
    assert result['sum_d_biet_chu_m1'] == 478664.0, f"Tổng D biết chữ M1 sai: {result['sum_d_biet_chu_m1']}"
    assert result['sum_f_biet_chu_m2'] == 477517.0, f"Tổng F biết chữ M2 sai: {result['sum_f_biet_chu_m2']}"
    
    # Kiểm tra tỷ lệ cấp tỉnh
    expected_rate_m1 = round(478664.0 / 479435.0 * 100, 2)
    expected_rate_m2 = round(477517.0 / 479435.0 * 100, 2)
    assert result['rate_m1_prov'] == expected_rate_m1, f"Tỷ lệ M1 sai: {result['rate_m1_prov']} vs {expected_rate_m1}"
    assert result['rate_m2_prov'] == expected_rate_m2, f"Tỷ lệ M2 sai: {result['rate_m2_prov']} vs {expected_rate_m2}"
    print(f"-> Test 2 PASS: Tổng hợp cấp tỉnh chuẩn xác 100% (Dân số: {result['sum_c_dan_so']:,.0f}, M1: {result['rate_m1_prov']}%, M2: {result['rate_m2_prov']}%).")
    
    # 3. Xuất file Excel
    os.makedirs('outputs', exist_ok=True)
    out_path = 'outputs/TONG_HOP_THU_3.M2_PCGD_2026.xlsx'
    saved_path = export_3_m2_test(result, 'MẪU CHUẨN - PCGD 2026(ok).xlsx', out_path)
    assert os.path.exists(saved_path), f"File không tồn tại: {saved_path}"
    print(f"-> Test 3 PASS: Xuất file Excel thành công tại {saved_path}.")
    
    # 4. Kiểm tra cấu trúc file Excel xuất ra
    wb = openpyxl.load_workbook(saved_path, data_only=True)
    assert 'TONG_HOP_THU' in wb.sheetnames, "Thiếu sheet TONG_HOP_THU"
    assert 'CHI_TIET_92_XA' in wb.sheetnames, "Thiếu sheet CHI_TIET_92_XA"
    
    ws_th = wb['TONG_HOP_THU']
    total_val = None
    for r in range(1, ws_th.max_row + 1):
        v1 = str(ws_th.cell(r, 2).value or '')
        if 'cộng toàn tỉnh' in v1.lower():
            total_val = ws_th.cell(r, 3).value
            rate_m1_val = ws_th.cell(r, 5).value
            rate_m2_val = ws_th.cell(r, 7).value
            break
            
    assert total_val == 479435.0, f"Giá trị ô tổng Cột C trong sheet TONG_HOP_THU sai: {total_val}"
    assert rate_m1_val == 99.84, f"Giá trị ô tỷ lệ M1 sai: {rate_m1_val}"
    assert rate_m2_val == 99.60, f"Giá trị ô tỷ lệ M2 sai: {rate_m2_val}"
    wb.close()
    print(f"-> Test 4 PASS: File Excel kiểm tra ngược dữ liệu khớp 100%.")
    
    elapsed = time.time() - start_time
    print(f"\nTOÀN BỘ 4/4 TESTS ĐẠT 100% TRONG {elapsed:.2f} GIÂY!")
    print("=" * 60)


if __name__ == '__main__':
    test_3_m2()

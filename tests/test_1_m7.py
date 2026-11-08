import os
import sys
from pathlib import Path
import openpyxl

# Thêm thư mục gốc vào PYTHONPATH
sys.path.insert(0, str(Path(__file__).parent.parent))

from processors.m1_m7 import process_1_m7
from exporters.excel_exporter import export_1_m7_test

def test_1_m7_aggregation():
    data_dir = Path("data")
    files = sorted([f for f in data_dir.glob("*.xlsx") if not f.name.startswith("~$")])
    
    valid_files = []
    for idx, f in enumerate(files, 1):
        valid_files.append({
            'filepath': str(f),
            'filename': f.name,
            'commune_name': f.stem,
            'commune_stt': idx
        })
        
    print(f"Bắt đầu kiểm thử Biểu 1.M7 trên {len(valid_files)} xã...")
    result = process_1_m7(valid_files)
    
    # 1. Kiểm tra số lượng xã
    assert result['total_communes_count'] == 92, f"Expected 92 communes, got {result['total_communes_count']}"
    print("PASS 1: Đủ 92/92 xã có dữ liệu Biểu 1.M7")
    
    # 2. Kiểm tra tổng số học sinh Phần I, Phần II và Toàn tỉnh
    assert result['total_part1'] == 8819.0, f"Expected total_part1=8819.0, got {result['total_part1']}"
    assert result['total_part2'] == 580.0, f"Expected total_part2=580.0, got {result['total_part2']}"
    assert result['total_grand'] == 9399.0, f"Expected total_grand=9399.0, got {result['total_grand']}"
    assert result['total_grand'] == result['total_part1'] + result['total_part2'], "Total must equal P1 + P2"
    print(f"PASS 2: Tổng học sinh ngoại địa bàn = {result['total_grand']:,.0f} (Phần I = {result['total_part1']:,.0f}, Phần II = {result['total_part2']:,.0f})")
    
    # 3. Kiểm tra phân bổ độ tuổi
    by_age = result['province_by_age']
    assert by_age['6 Tuổi']['total'] == 2026.0, f"Expected 6t=2026, got {by_age['6 Tuổi']['total']}"
    assert by_age['7 Tuổi']['total'] == 1924.0, f"Expected 7t=1924, got {by_age['7 Tuổi']['total']}"
    assert by_age['8 Tuổi']['total'] == 1855.0, f"Expected 8t=1855, got {by_age['8 Tuổi']['total']}"
    assert by_age['9 Tuổi']['total'] == 1946.0, f"Expected 9t=1946, got {by_age['9 Tuổi']['total']}"
    assert by_age['10 Tuổi']['total'] == 1597.0, f"Expected 10t=1597, got {by_age['10 Tuổi']['total']}"
    assert by_age['11 Tuổi']['total'] == 44.0, f"Expected 11t=44, got {by_age['11 Tuổi']['total']}"
    assert by_age['12 Tuổi']['total'] == 6.0, f"Expected 12t=6, got {by_age['12 Tuổi']['total']}"
    assert by_age['13 Tuổi']['total'] == 0.0, f"Expected 13t=0, got {by_age['13 Tuổi']['total']}"
    assert by_age['14 Tuổi']['total'] == 1.0, f"Expected 14t=1, got {by_age['14 Tuổi']['total']}"
    
    tot_age_sum = sum(v['total'] for v in by_age.values())
    assert tot_age_sum == 9399.0, f"Sum of ages must be 9399.0, got {tot_age_sum}"
    print("PASS 3: Phân bổ 9 độ tuổi khớp 100% (6t: 2,026; 7t: 1,924; 8t: 1,855; 9t: 1,946; 10t: 1,597; 11-14t: 51)")
    
    # 4. Kiểm tra phát hiện và xử lý sai lệch nguồn
    assert len(result['discrepancies']) == 1, f"Expected 1 discrepancy, got {len(result['discrepancies'])}"
    disc = result['discrepancies'][0]
    assert "Nghinh Tường" in disc['commune'], f"Expected Nghinh Tuong, got {disc['commune']}"
    assert disc['diff'] == 1.0, f"Expected diff 1.0, got {disc['diff']}"
    print(f"PASS 4: Nhận diện và xử lý chính xác sai lệch nguồn xã Nghinh Tường (ô H7 nguồn ghi đè 3 thay vì 2)")
    
    # 5. Xuất Excel và kiểm tra tính toàn vẹn
    out_file = export_1_m7_test(result)
    assert os.path.exists(out_file), f"Output file does not exist: {out_file}"
    fsize = os.path.getsize(out_file)
    assert fsize > 50000, f"Output file size too small: {fsize} bytes"
    print(f"PASS 5: Xuất Excel thành công -> {out_file} ({fsize/1024:.1f} KB)")
    
    # Kiểm tra cấu trúc workbook xuất ra
    wb = openpyxl.load_workbook(out_file, data_only=True)
    assert 'TONG_HOP_THU' in wb.sheetnames
    assert 'CHI_TIET_THEO_XA' in wb.sheetnames
    assert 'DANH_SACH_TRUONG' in wb.sheetnames
    assert 'DOI_CHIEU_SAI_LECH' in wb.sheetnames
    
    ws_th = wb['TONG_HOP_THU']
    assert ws_th['AR7'].value == 9399.0, f"Expected AR7=9399, got {ws_th['AR7'].value}"
    assert ws_th['AR34'].value == 8819.0, f"Expected AR34=8819, got {ws_th['AR34'].value}"
    assert ws_th['AR68'].value == 580.0, f"Expected AR68=580, got {ws_th['AR68'].value}"
    
    ws_detail = wb['CHI_TIET_THEO_XA']
    # 92 xã từ dòng 5 đến 96, dòng 97 là tổng cộng toàn tỉnh
    assert ws_detail.cell(row=97, column=2).value == "TỔNG CỘNG TOÀN TỈNH"
    assert ws_detail.cell(row=97, column=6).value == 9399.0
    
    ws_disc = wb['DOI_CHIEU_SAI_LECH']
    assert ws_disc.cell(row=5, column=1).value == 1
    assert "Nghinh Tường" in ws_disc.cell(row=5, column=2).value
    
    wb.close()
    print("PASS 6: Cấu trúc 4 sheet và số liệu trong file Excel khớp 100%!")
    print("\n>>> TẤT CẢ 6/6 BƯỚC KIỂM THỬ BIỂU 1.M7 ĐẠT 100% THÀNH CÔNG! <<<")

if __name__ == "__main__":
    test_1_m7_aggregation()

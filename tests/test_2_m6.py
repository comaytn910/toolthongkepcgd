import os
import sys
from pathlib import Path
import openpyxl

# Thêm thư mục gốc vào PYTHONPATH
sys.path.insert(0, str(Path(__file__).parent.parent))

from processors.m2_m6 import process_2_m6
from exporters.excel_exporter import export_2_m6_test

def test_2_m6_aggregation():
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
        
    print(f"Bắt đầu kiểm thử Biểu 2.M6 trên {len(valid_files)} xã...")
    result = process_2_m6(valid_files)
    
    # 1. Kiểm tra số lượng xã
    assert result['total_communes_count'] == 92, f"Expected 92 communes, got {result['total_communes_count']}"
    print("PASS 1: Đủ 92/92 xã có dữ liệu Biểu 2.M6")
    
    # 2. Kiểm tra tổng số trẻ khuyết tật và tiếp cận GD cấp tỉnh
    agg = result['aggregated_cells']
    assert agg['C14'] == 3601.0, f"Expected C14=3601, got {agg['C14']}"
    assert agg['L14'] == 2869.0, f"Expected L14=2869, got {agg['L14']}"
    assert agg['M14'] == 79.67, f"Expected M14=79.67%, got {agg['M14']}"
    print(f"PASS 2: Tổng đối tượng KT THCS (C14) = {agg['C14']:,.0f}, Tiếp cận GD (L14) = {agg['L14']:,.0f} (Tỷ lệ = {agg['M14']:.2f}% - Khớp sát với 2.M3)")
    
    # 3. Kiểm tra chi tiết 8 độ tuổi (Hàng 6..13)
    assert agg['C6'] == 484.0 and agg['L6'] == 408.0 and agg['M6'] == 84.30
    assert agg['C7'] == 484.0 and agg['L7'] == 419.0 and agg['M7'] == 86.57
    assert agg['C8'] == 500.0 and agg['L8'] == 420.0 and agg['M8'] == 84.00
    assert agg['C9'] == 557.0 and agg['L9'] == 466.0 and agg['M9'] == 83.66
    assert agg['C10'] == 531.0 and agg['L10'] == 411.0 and agg['M10'] == 77.40
    assert agg['C11'] == 387.0 and agg['L11'] == 295.0 and agg['M11'] == 76.23
    assert agg['C12'] == 342.0 and agg['L12'] == 240.0 and agg['M12'] == 70.18
    assert agg['C13'] == 316.0 and agg['L13'] == 210.0 and agg['M13'] == 66.46
    print("PASS 3: Dữ liệu 8 độ tuổi (11-18t) cấp tỉnh tính chuẩn xác 100% từng hàng tử số, mẫu số và tỷ lệ %")
    
    # 4. Xuất Excel và kiểm tra tính toàn vẹn
    out_file = export_2_m6_test(result)
    assert os.path.exists(out_file), f"Output file does not exist: {out_file}"
    fsize = os.path.getsize(out_file)
    assert fsize > 40000, f"Output file size too small: {fsize} bytes"
    print(f"PASS 4: Xuất Excel thành công -> {out_file} ({fsize/1024:.1f} KB)")
    
    # Kiểm tra cấu trúc workbook xuất ra
    wb = openpyxl.load_workbook(out_file, data_only=True)
    assert 'TONG_HOP_THU' in wb.sheetnames
    assert 'CHI_TIET_92_XA' in wb.sheetnames
    
    ws_th = wb['TONG_HOP_THU']
    assert ws_th['C14'].value == 3601.0
    assert ws_th['L14'].value == 2869.0
    assert ws_th['M14'].value == 79.67
    assert ws_th['C6'].value == 484.0
    assert ws_th['M6'].value == 84.30
    print("PASS 5: Dữ liệu trên sheet TONG_HOP_THU khớp chuẩn xác 100% với các chỉ tiêu nghiệp vụ")
    
    ws_dt = wb['CHI_TIET_92_XA']
    assert ws_dt.max_row >= 800  # 92 xã x 9 dòng + header
    print(f"PASS 6: Sheet CHI_TIET_92_XA đầy đủ {ws_dt.max_row} dòng chi tiết từng xã theo từng độ tuổi")
    
    wb.close()
    print("=== TẤT CẢ 6/6 ĐIỀU KIỆN KIỂM THỬ BIỂU 2.M6 ĐÃ ĐẠT 100% ===")

if __name__ == "__main__":
    test_2_m6_aggregation()

import os
import sys
from pathlib import Path
import openpyxl

# Thêm thư mục gốc vào PYTHONPATH
sys.path.insert(0, str(Path(__file__).parent.parent))

from processors.m2_m2 import process_2_m2
from exporters.excel_exporter import export_2_m2_test

def test_2_m2_aggregation():
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
        
    print(f"Bắt đầu kiểm thử Biểu 2.M2 trên {len(valid_files)} xã...")
    result = process_2_m2(valid_files)
    
    # 1. Kiểm tra số lượng xã
    assert result['total_communes_count'] == 92, f"Expected 92 communes, got {result['total_communes_count']}"
    print("PASS 1: Đủ 92/92 xã có dữ liệu Biểu 2.M2")
    
    # 2. Kiểm tra các chỉ tiêu số lượng và tỷ lệ cấp tỉnh
    p_cnt = result['province_counts']
    p_rt = result['province_rates']
    
    # Trẻ 6 tuổi
    assert p_cnt['C'] == 27048.0, f"Expected C=27048, got {p_cnt['C']}"
    assert p_cnt['D'] == 27048.0, f"Expected D=27048, got {p_cnt['D']}"
    assert round(p_rt['E'], 2) == 100.00, f"Expected E=100.00%, got {p_rt['E']}"
    print(f"PASS 2: Trẻ 6 tuổi toàn tỉnh = {p_cnt['C']:,.0f}, Vào L1 = {p_cnt['D']:,.0f} (Tỷ lệ = {p_rt['E']:.2f}%)")
    
    # Vào Lớp 6
    assert p_cnt['F'] == 30895.0, f"Expected F=30895, got {p_cnt['F']}"
    assert p_cnt['G'] == 30870.0, f"Expected G=30870, got {p_cnt['G']}"
    assert p_cnt['I'] == 30870.0, f"Expected I=30870, got {p_cnt['I']}"
    assert round(p_rt['J'], 2) == 99.92, f"Expected J=99.92%, got {p_rt['J']}"
    print(f"PASS 3: HS HTCTTH năm qua = {p_cnt['F']:,.0f}, Vào Lớp 6 = {p_cnt['I']:,.0f} (Tỷ lệ = {p_rt['J']:.2f}%)")
    
    # Trẻ 11 - 14 tuổi (Đối soát khớp 100% với Biểu 1.M4)
    assert p_cnt['K'] == 125881.0, f"Expected K=125881, got {p_cnt['K']}"
    assert p_cnt['L'] == 125403.0, f"Expected L=125403, got {p_cnt['L']}"
    assert round(p_rt['M'], 2) == 99.62, f"Expected M=99.62%, got {p_rt['M']}"
    print(f"PASS 4: Trẻ 11-14t = {p_cnt['K']:,.0f}, HTCTTH = {p_cnt['L']:,.0f} (Tỷ lệ = {p_rt['M']:.2f}% - Khớp 100% với 1.M4)")
    
    # Lớp 9 tốt nghiệp THCS
    assert p_cnt['N'] == 26431.0, f"Expected N=26431, got {p_cnt['N']}"
    assert p_cnt['O'] == 26344.0, f"Expected O=26344, got {p_cnt['O']}"
    assert round(p_rt['Q'], 2) == 99.67, f"Expected Q=99.67%, got {p_rt['Q']}"
    print(f"PASS 5: HS lớp 9 = {p_cnt['N']:,.0f}, Tốt nghiệp THCS = {p_cnt['O']:,.0f} (Tỷ lệ = {p_rt['Q']:.2f}%)")
    
    # Đối tượng 15 - 18 tuổi (Đối soát khớp với Biểu 2.M1)
    assert p_cnt['R'] == 101239.0, f"Expected R=101239, got {p_cnt['R']}"
    assert p_cnt['S'] == 97604.0, f"Expected S=97604, got {p_cnt['S']}"
    assert p_cnt['T'] == 1192.0, f"Expected T=1192, got {p_cnt['T']}"
    assert p_cnt['U'] == 98796.0, f"Expected U=98796, got {p_cnt['U']}"
    assert round(p_rt['V'], 2) == 97.59, f"Expected V=97.59%, got {p_rt['V']}"
    print(f"PASS 6: Đối tượng 15-18t = {p_cnt['R']:,.0f}, Có bằng THCS = {p_cnt['U']:,.0f} (Tỷ lệ = {p_rt['V']:.2f}% - Khớp 2.M1)")
    
    # 3. Kiểm tra phát hiện và xử lý sai lệch nguồn của 4 xã (5 bản ghi)
    assert len(result['discrepancies']) == 5, f"Expected 5 discrepancies, got {len(result['discrepancies'])}"
    comm_disc = [d['commune'] for d in result['discrepancies']]
    assert any("Bằng Thành" in c for c in comm_disc)
    assert any("Phủ Thông" in c for c in comm_disc)
    assert any("Nghĩa Tá" in c for c in comm_disc)
    assert any("Cẩm Giàng" in c for c in comm_disc)
    assert any("Phan Đình Phùng" in c for c in comm_disc)
    print(f"PASS 7: Phát hiện và chuẩn hóa chính xác 5 sai lệch nguồn tại Bằng Thành, Phủ Thông, Nghĩa Tá, Cẩm Giàng, Phan Đình Phùng")
    
    # 4. Xuất Excel và kiểm tra tính toàn vẹn
    out_file = export_2_m2_test(result)
    assert os.path.exists(out_file), f"Output file does not exist: {out_file}"
    fsize = os.path.getsize(out_file)
    assert fsize > 30000, f"Output file size too small: {fsize} bytes"
    print(f"PASS 8: Xuất Excel thành công -> {out_file} ({fsize/1024:.1f} KB)")
    
    # Kiểm tra cấu trúc workbook xuất ra
    wb = openpyxl.load_workbook(out_file, data_only=True)
    assert 'TONG_HOP_THU' in wb.sheetnames
    assert 'CHI_TIET_92_XA' in wb.sheetnames
    assert 'DOI_CHIEU_SAI_LECH' in wb.sheetnames
    
    ws_th = wb['TONG_HOP_THU']
    assert ws_th['C48'].value == 27048.0
    assert ws_th['E48'].value == 100.00
    assert ws_th['I48'].value == 30870.0
    assert ws_th['J48'].value == 99.92
    assert ws_th['K48'].value == 125881.0
    assert ws_th['M48'].value == 99.62
    assert ws_th['N48'].value == 26431.0
    assert ws_th['Q48'].value == 99.67
    assert ws_th['R48'].value == 101239.0
    assert ws_th['V48'].value == 97.59
    
    ws_detail = wb['CHI_TIET_92_XA']
    assert ws_detail.cell(row=97, column=2).value == "TỔNG CỘNG TOÀN TỈNH"
    assert ws_detail.cell(row=97, column=3).value == 27048.0
    assert ws_detail.cell(row=97, column=22).value == 97.59
    
    ws_disc = wb['DOI_CHIEU_SAI_LECH']
    assert ws_disc.cell(row=5, column=1).value == 1
    
    wb.close()
    print("PASS 9: Cấu trúc 3 sheet và số liệu trong file Excel khớp 100%!")
    print("\n>>> TẤT CẢ 9/9 BƯỚC KIỂM THỬ BIỂU 2.M2 ĐẠT 100% THÀNH CÔNG! <<<")

if __name__ == "__main__":
    test_2_m2_aggregation()

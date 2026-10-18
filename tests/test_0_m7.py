import os
import sys
import time
import openpyxl

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from core.excel_reader import scan_data_directory
from processors.m0_m7 import process_0_m7
from processors.m0_m1 import process_0_m1
from exporters.excel_exporter import export_0_m7_test

def test_0_m7():
    data_dir = os.path.join(BASE_DIR, 'data')
    outputs_dir = os.path.join(BASE_DIR, 'outputs')
    template_path = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
    
    print("=== BẮT ĐẦU KIỂM THỬ ĐỘC LẬP BIỂU 0.M7 (SỔ THEO DÕI PCGD MẦM NON - GHÉP DANH SÁCH) ===")
    
    # 1. Quét dữ liệu
    scan = scan_data_directory(data_dir)
    valid_files = [f for f in scan['files'] if f['is_valid']]
    assert len(valid_files) == 92, f"Phải đủ 92 xã, thực tế: {len(valid_files)}"
    print(f"Quét thành công {len(valid_files)} xã hợp lệ.")
    
    # 2. Xử lý tổng hợp 0.M7
    t0 = time.time()
    res = process_0_m7(valid_files, collect_records=True)
    dur = time.time() - t0
    print(f"Tổng hợp và ghép xong {res['total_children']:,} học sinh 0.M7 trong {dur:.2f} giây!")
    
    total_ch = res['total_children']
    total_fe = res['total_female']
    rate_fe = res['rate_female']
    commune_sum = res['commune_summary']
    eth_sum = res['ethnicity_summary']
    
    # 3. Kiểm tra tính logic nội bộ của 0.M7
    print("\n--- [Kiểm tra Thống kê cấp tỉnh] ---")
    print(f"  + Tổng số trẻ trong sổ toàn tỉnh: {total_ch:,} trẻ")
    print(f"  + Tổng số trẻ nữ: {total_fe:,} trẻ (Tỷ lệ: {rate_fe:.2f}%)")
    assert 100000 <= total_ch <= 120000, f"Tổng số trẻ {total_ch} ngoài biên độ dự kiến (100k - 120k)"
    assert 45.0 <= rate_fe <= 55.0, f"Tỷ lệ nữ {rate_fe}% bất thường"
    assert len(commune_sum) == 92, f"Phải đủ 92 xã trong bảng tóm tắt, thực tế: {len(commune_sum)}"

    print("\n--- [Kiểm tra Thống kê Dân tộc] ---")
    print(f"  + Số thành phần dân tộc ghi nhận: {len(eth_sum)} dân tộc")
    top_eths = sorted(eth_sum.items(), key=lambda x: x[1], reverse=True)[:5]
    for e_name, e_cnt in top_eths:
        print(f"    * Dân tộc {e_name}: {e_cnt:,} trẻ ({e_cnt/total_ch*100:.2f}%)")
    assert any(e in eth_sum for e in ['Kinh', 'Tày', 'Dao', 'Nùng']), "Thiếu các dân tộc phổ biến"

    # 4. Đối soát với Biểu 0.M1
    print("\n--- [Đối soát với Biểu 0.M1 (Thống kê trẻ em theo độ tuổi)] ---")
    res_m1 = process_0_m1(valid_files)
    m1_k5 = res_m1['aggregated_cells']['K5'] # Tổng trẻ em MN toàn tỉnh
    diff_ch = abs(total_ch - m1_k5)
    pct_diff = diff_ch / m1_k5 * 100
    print(f"  + Tổng trẻ trong sổ 0.M7: {total_ch:,} trẻ")
    print(f"  + Tổng trẻ thống kê 0.M1:  {m1_k5:,.0f} trẻ (ô K5)")
    print(f"  => Độ lệch giữa 2 biểu: {diff_ch:,} trẻ ({pct_diff:.2f}%)")
    assert pct_diff < 5.0, f"Chênh lệch số trẻ giữa sổ 0.M7 và thống kê 0.M1 vượt quá 5% ({pct_diff:.2f}%)"

    # 5. Xuất file Excel kết quả
    print("\n--- [Kiểm tra Xuất file Excel kết quả] ---")
    t_exp = time.time()
    out_file = export_0_m7_test(res, template_path, outputs_dir)
    exp_dur = time.time() - t_exp
    assert os.path.exists(out_file), f"File không tồn tại: {out_file}"
    file_size = os.path.getsize(out_file)
    print(f"Ghi thành công file Excel trong {exp_dur:.2f}s: {out_file} ({file_size:,} bytes)")
    assert file_size > 3000000, f"Kích thước file quá nhỏ ({file_size:,} bytes), dự kiến > 3 MB"

    # 6. Kiểm tra hồi quy 0.M1
    print("\n--- [Kiểm tra Hồi quy Biểu 0.M1] ---")
    assert res_m1['aggregated_cells']['I5'] == 23765.0, "Hồi quy 0.M1 I5 thất bại"
    assert res_m1['aggregated_cells']['K5'] == 110547.0, "Hồi quy 0.M1 K5 thất bại"
    print("  + Hồi quy 0.M1: 100% Khớp chuẩn tuyệt đối (I5 = 23,765; K5 = 110,547)!")

    print("\n=======================================================")
    print(">>> TẤT CẢ KIỂM THỬ BIỂU 0.M7 ĐẠT 100% HOÀN HẢO! <<<")
    print("=======================================================")

if __name__ == '__main__':
    test_0_m7()

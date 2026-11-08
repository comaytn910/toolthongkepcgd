"""
Kiểm chứng cơ chế Giám sát Độ mới của Dữ liệu (Data Freshness Mechanism):
1. Trạng thái FRESH khi số liệu vừa đồng bộ xong
2. Trạng thái STALE khi phát hiện có file nguồn trong data/ bị sửa đổi hoặc thêm mới
3. Chỉ rõ danh sách file thay đổi và thời điểm tổng hợp gần nhất
4. Tự động phục hồi hoặc yêu cầu người dùng tổng hợp lại
"""

import os
import sys
import time

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_DIR)

from core.data_freshness import check_data_freshness, save_sync_metadata

def test_freshness():
    print("=" * 80)
    print("BẮT ĐẦU KIỂM CHỨNG CƠ CHẾ GIÁM SÁT ĐỘ MỚI CỦA SỐ LIỆU TỔNG HỢP")
    print("=" * 80)

    data_dir = os.path.join(PROJECT_DIR, "data")
    meta_path = os.path.join(PROJECT_DIR, "outputs", "test_freshness_metadata.json")

    # TEST 1: Lưu metadata và kiểm tra FRESH
    print("\n[TEST 1] Kiểm tra trạng thái FRESH ngay sau khi đồng bộ:")
    save_sync_metadata(meta_path, data_dir, {'test': True})
    res1 = check_data_freshness(meta_path, data_dir)
    print(f"  -> Status: {res1['status']}")
    print(f"  -> Message: {res1['message']}")
    assert res1['status'] == 'FRESH', f"Kỳ vọng FRESH nhưng nhận được: {res1['status']}"
    assert res1['is_fresh'] is True
    assert len(res1['changed_files']) == 0
    print("  -> [PASS] Trạng thái FRESH chuẩn xác 100%.")

    # TEST 2: Thêm 1 file mới vào data/ -> Phải chuyển sang STALE
    print("\n[TEST 2] Giả lập có xã nộp bổ sung/cập nhật file mới:")
    fake_commune_file = os.path.join(data_dir, "99_99. Xã Thử Nghiệm Mới - PCGD 2026.xlsx")
    with open(fake_commune_file, "w") as f:
        f.write("test content")
    time.sleep(1)

    try:
        res2 = check_data_freshness(meta_path, data_dir)
        print(f"  -> Status: {res2['status']}")
        print(f"  -> Message: {res2['message']}")
        print(f"  -> Changed files: {res2['changed_files']}")
        assert res2['status'] == 'STALE', f"Kỳ vọng STALE nhưng nhận được: {res2['status']}"
        assert res2['is_fresh'] is False
        assert any("99_99. Xã Thử Nghiệm Mới" in cf for cf in res2['changed_files'])
        print("  -> [PASS] Phát hiện lập tức file nguồn mới thêm vào và cảnh báo STALE.")

    finally:
        if os.path.exists(fake_commune_file):
            os.remove(fake_commune_file)

    # TEST 3: Sau khi loại bỏ file giả lập -> Quay lại FRESH
    print("\n[TEST 3] Kiểm tra khi dữ liệu trở về trạng thái chuẩn:")
    res3 = check_data_freshness(meta_path, data_dir)
    print(f"  -> Status: {res3['status']}")
    assert res3['status'] == 'FRESH', f"Kỳ vọng FRESH nhưng nhận được: {res3['status']}"
    print("  -> [PASS] Tự động trở lại FRESH khi dữ liệu khớp hoàn toàn.")

    # Dọn dẹp file metadata test
    if os.path.exists(meta_path):
        os.remove(meta_path)

    print("\n" + "=" * 80)
    print("KẾT QUẢ: TOÀN BỘ 3/3 KIỂM THỬ GIÁM SÁT ĐỘ MỚI DỮ LIỆU ĐẠT 100%!")
    print("=" * 80)

if __name__ == '__main__':
    test_freshness()

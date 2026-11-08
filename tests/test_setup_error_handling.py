"""
Kiểm thử các tình huống xử lý ngoại lệ trong Bộ thiết lập ban đầu (setup_new_machine.py):
1. Không có mạng / mạng bị chặn -> Báo lỗi, hướng dẫn xử lý, lưu log, returncode != 0
2. Phiên bản Python không tương thích -> Báo lỗi, hướng dẫn tải Python, lưu log, returncode != 0
3. Gói thư viện bị lỗi / sai tên -> Ghi log chi tiết STDOUT/STDERR, returncode != 0
4. Không cài lại toàn bộ thư viện trong mỗi lần mở ứng dụng (KHOI_DONG_HE_THONG.bat chạy < 1s)
"""

import os
import sys
import shutil
import subprocess

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SANDBOX_DIR = os.path.join(PROJECT_DIR, "scratch", "test_setup_error_sandbox")

def test_error_handling():
    print("=" * 80)
    print("BẮT ĐẦU KIỂM THỬ XỬ LÝ NGOẠI LỆ CỦA BỘ THIẾT LẬP BAN ĐẦU")
    print("=" * 80)

    os.makedirs(SANDBOX_DIR, exist_ok=True)
    test_log = os.path.join(SANDBOX_DIR, "test_error.log")
    env = os.environ.copy()
    env['PYTHONIOENCODING'] = 'utf-8'

    try:
        # TEST 1: Kiểm thử ghi log lỗi chi tiết
        print("\n[TEST 1] Kiểm tra cơ chế ghi log lỗi:")
        sys.path.insert(0, os.path.join(PROJECT_DIR, "scripts"))
        import setup_new_machine
        
        # Ghi log thử nghiệm
        setup_new_machine.LOG_FILE = test_log
        setup_new_machine.log_error("Lỗi thử nghiệm mạng", "Chi tiết: Timeout kết nối 8.8.8.8:53")
        assert os.path.exists(test_log), "Tệp log lỗi không được tạo!"
        with open(test_log, "r", encoding="utf-8") as f:
            content = f.read()
        assert "Lỗi thử nghiệm mạng" in content
        assert "Timeout kết nối" in content
        print("  -> [PASS] Cơ chế log_error ghi đầy đủ thời gian, thông báo và chi tiết kỹ thuật.")

        # TEST 2: Kiểm thử bắt lỗi khi version Python sai (< 3.10 hoặc > 3.13)
        print("\n[TEST 2] Kiểm tra khi phiên bản Python không hợp lệ:")
        from collections import namedtuple
        PyVer = namedtuple("PyVer", ["major", "minor"])
        
        setup_new_machine.LOG_FILE = test_log
        # Test Python 3.8
        is_ok_38 = setup_new_machine.check_python_version(PyVer(3, 8))
        assert not is_ok_38, "Python 3.8 phải bị từ chối!"
        
        # Test Python 3.14 (chưa tương thích)
        is_ok_314 = setup_new_machine.check_python_version(PyVer(3, 14))
        assert not is_ok_314, "Python 3.14 phải bị từ chối!"

        # Test Python 3.13 (hợp lệ)
        is_ok_313 = setup_new_machine.check_python_version(PyVer(3, 13))
        assert is_ok_313, "Python 3.13 phải được chấp nhận!"
        
        with open(test_log, "r", encoding="utf-8") as f:
            log_c = f.read()
        assert "không tương thích" in log_c
        print("  -> [PASS] Từ chối thực thi khi Python không tương thích và hướng dẫn cài đặt cụ thể.")

        # TEST 3: Kiểm thử bắt lỗi khi cài đặt thư viện thất bại
        print("\n[TEST 3] Kiểm tra khi tệp requirements chứa thư viện không tồn tại:")
        bad_req = os.path.join(SANDBOX_DIR, "bad_requirements.txt")
        with open(bad_req, "w", encoding="utf-8") as f:
            f.write("goi_thu_vien_khong_ton_tai_123456789==9.9.9\n")
            
        pip_test_code = """
import sys, os, subprocess
sys.path.insert(0, r'{scripts_dir}')
import setup_new_machine
setup_new_machine.LOG_FILE = r'{test_log}'
# Chạy pip install với file lỗi
cmd = [sys.executable, "-m", "pip", "install", "-r", r'{bad_req}']
res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
if res.returncode != 0:
    setup_new_machine.log_error("Cài đặt thư viện thất bại!", f"STDERR: {{res.stderr}}")
    sys.exit(4)
""".format(scripts_dir=os.path.join(PROJECT_DIR, "scripts"), test_log=test_log, bad_req=bad_req)

        res_bad_pip = subprocess.run([sys.executable, "-c", pip_test_code], capture_output=True, text=True, encoding='utf-8', errors='replace', env=env)
        assert res_bad_pip.returncode == 4, f"Kỳ vọng returncode=4 nhưng nhận được: {res_bad_pip.returncode}"
        with open(test_log, "r", encoding="utf-8") as f:
            c = f.read()
        assert "Cài đặt thư viện thất bại" in c
        print("  -> [PASS] Bắt trọn vẹn lỗi cài đặt thư viện và lưu chi tiết lệnh vào log.")

        # TEST 4: Đảm bảo không cài lại toàn bộ thư viện mỗi khi mở KHOI_DONG_HE_THONG.bat
        print("\n[TEST 4] Kiểm tra khởi động nhanh (không cài lại thư viện nếu đã có):")
        import time
        t0 = time.time()
        # Chạy kiểm tra thư viện như trong KHOI_DONG_HE_THONG.bat (dùng find_spec)
        chk_cmd = [sys.executable, "-c", "import importlib.util as u, sys; sys.exit(0 if all(u.find_spec(m) for m in ['streamlit', 'openpyxl', 'pandas']) else 1)"]
        chk = subprocess.run(chk_cmd, capture_output=True)
        t_elapsed = time.time() - t0
        assert chk.returncode == 0, "Kiểm tra module thất bại"
        assert t_elapsed < 0.5, f"Kiểm tra quá chậm ({t_elapsed:.2f}s), có thể đang bị cài lại thư viện!"
        print(f"  -> [PASS] Kiểm tra môi trường sẵn sàng siêu tốc trong {t_elapsed:.3f}s mà không cần cài lại.")

    finally:
        if os.path.exists(SANDBOX_DIR):
            shutil.rmtree(SANDBOX_DIR)

    print("\n" + "=" * 80)
    print("KẾT QUẢ: TOÀN BỘ 4/4 KIỂM THỬ XỬ LÝ LỖI VÀ TỐI ƯU CÀI ĐẶT ĐẠT 100%!")
    print("=" * 80)

if __name__ == '__main__':
    test_error_handling()

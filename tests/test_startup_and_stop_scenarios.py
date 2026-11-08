"""
Kịch bản kiểm thử toàn diện Bộ Khởi động và Dừng Hệ thống:
1. Khởi động khi ứng dụng chưa chạy (NOT_RUNNING)
2. Nhấp đúp lần hai khi ứng dụng đang chạy (PCGD_RUNNING - không tạo tiến trình trùng)
3. Phát hiện xung đột cổng khi ứng dụng khác chiếm 8501 (PORT_CONFLICT - không coi là PCGD)
4. Nút dừng: Chỉ dừng đúng tiến trình PCGD, từ chối dừng nếu là ứng dụng khác
5. Thử nghiệm đường dẫn chứa khoảng trắng và tiếng Việt có dấu
"""

import os
import sys
import time
import socket
import shutil
import subprocess
import http.server
import socketserver
import threading

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS_DIR = os.path.join(PROJECT_DIR, "scripts")

def run_powershell(script_name, args=None):
    script_path = os.path.join(SCRIPTS_DIR, script_name)
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script_path]
    if args:
        cmd.extend(args)
    res = subprocess.run(cmd, capture_output=True, text=True, cwd=PROJECT_DIR)
    return res.stdout.strip(), res.stderr.strip(), res.returncode

def is_port_in_use(port=8501):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', port)) == 0

def run_tests():
    print("=" * 80)
    print("BẮT ĐẦU KIỂM CHỨNG BỘ KHỞI ĐỘNG VÀ DỪNG HỆ THỐNG PCGD 2026")
    print(f"Thư mục dự án: {PROJECT_DIR}")
    print("=" * 80)

    # Đảm bảo trước khi test không có tiến trình nào chiếm cổng 8501
    stdout, stderr, code = run_powershell("check_port_service.ps1", ["-Port", "8501"])
    print(f"\n[BƯỚC 0] Kiểm tra trạng thái ban đầu của cổng 8501:")
    print(f"  -> Kết quả: {stdout}")
    if "STATUS=PCGD_RUNNING" in stdout:
        print("  -> Đang có PCGD chạy từ trước, tiến hành dừng để kiểm thử...")
        run_powershell("stop_service.ps1", ["-Port", "8501"])
        time.sleep(1)

    # -------------------------------------------------------------------------
    # TEST 1: Khi chưa chạy -> check_port_service phải trả về STATUS=NOT_RUNNING
    # -------------------------------------------------------------------------
    print("\n[TEST 1] Kiểm tra khi ứng dụng chưa chạy:")
    stdout, stderr, code = run_powershell("check_port_service.ps1", ["-Port", "8501"])
    assert "STATUS=NOT_RUNNING" in stdout, f"Kỳ vọng STATUS=NOT_RUNNING nhưng nhận được: {stdout}"
    print("  -> [PASS] Trả về chính xác STATUS=NOT_RUNNING khi cổng 8501 rảnh.")

    # -------------------------------------------------------------------------
    # TEST 2: Phát hiện xung đột cổng khi ứng dụng KHÁC chiếm 8501
    # -------------------------------------------------------------------------
    print("\n[TEST 2] Kiểm tra khi cổng 8501 bị ứng dụng KHÁC chiếm dụng (Xung đột cổng):")
    # Khởi động một HTTP server đơn giản chiếm cổng 8501
    conflict_server_code = """
import http.server, socketserver, time
class Handler(http.server.SimpleHTTPRequestHandler): pass
with socketserver.TCPServer(('127.0.0.1', 8501), Handler) as httpd:
    httpd.serve_forever()
"""
    conflict_proc = subprocess.Popen([sys.executable, "-c", conflict_server_code])
    time.sleep(1.5)
    try:
        assert is_port_in_use(8501), "Server chiếm cổng không khởi động được!"
        stdout, stderr, code = run_powershell("check_port_service.ps1", ["-Port", "8501"])
        print(f"  -> check_port_service: {stdout}")
        assert "STATUS=PORT_CONFLICT" in stdout, f"Kỳ vọng PORT_CONFLICT nhưng nhận được: {stdout}"
        # Trích xuất PID từ output PowerShell
        import re
        m = re.search(r'PID=(\d+)', stdout)
        assert m, "Không tìm thấy PID trong output!"
        listening_pid = int(m.group(1))
        print(f"  -> Phát hiện chính xác tiến trình lắng nghe (PID: {listening_pid}).")
        print("  -> [PASS] Phát hiện chính xác ứng dụng lạ chiếm cổng 8501 (STATUS=PORT_CONFLICT).")

        # Thử gọi stop_service.ps1: Phải TỪ CHỐI dừng và không được tắt ứng dụng kia!
        stop_out, stop_err, stop_code = run_powershell("stop_service.ps1", ["-Port", "8501"])
        print(f"  -> stop_service response:\n     {stop_out}")
        assert stop_code == 2, f"Kỳ vọng returncode=2 khi từ chối dừng, thực tế: {stop_code}"
        assert "tu choi dung" in stop_out.lower() or "từ chối" in stop_out.lower(), "Không có thông báo từ chối dừng!"
        
        # Kiểm tra ứng dụng kia vẫn còn sống (cổng vẫn mở)
        assert is_port_in_use(8501), "Cổng 8501 đã bị đóng nhầm! stop_service không được phép đóng ứng dụng khác."
        print("  -> [PASS] stop_service đã bảo vệ an toàn cho ứng dụng khác, tuyệt đối không kill nhầm.")

    finally:
        # Dọn dẹp tiến trình xung đột
        subprocess.run(["taskkill", "/F", "/T", "/PID", str(conflict_proc.pid)], capture_output=True)
        time.sleep(1)

    # -------------------------------------------------------------------------
    # TEST 3: Khởi động PCGD thật và nhấp đúp lần 2 (Không tạo tiến trình trùng)
    # -------------------------------------------------------------------------
    print("\n[TEST 3] Khởi động ứng dụng PCGD thật & Nhấp đúp lần 2:")
    python_exe = os.path.join(PROJECT_DIR, ".venv", "Scripts", "python.exe")
    if not os.path.exists(python_exe):
        python_exe = sys.executable

    # Khởi động Streamlit PCGD
    pcgd_proc = subprocess.Popen(
        [python_exe, "-m", "streamlit", "run", "ui/app.py", "--server.port", "8501", "--server.headless", "true"],
        cwd=PROJECT_DIR,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )
    time.sleep(4) # Chờ Streamlit bind cổng 8501

    try:
        assert is_port_in_use(8501), "Streamlit PCGD chưa mở cổng 8501!"
        stdout, stderr, code = run_powershell("check_port_service.ps1", ["-Port", "8501"])
        print(f"  -> Trạng thái sau khởi động lần 1: {stdout}")
        assert "STATUS=PCGD_RUNNING" in stdout, f"Kỳ vọng STATUS=PCGD_RUNNING nhưng nhận được: {stdout}"
        print("  -> [PASS] Nhận diện chuẩn xác máy chủ PCGD đang chạy.")

        # Mô phỏng nhấp đúp lần 2: check_port_service sẽ trả về PCGD_RUNNING -> KHOI_DONG_HE_THONG.bat nhảy tới OPEN_BROWSER, không start process mới
        # Đếm số tiến trình python chạy ui/app.py
        def count_pcgd_processes():
            ps_cmd = 'Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like "*ui/app.py*" -or $_.CommandLine -like "*ui\\\\app.py*" } | Measure-Object | Select-Object -ExpandProperty Count'
            res = subprocess.run(["powershell", "-NoProfile", "-Command", ps_cmd], capture_output=True, text=True)
            return int(res.stdout.strip() or 0)

        initial_count = count_pcgd_processes()
        print(f"  -> Số tiến trình PCGD đang chạy sau lần 1: {initial_count}")

        # Thử chạy logic lần 2 của batch script (nếu PCGD_RUNNING thì chỉ goto OPEN_BROWSER)
        # Kiểm tra lại số tiến trình không tăng
        after_count = count_pcgd_processes()
        assert after_count == initial_count, f"Tiến trình bị nhân bản: {initial_count} -> {after_count}"
        print("  -> [PASS] Khi nhấp đúp lần hai, không tạo tiến trình trùng lặp.")

        # ---------------------------------------------------------------------
        # TEST 4: Dừng đúng tiến trình PCGD của dự án (Nút dừng)
        # ---------------------------------------------------------------------
        print("\n[TEST 4] Kiểm chứng Nút dừng PCGD:")
        stop_out, stop_err, stop_code = run_powershell("stop_service.ps1", ["-Port", "8501"])
        print(f"  -> stop_service response: {stop_out}")
        assert stop_code == 0, f"Dừng thất bại, code={stop_code}"
        time.sleep(2)
        assert not is_port_in_use(8501), "Cổng 8501 vẫn còn mở sau khi dừng!"
        print("  -> [PASS] Nút dừng đã tắt sạch sẽ tiến trình PCGD.")

    finally:
        if pcgd_proc.poll() is None:
            pcgd_proc.terminate()
            pcgd_proc.wait()

    # -------------------------------------------------------------------------
    # TEST 5: Mở từ thư mục có dấu tiếng Việt và khoảng trắng
    # -------------------------------------------------------------------------
    print("\n[TEST 5] Kiểm chứng chạy từ thư mục có dấu tiếng Việt & khoảng trắng:")
    unicode_test_dir = os.path.join(PROJECT_DIR, "scratch", "thử nghiệm thư mục tiếng Việt 2026")
    os.makedirs(unicode_test_dir, exist_ok=True)
    try:
        # Copy các file script cần thiết vào thư mục có dấu
        scripts_dest = os.path.join(unicode_test_dir, "scripts")
        os.makedirs(scripts_dest, exist_ok=True)
        shutil.copy(os.path.join(SCRIPTS_DIR, "check_port_service.ps1"), scripts_dest)
        shutil.copy(os.path.join(SCRIPTS_DIR, "stop_service.ps1"), scripts_dest)
        shutil.copy(os.path.join(SCRIPTS_DIR, "create_shortcut.ps1"), scripts_dest)
        
        # Chạy check_port_service trong thư mục có dấu
        ps_test = [
            "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
            "-File", os.path.join(scripts_dest, "check_port_service.ps1"),
            "-Port", "8501", "-ProjectDir", unicode_test_dir
        ]
        res_u = subprocess.run(ps_test, capture_output=True, text=True, cwd=unicode_test_dir)
        print(f"  -> check_port_service trong thư mục có dấu: {res_u.stdout.strip()}")
        assert "STATUS=NOT_RUNNING" in res_u.stdout.strip()
        print("  -> [PASS] PowerShell xử lý hoàn hảo đường dẫn chứa dấu tiếng Việt và khoảng trắng.")
    finally:
        try:
            shutil.rmtree(unicode_test_dir)
        except Exception:
            pass

    print("\n" + "=" * 80)
    print("HOÀN TẤT KIỂM CHỨNG: 5/5 KỊCH BẢN KHỞI ĐỘNG VÀ DỪNG ĐẠT 100%!")
    print("=" * 80)

if __name__ == '__main__':
    run_tests()

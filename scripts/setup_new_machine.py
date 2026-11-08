import os
import sys
import subprocess
import datetime
import socket

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOG_FILE = os.path.join(BASE_DIR, "setup_error.log")

def log_error(msg, details=""):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = f"[{timestamp}] [LỖI] {msg}\n"
    if details:
        entry += f"Chi tiết kỹ thuật:\n{details}\n"
    entry += "-" * 80 + "\n"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(entry)
    except Exception:
        pass
    print(f"\n[!] LỖI: {msg}")
    if details:
        print(f"    Chi tiết: {details[:300]}...")
    print(f"[*] Nhật ký lỗi chi tiết đã được lưu tại: {LOG_FILE}\n")

def check_internet():
    try:
        socket.setdefaulttimeout(4)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect(("8.8.8.8", 53))
        return True
    except Exception:
        try:
            socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect(("1.1.1.1", 53))
            return True
        except Exception:
            return False

def check_python_version(py_ver=None):
    if py_ver is None:
        py_ver = sys.version_info
    print(f"[*] Đang kiểm tra Python: {py_ver.major}.{py_ver.minor}")
    if py_ver.major != 3 or py_ver.minor < 10 or py_ver.minor > 13:
        msg = f"Phiên bản Python hiện tại ({py_ver.major}.{py_ver.minor}) không tương thích!"
        details = (
            "Hệ thống PCGD 2026 yêu cầu Python phiên bản từ 3.10 đến 3.13.\n"
            "Hướng dẫn xử lý:\n"
            "1. Tải Python 3.13 chính thức từ: https://www.python.org/downloads/\n"
            "2. Khi cài đặt, BẮT BUỘC tích chọn 'Add python.exe to PATH'.\n"
            "3. Sau khi cài xong, mở lại cửa sổ này và chạy lại CAI_DAT_BAN_DAU.bat."
        )
        log_error(msg, details)
        return False
    print("    -> Phiên bản Python đạt yêu cầu.\n")
    return True

def setup():
    print("=========================================================================================")
    print("  THIẾT LẬP BAN ĐẦU HỆ THỐNG PCGD 2026 - SỞ GIÁO DỤC VÀ ĐÀO TẠO THÁI NGUYÊN")
    print("=========================================================================================\n")
    
    # 1. Kiểm tra Python
    if not check_python_version():
        sys.exit(1)
    
    # 2. Tạo môi trường ảo .venv
    venv_dir = os.path.join(BASE_DIR, ".venv")
    venv_python = os.path.join(venv_dir, "Scripts", "python.exe")
    
    print(f"[*] Đang kiểm tra/tạo môi trường ảo độc lập tại: {venv_dir}")
    if not os.path.exists(venv_python):
        print("    -> Đang khởi tạo môi trường ảo mới (.venv)...")
        res = subprocess.run([sys.executable, "-m", "venv", venv_dir], capture_output=True, text=True)
        if res.returncode != 0:
            log_error("Không thể tạo môi trường ảo .venv!", res.stderr)
            sys.exit(2)
        print("    -> Khởi tạo .venv thành công.")
    else:
        print("    -> Môi trường ảo .venv đã tồn tại.")
        
    # 3. Kiểm tra kết nối mạng trước khi tải thư viện
    print("\n[*] Đang kiểm tra kết nối Internet để chuẩn bị tải thư viện...")
    if not check_internet():
        msg = "Không có kết nối Internet hoặc mạng bị chặn kết nối PyPI!"
        details = (
            "Máy tính cần có kết nối mạng để tải các gói thư viện lần đầu.\n"
            "Hướng dẫn xử lý:\n"
            "1. Kiểm tra dây mạng hoặc kết nối Wi-Fi.\n"
            "2. Nếu cơ quan dùng Proxy/Firewall chặn, hãy cắm mạng không chặn hoặc dùng USB 4G tạm thời.\n"
            "3. Chạy lại CAI_DAT_BAN_DAU.bat sau khi đã có mạng."
        )
        log_error(msg, details)
        sys.exit(3)
    print("    -> Kết nối Internet hoạt động tốt.\n")
    
    # 4. Cài đặt các thư viện từ requirements-lock.txt hoặc requirements.txt
    lock_file = os.path.join(BASE_DIR, "requirements-lock.txt")
    req_file = os.path.join(BASE_DIR, "requirements.txt")
    
    target_req = lock_file if os.path.exists(lock_file) else req_file
    print(f"[*] Đang cài đặt các thư viện đã được kiểm chứng từ: {os.path.basename(target_req)}")
    print("    (Quá trình này mất khoảng 30 - 90 giây tùy tốc độ mạng, vui lòng chờ)...")
    
    pip_cmd = [venv_python, "-m", "pip", "install", "-r", target_req]
    res_pip = subprocess.run(pip_cmd, capture_output=True, text=True)
    
    if res_pip.returncode != 0:
        # Thử lại với requirements.txt nếu lockfile gặp trục trặc
        if target_req != req_file and os.path.exists(req_file):
            print("    [!] Cài từ lockfile gặp cảnh báo, đang thử lại với requirements.txt...")
            pip_cmd_fallback = [venv_python, "-m", "pip", "install", "-r", req_file]
            res_pip = subprocess.run(pip_cmd_fallback, capture_output=True, text=True)
            
        if res_pip.returncode != 0:
            msg = "Cài đặt thư viện qua pip thất bại!"
            details = f"Lệnh chạy: {' '.join(pip_cmd)}\nSTDOUT:\n{res_pip.stdout}\nSTDERR:\n{res_pip.stderr}"
            log_error(msg, details)
            sys.exit(4)
            
    print("    -> Cài đặt thư viện thành công 100%.\n")
    
    # 5. Kiểm thử import độc lập
    print("[*] Đang kiểm tra tính tương thích của các thư viện cốt lõi...")
    test_import_cmd = [venv_python, "-c", "import streamlit, openpyxl, pandas; print('MODULES_OK')"]
    res_import = subprocess.run(test_import_cmd, capture_output=True, text=True)
    if "MODULES_OK" not in res_import.stdout:
        msg = "Thư viện cài đặt xong nhưng không thể nạp (import) vào Python!"
        log_error(msg, res_import.stderr)
        sys.exit(5)
    print("    -> Các module streamlit, openpyxl, pandas đều hoạt động trơn tru.\n")
    
    # 6. Tạo shortcut Desktop
    print("[*] Đang tạo lối tắt tiện ích trên màn hình Desktop...")
    ps_shortcut = os.path.join(BASE_DIR, "scripts", "create_shortcut.ps1")
    if os.path.exists(ps_shortcut):
        subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps_shortcut], capture_output=True)
        print("    -> Đã tạo biểu tượng 'HE THONG PCGD 2026 - THAI NGUYEN' ngoài Desktop.")
        
    print("\n=========================================================================================")
    print("  [XÁC NHẬN] THIẾT LẬP HOÀN TẤT VÀ KIỂM CHỨNG 100% THÀNH CÔNG!")
    print("  Từ nay, Người đẹp chỉ cần:")
    print("  1. Click đúp vào biểu tượng ngoài Desktop 'HE THONG PCGD 2026 - THAI NGUYEN'")
    print("  2. Hoặc click đúp tệp KHOI_DONG_HE_THONG.bat")
    print("  Hệ thống sẽ mở ngay trong 2 giây mà không cần cài đặt lại!")
    print("=========================================================================================\n")

if __name__ == '__main__':
    setup()

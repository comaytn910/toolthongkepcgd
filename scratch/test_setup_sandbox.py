import os
import sys
import subprocess
import shutil

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def test_sandbox_setup():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    sandbox_dir = os.path.join(base_dir, "scratch", "test_portable_setup")
    
    print("=== BẮT ĐẦU KIỂM THỬ THIẾT LẬP TRÊN SANDBOX MÁY MỚI ===")
    print(f"Thư mục sandbox: {sandbox_dir}")
    
    # 1. Dọn dẹp sandbox cũ nếu có
    if os.path.exists(sandbox_dir):
        shutil.rmtree(sandbox_dir)
    os.makedirs(sandbox_dir, exist_ok=True)
    
    # 2. Tạo .venv trong sandbox
    sandbox_venv = os.path.join(sandbox_dir, ".venv")
    print(f"[*] [Bước 1] Đang tạo môi trường ảo thử nghiệm: {sandbox_venv}...")
    res = subprocess.run([sys.executable, "-m", "venv", sandbox_venv], capture_output=True, text=True)
    assert res.returncode == 0, f"Lỗi tạo venv: {res.stderr}"
    assert os.path.exists(os.path.join(sandbox_venv, "Scripts", "python.exe")), "Không tìm thấy python.exe trong sandbox venv"
    print("    -> [ĐẠT] Tạo .venv thành công!")
    
    # 3. Thử cài đặt 1 thư viện nhẹ đại diện (ví dụ: et_xmlfile)
    sandbox_py = os.path.join(sandbox_venv, "Scripts", "python.exe")
    sandbox_pip = os.path.join(sandbox_venv, "Scripts", "pip.exe")
    
    print("[*] [Bước 2] Thử nghiệm pip install trong sandbox...")
    res_pip = subprocess.run([sandbox_pip, "install", "et_xmlfile==2.0.0"], capture_output=True, text=True)
    assert res_pip.returncode == 0, f"Lỗi pip install: {res_pip.stderr}"
    print("    -> [ĐẠT] pip install hoạt động hoàn hảo trong sandbox!")
    
    # 4. Thử nghiệm import
    res_imp = subprocess.run([sandbox_py, "-c", "import et_xmlfile; print('SANDBOX_OK')"], capture_output=True, text=True)
    assert "SANDBOX_OK" in res_imp.stdout, f"Lỗi import: {res_imp.stderr}"
    print("    -> [ĐẠT] Import module thành công trong sandbox!")
    
    # 5. Dọn dẹp sandbox
    print("[*] [Bước 3] Dọn dẹp thư mục sandbox để giải phóng dung lượng...")
    shutil.rmtree(sandbox_dir)
    print("    -> [ĐẠT] Đã dọn dẹp sạch sẽ thư mục sandbox.")
    
    print("\n=== KẾT QUẢ: KIỂM THỬ THIẾT LẬP MÁY MỚI TRÊN SANDBOX ĐẠT 100% ===")

if __name__ == '__main__':
    test_sandbox_setup()

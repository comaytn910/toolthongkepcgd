import os
import sys
import subprocess

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def create_desktop_shortcut():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    target_bat = os.path.join(base_dir, "KHOI_DONG_HE_THONG.bat")
    desktop = os.path.join(os.environ["USERPROFILE"], "Desktop")
    shortcut_path = os.path.join(desktop, "HỆ THỐNG PCGD 2026 - THÁI NGUYÊN.lnk")
    
    # PowerShell script to create Windows shortcut using WScript.Shell
    # Escaping double quotes and backslashes for PowerShell
    target_clean = target_bat.replace("\\", "\\\\")
    base_clean = base_dir.replace("\\", "\\\\")
    shortcut_clean = shortcut_path.replace("\\", "\\\\")
    
    ps_script = f'''
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("{shortcut_clean}")
$Shortcut.TargetPath = "{target_clean}"
$Shortcut.WorkingDirectory = "{base_clean}"
$Shortcut.Description = "Hệ thống Tổng hợp số liệu Phổ cập Giáo dục & Xóa mù chữ năm 2026 - Sở GDĐT Thái Nguyên"
$Shortcut.IconLocation = "$env:SystemRoot\\System32\\imageres.dll,114"
$Shortcut.Save()
Write-Host "SUCCESS"
'''
    try:
        res = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True,
            text=True,
            check=True
        )
        if "SUCCESS" in res.stdout:
            print(f"[OK] Đã tạo thành công lối tắt (Shortcut) trên Desktop:")
            print(f"     -> {shortcut_path}")
            return True
        else:
            print(f"[CẢNH BÁO] Không nhận được phản hồi SUCCESS: {res.stdout} - {res.stderr}")
            return False
    except Exception as e:
        print(f"[LỖI] Không thể tạo shortcut: {e}")
        return False

if __name__ == '__main__':
    create_desktop_shortcut()

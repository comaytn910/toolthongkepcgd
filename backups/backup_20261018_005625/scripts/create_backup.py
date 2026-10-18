import os
import sys
import shutil
import datetime
import zipfile

sys.stdout.reconfigure(encoding='utf-8')

def create_project_backup():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    timestamp_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_root = os.path.join(base_dir, "backups")
    os.makedirs(backup_root, exist_ok=True)
    
    backup_name = f"backup_{timestamp_str}"
    target_backup_dir = os.path.join(backup_root, backup_name)
    os.makedirs(target_backup_dir, exist_ok=True)
    
    # Danh sách các thư mục cần sao lưu
    dirs_to_copy = ['core', 'processors', 'exporters', 'ui', 'config', 'tests', 'outputs', 'scripts']
    
    # Danh sách các tệp riêng lẻ ở thư mục gốc
    files_to_copy = [
        'AGENTS.md',
        'PROJECT_REQUIREMENTS.md',
        'PROGRESS.md',
        'HANDOFF.md',
        'danhmuc_92_xa.json',
        'all_sheets_structure.json',
        'sheets_summary.json'
    ]
    
    copied_files_count = 0
    
    # 1. Sao chép các tệp riêng lẻ
    for fname in files_to_copy:
        src = os.path.join(base_dir, fname)
        if os.path.exists(src):
            dst = os.path.join(target_backup_dir, fname)
            shutil.copy2(src, dst)
            copied_files_count += 1
            
    # 2. Sao chép các thư mục (loại trừ __pycache__ và cache)
    def ignore_patterns(path, names):
        ignored = []
        for n in names:
            if n in ['__pycache__', '.pytest_cache', '.git']:
                ignored.append(n)
            elif n.endswith('.pyc') or n.endswith('.pyo'):
                ignored.append(n)
        return ignored

    for dname in dirs_to_copy:
        src_d = os.path.join(base_dir, dname)
        if os.path.exists(src_d):
            dst_d = os.path.join(target_backup_dir, dname)
            shutil.copytree(src_d, dst_d, ignore=ignore_patterns, dirs_exist_ok=True)
            for root, _, files in os.walk(dst_d):
                copied_files_count += len(files)
                
    # 3. Tạo thêm file nén zip dự phòng
    zip_filename = os.path.join(backup_root, f"{backup_name}.zip")
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(target_backup_dir):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, target_backup_dir)
                zf.write(full_path, rel_path)
                
    print(f"Đã tạo bản sao lưu thành công tại thư mục: {target_backup_dir}")
    print(f"Đã tạo file zip sao lưu: {zip_filename} ({os.path.getsize(zip_filename):,} bytes)")
    print(f"Tổng số tệp được sao lưu: {copied_files_count} tệp.")
    return target_backup_dir, zip_filename

if __name__ == '__main__':
    create_project_backup()

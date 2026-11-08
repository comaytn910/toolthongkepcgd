"""
Script tạo bản sao lưu dự phòng an toàn (Backup)
cho toàn bộ mã nguồn, cấu hình, tài liệu bàn giao và các sản phẩm xuất sắc.
"""

import os
import sys
import zipfile
import time

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BACKUP_DIR = os.path.join(BASE_DIR, 'backups')
os.makedirs(BACKUP_DIR, exist_ok=True)

timestamp = "20261019_012500"
ZIP_NAME = f"backup_{timestamp}.zip"
ZIP_PATH = os.path.join(BACKUP_DIR, ZIP_NAME)

ITEMS_TO_BACKUP = [
    'core',
    'processors',
    'exporters',
    'config',
    'ui',
    'scripts',
    'tests',
    'AGENTS.md',
    'PROJECT_REQUIREMENTS.md',
    'PROGRESS.md',
    'HANDOFF.md',
    'requirements.txt',
    'outputs/SO_CAI_TONG_HOP_PCGD_2026_THAI_NGUYEN.xlsx',
    'outputs/BAO_CAO_DOI_SOAT_63_QUY_TAC_PCGD_2026.xlsx'
]

print("==========================================================================")
print(f"BẮT ĐẦU TẠO BẢN SAO LƯU: {ZIP_NAME}")
print("==========================================================================")
t0 = time.time()

file_count = 0
total_uncompressed_bytes = 0

with zipfile.ZipFile(ZIP_PATH, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
    for item in ITEMS_TO_BACKUP:
        full_path = os.path.join(BASE_DIR, item)
        if not os.path.exists(full_path):
            print(f"   [CẢNH BÁO] Không tìm thấy: {item}")
            continue

        if os.path.isfile(full_path):
            arcname = item
            zf.write(full_path, arcname)
            sz = os.path.getsize(full_path)
            total_uncompressed_bytes += sz
            file_count += 1
            print(f"   + [Tệp] {item:<50} ({sz/1024:.1f} KB)")
        elif os.path.isdir(full_path):
            for root, dirs, files in os.walk(full_path):
                # Bỏ qua __pycache__
                if '__pycache__' in root:
                    continue
                for f in files:
                    if f.endswith('.pyc') or f.endswith('.tmp'):
                        continue
                    file_full = os.path.join(root, f)
                    rel_path = os.path.relpath(file_full, BASE_DIR)
                    zf.write(file_full, rel_path)
                    sz = os.path.getsize(file_full)
                    total_uncompressed_bytes += sz
                    file_count += 1
            print(f"   + [Thư mục] {item}/")

zip_size = os.path.getsize(ZIP_PATH)
dur = time.time() - t0

print("\n--- KIỂM TRA TÍNH TOÀN VẸN BẢN SAO LƯU ---")
with zipfile.ZipFile(ZIP_PATH, 'r') as zf_test:
    bad_file = zf_test.testzip()
    if bad_file:
        print(f"   [LỖI CRC] Tệp hỏng trong zip: {bad_file}")
        sys.exit(1)
    else:
        print(f"   [PASS] 100% {len(zf_test.namelist())} tệp trong zip đều toàn vẹn, kiểm tra CRC thành công!")

print("==========================================================================")
print(f"TẠO BẢN SAO LƯU THÀNH CÔNG TRONG {dur:.2f} GIÂY!")
print(f"Tệp sao lưu: {ZIP_PATH}")
print(f"Tổng số tệp: {file_count} tệp | Kích thước gốc: {total_uncompressed_bytes/1024/1024:.2f} MB | Nén: {zip_size/1024/1024:.2f} MB")
print("==========================================================================\n")

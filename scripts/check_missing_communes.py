import os
import sys
import glob
import json

sys.stdout.reconfigure(encoding='utf-8')
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

with open(os.path.join(BASE_DIR, 'danhmuc_92_xa.json'), 'r', encoding='utf-8') as f:
    dm_92 = json.load(f)

files_in_data = glob.glob(os.path.join(BASE_DIR, 'data', '*.xlsx'))
print(f"Số file thực tế trong data/: {len(files_in_data)}")

# Map stt xã
existing_stt = set()
for fp in files_in_data:
    fn = os.path.basename(fp)
    # Lấy STT đầu tiên
    prefix = fn.split('.')[0].strip()
    try:
        stt = int(prefix.split('_')[0])
        existing_stt.add(stt)
    except:
        pass

missing = []
for item in dm_92:
    if item['stt'] not in existing_stt:
        missing.append(f"{item['stt']:02d}. {item['raw_name']}")

print(f"Số xã có mặt: {len(existing_stt)}")
print(f"Số xã vắng mặt: {len(missing)}")
print("Danh sách các xã không có trong data/:")
for m in missing:
    print(f"  - {m}")

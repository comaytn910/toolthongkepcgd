import os
import sys
import json
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict

# Thêm thư mục gốc vào sys.path
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from config.communes import match_commune_name
from config.sheets_catalog import SHEET_CATALOG

def get_sheets_from_zip(file_path):
    """Đọc danh sách sheet trực tiếp từ workbook.xml qua zipfile siêu tốc (< 10ms)"""
    with zipfile.ZipFile(file_path, 'r') as z:
        with z.open('xl/workbook.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            # Namespace trong workbook.xml
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            sheets_elem = root.find('ns:sheets', ns)
            if sheets_elem is None:
                # Thử không namespace
                sheets = [elem.attrib.get('name') for elem in root.iter() if elem.tag.endswith('sheet')]
            else:
                sheets = [elem.attrib.get('name') for elem in sheets_elem.findall('ns:sheet', ns)]
            return [s for s in sheets if s]

def run_inventory():
    data_dir = os.path.join(BASE_DIR, 'data')
    danhmuc_path = os.path.join(BASE_DIR, 'danhmuc_92_xa.json')
    
    with open(danhmuc_path, 'r', encoding='utf-8') as f:
        danhmuc_92 = json.load(f)
        
    canonical_map = {item['stt']: item['raw_name'] for item in danhmuc_92}
    
    # 1. Tìm toàn bộ file Excel trong data/
    all_files = []
    for root, dirs, files in os.walk(data_dir):
        for f in files:
            if f.startswith('~$') or f.startswith('.'):
                continue
            if not (f.endswith('.xlsx') or f.endswith('.xls')):
                continue
            full_path = os.path.join(root, f)
            rel_path = os.path.relpath(full_path, data_dir)
            all_files.append((full_path, rel_path, f))
            
    print(f"=== TỔNG SỐ FILE EXCEL TÌM THẤY TRONG DATA: {len(all_files)} ===")
    
    results = []
    commune_to_files = defaultdict(list)
    sheet_coverage = defaultdict(int)
    missing_sheets_by_file = {}
    corrupted_files = []
    
    for full_path, rel_path, filename in sorted(all_files, key=lambda x: x[2]):
        file_info = {
            'filename': filename,
            'rel_path': rel_path,
            'size': os.path.getsize(full_path),
            'stt': None,
            'commune_name': None,
            'sheets_found': [],
            'biz_sheets_found': [],
            'missing_biz_sheets': [],
            'status': 'OK',
            'error': None
        }
        
        try:
            sheets = get_sheets_from_zip(full_path)
            file_info['sheets_found'] = sheets
            biz_found = [s for s in sheets if s in SHEET_CATALOG]
            missing_biz = [s for s in SHEET_CATALOG if s not in sheets]
            file_info['biz_sheets_found'] = biz_found
            file_info['missing_biz_sheets'] = missing_biz
            
            for s in biz_found:
                sheet_coverage[s] += 1
                
            if missing_biz:
                missing_sheets_by_file[filename] = missing_biz
                
            # Nhận diện xã từ tên file
            # Format: "01_01. Xã Yên Trạch - Nhập hoàn chỉnh PCGD 2026.xlsx"
            raw_commune = os.path.splitext(filename)[0]
            # Loại bỏ prefix "XX_" nếu có
            if raw_commune[:3].isdigit() and raw_commune[2] == '_':
                raw_commune = raw_commune[3:]
                
            stt, name, score = match_commune_name(raw_commune)
            file_info['stt'] = stt
            file_info['commune_name'] = name if name else raw_commune
            
            if name:
                commune_to_files[stt].append(file_info)
            else:
                commune_to_files['UNKNOWN'].append(file_info)
                
        except Exception as e:
            file_info['status'] = 'CORRUPTED'
            file_info['error'] = str(e)
            corrupted_files.append((filename, str(e)))
            
        results.append(file_info)
        
    matched_stts = set(commune_to_files.keys()) - {'UNKNOWN'}
    missing_stts = [stt for stt in range(1, 93) if stt not in matched_stts]
    duplicate_stts = [stt for stt, flist in commune_to_files.items() if stt != 'UNKNOWN' and len(flist) > 1]
    
    summary = {
        'total_files': len(all_files),
        'unique_communes_matched': len(matched_stts),
        'missing_communes_count': len(missing_stts),
        'missing_communes': [{'stt': stt, 'name': canonical_map.get(stt, '')} for stt in missing_stts],
        'duplicate_communes_count': len(duplicate_stts),
        'duplicate_communes': [
            {
                'stt': stt,
                'name': canonical_map.get(stt, 'Unknown'),
                'files': [f['filename'] for f in commune_to_files[stt]]
            } for stt in duplicate_stts
        ],
        'corrupted_files': corrupted_files,
        'missing_sheets_summary': {k: v for k, v in missing_sheets_by_file.items()}
    }
    
    print("\n" + "="*60)
    print("BÁO CÁO KIỂM KÊ NGUỒN 92 XÃ/PHƯỜNG")
    print("="*60)
    print(f"Tổng số file Excel hợp lệ quét được: {summary['total_files']}")
    print(f"Số đơn vị xã/phường khớp danh mục chuẩn: {summary['unique_communes_matched']} / 92")
    print(f"Số đơn vị còn thiếu: {summary['missing_communes_count']}")
    print(f"Số xã bị trùng lặp file: {summary['duplicate_communes_count']}")
    print(f"Số file bị lỗi/hỏng (Corrupted): {len(corrupted_files)}")
    
    if duplicate_stts:
        print("\nCÁC XÃ BỊ TRÙNG LẶP FILE:")
        for dup in summary['duplicate_communes']:
            print(f"  - [{dup['stt']:02d}] {dup['name']}: {', '.join(dup['files'])}")
            
    if missing_stts:
        print("\nCÁC XÃ CÒN THIẾU FILE:")
        for mis in summary['missing_communes']:
            print(f"  - [{mis['stt']:02d}] {mis['name']}")
            
    print("\nKIỂM TRA ĐỘ PHỦ 33 BIỂU NGHIỆP VỤ:")
    all_33_ok = True
    for code, info in SHEET_CATALOG.items():
        cnt = sheet_coverage[code]
        if cnt < len(all_files):
            all_33_ok = False
            print(f"  [THIẾU/CẢNH BÁO] {code:6s} ({info['name'][:30]:30s}): Có ở {cnt}/{len(all_files)} file")
        else:
            print(f"  [ĐỦ 92/92]       {code:6s} ({info['name'][:30]:30s}): Có ở {cnt}/{len(all_files)} file")
            
    # Ghi kết quả kiểm kê ra file json
    out_inv_json = os.path.join(BASE_DIR, 'inventory_all_92.json')
    with open(out_inv_json, 'w', encoding='utf-8') as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f"\nĐã ghi báo cáo kiểm kê chi tiết vào: {out_inv_json}")
    
    return summary, results

if __name__ == '__main__':
    run_inventory()

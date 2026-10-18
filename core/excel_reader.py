import os
import openpyxl
from config.communes import match_commune_name, CANONICAL_COMMUNES
from config.sheets_catalog import SHEET_CATALOG, SUPPORTING_SHEETS

def scan_data_directory(data_dir: str):
    """
    Quét trực tiếp thư mục data để kiểm tra các file Excel hiện có.
    Không sửa đổi file nguồn, không nhân bản, không giả mạo dữ liệu.
    Hỗ trợ từ 1-2 file kiểm thử cho đến toàn bộ 92 file sau này.
    """
    if not os.path.exists(data_dir):
        return {
            'error': f'Thư mục {data_dir} chưa tồn tại.',
            'files': [],
            'valid_count': 0,
            'duplicate_communes': [],
            'summary': {}
        }
        
    excel_files = [
        f for f in os.listdir(data_dir)
        if (f.endswith('.xlsx') or f.endswith('.xls')) and not f.startswith('~$')
    ]
    
    if not excel_files:
        return {
            'error': 'Thư mục data hiện đang trống. Vui lòng đặt 1–2 file Excel của xã/phường vào thư mục data để kiểm thử.',
            'files': [],
            'valid_count': 0,
            'duplicate_communes': [],
            'summary': {}
        }
        
    inspected_files = []
    seen_communes = {}
    duplicate_communes = []
    
    for filename in sorted(excel_files):
        file_path = os.path.join(data_dir, filename)
        file_info = inspect_single_file(file_path, filename)
        
        # Kiểm tra trùng lặp xã
        commune_name = file_info['commune_name']
        if commune_name:
            if commune_name in seen_communes:
                prev_file = seen_communes[commune_name]
                file_info['errors'].append(f"Trùng xã: Xã '{commune_name}' đã được khai báo ở file '{prev_file}'!")
                file_info['is_valid'] = False
                if commune_name not in duplicate_communes:
                    duplicate_communes.append(commune_name)
            else:
                seen_communes[commune_name] = filename
                
        inspected_files.append(file_info)
        
    valid_count = sum(1 for f in inspected_files if f['is_valid'])
    
    return {
        'files': inspected_files,
        'valid_count': valid_count,
        'total_files': len(inspected_files),
        'duplicate_communes': duplicate_communes,
        'summary': {
            'total_files': len(inspected_files),
            'valid_files': valid_count,
            'has_duplicates': len(duplicate_communes) > 0,
            'communes_found': list(seen_communes.keys())
        }
    }

import zipfile
import xml.etree.ElementTree as ET

def get_sheets_from_zip(file_path: str):
    """Đọc danh sách tên sheet trực tiếp từ workbook.xml siêu tốc (< 10ms)"""
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            with z.open('xl/workbook.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                sheets_elem = root.find('ns:sheets', ns)
                if sheets_elem is None:
                    sheets = [elem.attrib.get('name') for elem in root.iter() if elem.tag.endswith('sheet')]
                else:
                    sheets = [elem.attrib.get('name') for elem in sheets_elem.findall('ns:sheet', ns)]
                return [s for s in sheets if s]
    except Exception:
        return []

def inspect_single_file(file_path: str, filename: str):
    """
    Kiểm tra khả năng đọc của một file Excel cụ thể:
    - Nhận diện xã/phường (qua tên file hoặc cell A2)
    - Kiểm kê các sheet tìm thấy so với 33 biểu nghiệp vụ
    - Kiểm tra biểu 0.M1 có tồn tại và đọc được không
    - Liệt kê các cảnh báo hoặc lỗi nếu có
    """
    errors = []
    warnings = []
    sheets_found = []
    commune_stt = None
    commune_name = None
    has_0_m1 = False
    
    try:
        # Đọc danh sách sheet siêu tốc qua zipfile
        sheets_found = get_sheets_from_zip(file_path)
        if not sheets_found:
            # Fallback sang openpyxl nếu zipfile không đọc được
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=False)
            sheets_found = wb.sheetnames
            wb.close()
        
        # 1. Nhận diện xã/phường từ tên file
        raw_commune = os.path.splitext(filename)[0]
        # Loại bỏ tiền tố '01_', '02_' nếu có
        if len(raw_commune) >= 3 and raw_commune[:2].isdigit() and raw_commune[2] == '_':
            raw_commune = raw_commune[3:]
            
        c_stt, c_name, score = match_commune_name(raw_commune)
        
        # Nếu chưa khớp hoặc không chắc chắn, thử đọc cell A2 trong TỔNG_QUAN hoặc ĐỐI_SOÁT
        if not c_name and any(s in sheets_found for s in ['TỔNG_QUAN', 'ĐỐI_SOÁT']):
            try:
                wb_c = openpyxl.load_workbook(file_path, read_only=True, data_only=False)
                for s_check in ['TỔNG_QUAN', 'ĐỐI_SOÁT']:
                    if s_check in sheets_found:
                        ws_c = wb_c[s_check]
                        val_a2 = ws_c['A2'].value
                        if val_a2 and isinstance(val_a2, str):
                            clean_val = val_a2.replace('Đơn vị/File:', '').replace('Đơn vị:', '').strip()
                            if 'MẪU CHUẨN' not in clean_val:
                                c_stt_c, c_name_c, score_c = match_commune_name(clean_val)
                                if c_name_c:
                                    c_stt, c_name = c_stt_c, c_name_c
                                    break
                wb_c.close()
            except Exception:
                pass
                
        if c_name:
            commune_stt = c_stt
            commune_name = c_name
        else:
            commune_name = f"Chưa xác định ({raw_commune[:30]})"
            warnings.append(f"Chưa khớp chính xác với danh mục 92 xã chuẩn (Nhận diện tạm: {raw_commune[:30]}).")
            
        # 2. Kiểm tra biểu 0.M1
        if '0.M1' in sheets_found:
            has_0_m1 = True
        else:
            errors.append("File thiếu sheet biểu '0.M1'!")
            
        # 3. Thống kê biểu nghiệp vụ
        biz_sheets = [s for s in sheets_found if s in SHEET_CATALOG]
        missing_biz = [s for s in SHEET_CATALOG if s not in sheets_found]
        
        if len(biz_sheets) < 33:
            warnings.append(f"Tìm thấy {len(biz_sheets)}/33 biểu nghiệp vụ (Thiếu {len(missing_biz)} biểu).")
        
    except Exception as e:
        errors.append(f"Lỗi khi đọc file Excel: {str(e)}")
        
    is_valid = len(errors) == 0 and has_0_m1
    
    return {
        'filename': filename,
        'filepath': file_path,
        'commune_stt': commune_stt,
        'commune_name': commune_name,
        'has_0_m1': has_0_m1,
        'sheets_count': len(sheets_found),
        'biz_sheets_count': len([s for s in sheets_found if s in SHEET_CATALOG]),
        'is_valid': is_valid,
        'errors': errors,
        'warnings': warnings
    }

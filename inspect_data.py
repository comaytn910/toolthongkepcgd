import sys
import os

# Đảm bảo đường dẫn import
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from core.excel_reader import scan_data_directory

def main():
    sys.stdout.reconfigure(encoding='utf-8')
    data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
    print(f"=== KIỂM TRA THƯ MỤC DỮ LIỆU: {data_dir} ===\n")
    
    result = scan_data_directory(data_dir)
    
    if 'error' in result and not result['files']:
        print(f"[THÔNG BÁO] {result['error']}")
        return
        
    print(f"Tìm thấy tổng số: {result['total_files']} file. Hợp lệ: {result['valid_count']} file.\n")
    
    for idx, f in enumerate(result['files'], start=1):
        status_str = "[HỢP LỆ]" if f['is_valid'] else "[CÓ LỖI]"
        print(f"{idx}. File: {f['filename']}")
        print(f"   - Xã/phường nhận diện: {f['commune_name']} (STT: {f['commune_stt']})")
        print(f"   - Số biểu tìm thấy: {f['biz_sheets_count']}/33 biểu nghiệp vụ (Tổng sheet: {f['sheets_count']})")
        print(f"   - Biểu 0.M1: {'CÓ' if f['has_0_m1'] else 'THIẾU'}")
        print(f"   - Trạng thái: {status_str}")
        if f['errors']:
            print(f"   - LỖI: {'; '.join(f['errors'])}")
        if f['warnings']:
            print(f"   - Cảnh báo: {'; '.join(f['warnings'])}")
        print()

if __name__ == '__main__':
    main()

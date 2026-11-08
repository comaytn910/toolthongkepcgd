import os
import sys
import subprocess
import openpyxl
import io
import glob

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def find_m1_diffs():
    data_files = sorted(glob.glob("data/*.xlsx"))
    print(f"Tổng số file cần rà soát: {len(data_files)}")
    
    diff_communes = []
    
    for fpath in data_files:
        rel_path = fpath.replace("\\", "/")
        # Lấy file từ git HEAD
        cmd = ["git", "show", f"HEAD:{rel_path}"]
        proc = subprocess.run(cmd, capture_output=True)
        if proc.returncode != 0:
            print(f"File mới chưa có trong git: {rel_path}")
            continue
            
        head_bytes = io.BytesIO(proc.stdout)
        try:
            wb_head = openpyxl.load_workbook(head_bytes, data_only=True)
            wb_curr = openpyxl.load_workbook(fpath, data_only=True)
            
            if '0.M1' not in wb_head.sheetnames or '0.M1' not in wb_curr.sheetnames:
                continue
                
            ws_h = wb_head['0.M1']
            ws_c = wb_curr['0.M1']
            
            # So sánh I5 (trẻ 5t) và K5 (tổng trẻ MN)
            h_i5 = ws_h['I5'].value or 0
            c_i5 = ws_c['I5'].value or 0
            h_k5 = ws_h['K5'].value or 0
            c_k5 = ws_c['K5'].value or 0
            
            if h_i5 != c_i5 or h_k5 != c_k5:
                diff_communes.append({
                    'file': os.path.basename(fpath),
                    'head_i5': h_i5, 'curr_i5': c_i5, 'diff_i5': c_i5 - h_i5,
                    'head_k5': h_k5, 'curr_k5': c_k5, 'diff_k5': c_k5 - h_k5
                })
        except Exception as e:
            print(f"Lỗi đọc file {rel_path}: {e}")
            
    print(f"\n=== KẾT QUẢ RÀ SOÁT: CÓ {len(diff_communes)} XÃ BỊ THAY ĐỔI SỐ LIỆU 0.M1 GIỮA HEAD VÀ HIỆN TẠI ===")
    for item in diff_communes:
        print(f"  - {item['file']}:")
        print(f"      + I5 (Trẻ 5t): HEAD = {item['head_i5']} -> HIỆN TẠI = {item['curr_i5']} (Lệch: {item['diff_i5']:+g})")
        print(f"      + K5 (Tổng MN): HEAD = {item['head_k5']} -> HIỆN TẠI = {item['curr_k5']} (Lệch: {item['diff_k5']:+g})")

if __name__ == '__main__':
    find_m1_diffs()

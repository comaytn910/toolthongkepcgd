import os
import sys
import time
import openpyxl

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def benchmark_fast_read():
    data_dir = os.path.join(BASE_DIR, 'data')
    files = [os.path.join(data_dir, f) for f in sorted(os.listdir(data_dir)) if f.endswith('.xlsx') and not f.startswith('~$')]
    print(f"Bắt đầu test tốc độ đọc sheet 0.M5 trên {len(files)} file...")
    t0 = time.time()
    
    total_e8 = 0.0
    for idx, f in enumerate(files, 1):
        wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        ws = wb['0.M5']
        # Đọc hàng 8
        for row in ws.iter_rows(min_row=8, max_row=8, min_col=5, max_col=5, values_only=True):
            val = row[0]
            if isinstance(val, (int, float)):
                total_e8 += val
        wb.close()
        
    dur = time.time() - t0
    print(f"Đã đọc xong {len(files)} file trong {dur:.2f} giây! (Trung bình {dur/len(files)*1000:.1f}ms/file)")
    print(f"Tổng E8 (Chi NS thường xuyên): {total_e8:,.2f}")

if __name__ == '__main__':
    benchmark_fast_read()

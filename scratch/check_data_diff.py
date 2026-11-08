import os
import sys
import subprocess
import openpyxl
import io

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def check_diff():
    fname = "data/22_22. Xã Thần Sa - Nhập hoàn chỉnh PCGD 2026.xlsx"
    cmd = ["git", "show", f"HEAD:{fname}"]
    proc = subprocess.run(cmd, capture_output=True)
    if proc.returncode != 0:
        print("Lỗi git show")
        return
    
    head_bytes = io.BytesIO(proc.stdout)
    wb_head = openpyxl.load_workbook(head_bytes, data_only=True)
    wb_curr = openpyxl.load_workbook(fname, data_only=True)
    
    ws_head = wb_head['0.M1']
    ws_curr = wb_curr['0.M1']
    
    diffs = []
    for r in range(1, 41):
        for c in range(1, 12):
            cell_h = ws_head.cell(r, c).value
            cell_c = ws_curr.cell(r, c).value
            if cell_h != cell_c:
                coord = ws_head.cell(r, c).coordinate
                diffs.append((coord, cell_h, cell_c))
                
    print(f"File 22. Thần Sa - Khác biệt trong sheet 0.M1: {len(diffs)} ô")
    for d in diffs:
        print(f"  {d[0]}: HEAD = {d[1]} | CURR = {d[2]}")

if __name__ == '__main__':
    check_diff()

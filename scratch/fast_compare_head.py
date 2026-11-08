import os
import sys
import subprocess
import glob
import io
import zipfile
import xml.etree.ElementTree as ET

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from core.fast_excel import fast_read_sheet, parse_shared_strings, get_sheet_target_path

def fast_read_from_zip(z, sheet_name):
    target = get_sheet_target_path(z, sheet_name)
    if not target or target not in z.namelist():
        return {}
    shared_strings = parse_shared_strings(z)
    cells = {}
    with z.open(target) as f:
        for event, elem in ET.iterparse(f, events=('end',)):
            if elem.tag.endswith('c'):
                coord = elem.attrib.get('r')
                t_type = elem.attrib.get('t')
                v_elem = None
                f_elem = None
                for child in elem:
                    if child.tag.endswith('v'):
                        v_elem = child
                    elif child.tag.endswith('f'):
                        f_elem = child
                
                val = None
                formula = f_elem.text if f_elem is not None else None
                
                if v_elem is not None and v_elem.text is not None:
                    raw_val = v_elem.text
                    if t_type == 's':
                        try:
                            idx = int(raw_val)
                            val = shared_strings[idx] if idx < len(shared_strings) else None
                        except ValueError:
                            val = raw_val
                    elif t_type == 'b':
                        val = bool(int(raw_val))
                    else:
                        try:
                            val = float(raw_val) if '.' in raw_val else int(raw_val)
                        except ValueError:
                            val = raw_val
                cells[coord] = {'val': val, 'formula': formula}
                elem.clear()
    return cells

def compare_all_m1():
    files = sorted(glob.glob("data/*.xlsx"))
    print(f"Bắt đầu so sánh 0.M1 trên {len(files)} file giữa git HEAD và thư mục hiện tại...")
    
    diffs = []
    
    for fpath in files:
        rel_path = fpath.replace("\\", "/")
        # Lấy file từ git
        cmd = ["git", "show", f"HEAD:{rel_path}"]
        proc = subprocess.run(cmd, capture_output=True)
        if proc.returncode != 0:
            continue
            
        z_head = zipfile.ZipFile(io.BytesIO(proc.stdout))
        z_curr = zipfile.ZipFile(fpath)
        
        cells_head = fast_read_from_zip(z_head, '0.M1')
        cells_curr = fast_read_from_zip(z_curr, '0.M1')
        
        # So sánh các ô hàng 5-29, cột D..K
        changed_cells = []
        for r in range(5, 30):
            for col in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
                coord = f"{col}{r}"
                v_h = cells_head.get(coord, {}).get('val')
                v_c = cells_curr.get(coord, {}).get('val')
                if v_h != v_c:
                    changed_cells.append((coord, v_h, v_c))
                    
        if changed_cells:
            diffs.append({
                'file': os.path.basename(fpath),
                'changes': changed_cells
            })
            
    print(f"\n=== KẾT QUẢ: CÓ {len(diffs)} FILE BỊ THAY ĐỔI DỮ LIỆU TRONG SHEET 0.M1 ===")
    for d in diffs:
        print(f"\n[+] {d['file']}:")
        for c in d['changes']:
            print(f"    - {c[0]}: HEAD = {c[1]} -> CURR = {c[2]}")

if __name__ == '__main__':
    compare_all_m1()

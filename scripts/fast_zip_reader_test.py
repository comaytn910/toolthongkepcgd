import os
import sys
import time
import zipfile
import xml.etree.ElementTree as ET

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def get_sheet_target(z, sheet_name):
    # Đọc xl/workbook.xml để lấy sheetId và r:id
    with z.open('xl/workbook.xml') as f:
        tree = ET.parse(f)
        root = tree.getroot()
        sheet_elem = None
        for elem in root.iter():
            if elem.tag.endswith('sheet') and elem.attrib.get('name') == sheet_name:
                sheet_elem = elem
                break
        if sheet_elem is None:
            return None
        r_id = None
        for k, v in sheet_elem.attrib.items():
            if k.endswith('id'):
                r_id = v
                break
                
    # Đọc xl/_rels/workbook.xml.rels để ánh xạ r:id sang target file
    with z.open('xl/_rels/workbook.xml.rels') as f:
        tree = ET.parse(f)
        root = tree.getroot()
        for elem in root.iter():
            if elem.attrib.get('Id') == r_id:
                target = elem.attrib.get('Target')
                if not target.startswith('xl/'):
                    target = 'xl/' + target
                return target
    return None

def read_sheet_cells_fast(zip_path, sheet_name):
    cells = {}
    with zipfile.ZipFile(zip_path, 'r') as z:
        target = get_sheet_target(z, sheet_name)
        if not target or target not in z.namelist():
            return cells
        with z.open(target) as f:
            # Dùng iterparse để streaming cực nhanh
            for event, elem in ET.iterparse(f, events=('end',)):
                if elem.tag.endswith('c'):
                    coord = elem.attrib.get('r')
                    # Tìm tag <v> (giá trị đã tính toán)
                    v_elem = None
                    for child in elem:
                        if child.tag.endswith('v'):
                            v_elem = child
                            break
                    if v_elem is not None and v_elem.text is not None:
                        val_str = v_elem.text
                        try:
                            cells[coord] = float(val_str)
                        except ValueError:
                            cells[coord] = val_str
                    elem.clear()
    return cells

def test_fast_reader():
    data_dir = os.path.join(BASE_DIR, 'data')
    files = [os.path.join(data_dir, f) for f in sorted(os.listdir(data_dir)) if f.endswith('.xlsx') and not f.startswith('~$')]
    print(f"Bắt đầu đọc sheet 0.M5 của {len(files)} file bằng Fast Zip Reader...")
    t0 = time.time()
    
    total_e8 = 0.0
    for idx, f in enumerate(files, 1):
        cells = read_sheet_cells_fast(f, '0.M5')
        v = cells.get('E8', 0.0)
        if isinstance(v, (int, float)):
            total_e8 += v
            
    dur = time.time() - t0
    print(f"Hoàn thành {len(files)} file trong CHỈ {dur:.3f} giây! ({dur/len(files)*1000:.2f} ms/file)")
    print(f"Tổng E8 tính được: {total_e8:,.2f} Tr.đg")

if __name__ == '__main__':
    test_fast_reader()

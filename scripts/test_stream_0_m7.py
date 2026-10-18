import os
import sys
import time
import zipfile
import xml.etree.ElementTree as ET

def test_stream_0_m7():
    t0 = time.time()
    files = [f for f in os.listdir('data') if f.endswith('.xlsx') and not f.startswith('~$')]
    files.sort()
    
    ns = {'d': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    all_students = []
    commune_stats = []
    
    for fname in files:
        fpath = os.path.join('data', fname)
        c_name = fname.split(' - ')[0] if ' - ' in fname else fname
        
        with zipfile.ZipFile(fpath, 'r') as z:
            # 1. Tìm target sheet 0.M7
            wb_xml = z.read('xl/workbook.xml')
            root = ET.fromstring(wb_xml)
            r_id = None
            for s in root.findall('.//d:sheet', ns):
                if s.attrib.get('name', '').strip() == '0.M7':
                    r_id = s.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    break
            if not r_id:
                continue
            rels_xml = z.read('xl/_rels/workbook.xml.rels')
            r_root = ET.fromstring(rels_xml)
            target = None
            for rel in r_root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                if rel.attrib.get('Id') == r_id:
                    target = 'xl/' + rel.attrib.get('Target').lstrip('/')
                    break
            if not target:
                continue
                
            # Đọc shared strings nếu có
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                ss_root = ET.fromstring(z.read('xl/sharedStrings.xml'))
                for si in ss_root.findall('d:si', ns):
                    t = si.find('d:t', ns)
                    shared_strings.append(t.text if t is not None and t.text else '')
                    
            # Đọc sheet
            s_xml = z.read(target)
            s_root = ET.fromstring(s_xml)
            
            c_count = 0
            for r in s_root.findall('.//d:row', ns):
                r_idx = int(r.attrib.get('r', 0))
                if r_idx < 10:
                    continue
                row_vals = {}
                for c in r.findall('d:c', ns):
                    ref = c.attrib.get('r', '')
                    col = ''.join([ch for ch in ref if ch.isalpha()])
                    t_attr = c.attrib.get('t')
                    v = c.find('d:v', ns)
                    val = v.text if v is not None else None
                    if val is not None and t_attr == 's':
                        try:
                            val = shared_strings[int(val)]
                        except Exception:
                            pass
                    row_vals[col] = val
                    
                ho_ten = row_vals.get('C')
                if ho_ten and str(ho_ten).strip() and str(ho_ten).strip() not in ['None', '0', '0.0']:
                    c_count += 1
            
            commune_stats.append((c_name, c_count))
            
    dur = time.time() - t0
    total_st = sum(c[1] for c in commune_stats)
    print(f"Read {len(commune_stats)} communes, total {total_st:,} students in {dur:.2f}s")
    for item in commune_stats[:5]:
        print(f"  {item[0]}: {item[1]} students")

if __name__ == '__main__':
    test_stream_0_m7()

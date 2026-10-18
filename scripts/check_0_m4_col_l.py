import os
import zipfile
import xml.etree.ElementTree as ET

def inspect_0_m4_all():
    files = [f for f in os.listdir('data') if f.endswith('.xlsx') and not f.startswith('~$')]
    files.sort()
    
    col_l_vals = []
    has_m4_count = 0
    row_counts = []
    
    for fname in files:
        fpath = os.path.join('data', fname)
        try:
            with zipfile.ZipFile(fpath, 'r') as z:
                # Find sheet 0.M4
                wb_xml = z.read('xl/workbook.xml')
                root = ET.fromstring(wb_xml)
                ns = {'d': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
                sheet_target = None
                for s in root.findall('.//d:sheet', ns):
                    if s.attrib.get('name', '').strip() == '0.M4':
                        r_id = s.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                        # read rels
                        rels_xml = z.read('xl/_rels/workbook.xml.rels')
                        r_root = ET.fromstring(rels_xml)
                        for rel in r_root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                            if rel.attrib.get('Id') == r_id:
                                sheet_target = 'xl/' + rel.attrib.get('Target').lstrip('/')
                                break
                        break
                
                if not sheet_target:
                    continue
                
                has_m4_count += 1
                s_xml = z.read(sheet_target)
                s_root = ET.fromstring(s_xml)
                
                # Check L9, L10, K9, H9, K10, H10
                cells = {}
                for c in s_root.findall('.//d:c', ns):
                    ref = c.attrib.get('r')
                    if ref in ['L9', 'L10', 'K9', 'H9', 'K10', 'H10', 'L13', 'L14', 'C9', 'C11', 'H13', 'K13']:
                        v_el = c.find('d:v', ns)
                        f_el = c.find('d:f', ns)
                        val = v_el.text if v_el is not None else ''
                        formula = f_el.text if f_el is not None else ''
                        cells[ref] = (val, formula)
                
                col_l_vals.append({
                    'file': fname[:15],
                    'L9': cells.get('L9', ('', '')),
                    'K9': cells.get('K9', ('', '')),
                    'H9': cells.get('H9', ('', '')),
                    'L13': cells.get('L13', ('', '')),
                })
        except Exception as e:
            pass
            
    print(f"Total files with 0.M4: {has_m4_count}/{len(files)}")
    print("Sample L9 values across first 15 communes:")
    for item in col_l_vals[:15]:
        print(f"{item['file']}: K9={item['K9'][0]}, H9={item['H9'][0]} -> L9={item['L9'][0]} (formula: {item['L9'][1]}), L13={item['L13'][0]} ({item['L13'][1]})")

if __name__ == '__main__':
    inspect_0_m4_all()

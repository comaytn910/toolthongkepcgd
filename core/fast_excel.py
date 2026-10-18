import os
import zipfile
import xml.etree.ElementTree as ET

def parse_shared_strings(z):
    """Đọc sharedStrings.xml nếu có để giải mã các ô dạng t='s'"""
    if 'xl/sharedStrings.xml' not in z.namelist():
        return []
    strings = []
    with z.open('xl/sharedStrings.xml') as f:
        for event, elem in ET.iterparse(f, events=('end',)):
            if elem.tag.endswith('si'):
                # Gom text trong si (có thể có nhiều t con)
                text_parts = []
                for t in elem.iter():
                    if t.tag.endswith('t') and t.text:
                        text_parts.append(t.text)
                strings.append(''.join(text_parts))
                elem.clear()
    return strings

def get_sheet_target_path(z, sheet_name):
    """Tìm đường dẫn file XML tương ứng với sheet_name"""
    if 'xl/workbook.xml' not in z.namelist() or 'xl/_rels/workbook.xml.rels' not in z.namelist():
        return None
        
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

def fast_read_sheet(file_path: str, sheet_name: str):
    """
    Đọc toàn bộ dữ liệu của một sheet từ file .xlsx qua zipfile và XML iterparse.
    Tốc độ: ~10ms / file.
    Trả về dict: cells[coord] = {
        'val': float hoặc str hoặc None,
        'type': 'NUMBER' | 'EMPTY' | 'NOT_APPLICABLE' | 'TEXT' | 'ERROR_TYPE',
        'formula': str hoặc None,
        'raw': str hoặc None
    }
    """
    cells = {}
    if not os.path.exists(file_path):
        return cells
        
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            target = get_sheet_target_path(z, sheet_name)
            if not target or target not in z.namelist():
                return cells
                
            shared_strings = None
            
            with z.open(target) as f:
                for event, elem in ET.iterparse(f, events=('end',)):
                    if elem.tag.endswith('c'):
                        coord = elem.attrib.get('r')
                        cell_type = elem.attrib.get('t', '')
                        
                        formula_str = None
                        val_str = None
                        
                        for child in elem:
                            tag = child.tag
                            if tag.endswith('f'):
                                formula_str = f"={child.text}" if child.text else None
                            elif tag.endswith('v'):
                                val_str = child.text
                            elif tag.endswith('is'):
                                # inlineStr
                                for t_elem in child.iter():
                                    if t_elem.tag.endswith('t') and t_elem.text:
                                        val_str = t_elem.text
                                        break
                                        
                        # Giải mã giá trị và kiểu
                        parsed_val = None
                        parsed_type = 'EMPTY'
                        
                        if cell_type == 's' and val_str is not None:
                            # Shared string
                            if shared_strings is None:
                                shared_strings = parse_shared_strings(z)
                            try:
                                s_idx = int(val_str)
                                if 0 <= s_idx < len(shared_strings):
                                    parsed_val = shared_strings[s_idx]
                                else:
                                    parsed_val = ""
                            except ValueError:
                                parsed_val = val_str
                                
                            str_clean = str(parsed_val).strip()
                            if str_clean.lower() == 'x':
                                parsed_type = 'NOT_APPLICABLE'
                                parsed_val = 'x'
                            elif str_clean == '':
                                parsed_type = 'EMPTY'
                                parsed_val = None
                            else:
                                # Thử ép kiểu số nếu chuỗi là số
                                try:
                                    clean_num = str_clean.replace(',', '')
                                    parsed_val = float(clean_num)
                                    parsed_type = 'NUMBER'
                                except ValueError:
                                    parsed_type = 'TEXT'
                        elif cell_type == 'e':
                            # Formula error
                            parsed_type = 'ERROR_TYPE'
                            parsed_val = val_str
                        elif cell_type == 'b':
                            # Boolean
                            parsed_type = 'TEXT'
                            parsed_val = True if val_str == '1' else False
                        else:
                            # Number, general, str
                            if val_str is None or val_str.strip() == '':
                                parsed_type = 'EMPTY'
                                parsed_val = None
                            else:
                                str_clean = val_str.strip()
                                if str_clean.lower() == 'x':
                                    parsed_type = 'NOT_APPLICABLE'
                                    parsed_val = 'x'
                                else:
                                    try:
                                        clean_num = str_clean.replace(',', '')
                                        parsed_val = float(clean_num)
                                        parsed_type = 'NUMBER'
                                    except ValueError:
                                        parsed_val = str_clean
                                        parsed_type = 'TEXT'
                                        
                        cells[coord] = {
                            'val': parsed_val,
                            'type': parsed_type,
                            'formula': formula_str,
                            'raw': val_str
                        }
                        elem.clear()
    except Exception as e:
        pass
        
    return cells

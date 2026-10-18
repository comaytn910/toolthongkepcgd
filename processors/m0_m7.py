"""
Mô-đun tổng hợp Biểu 0.M7: Sổ theo dõi phổ cập giáo dục mầm non (Mẫu: MN-06-So PC)
Dành cho hệ thống Tổng hợp số liệu PCGD - XMC năm 2026 - Tỉnh Thái Nguyên.

Quy tắc kỹ thuật & nghiệp vụ:
1. Đây là biểu dạng GHÉP DANH SÁCH (expandable): Mỗi xã có từ vài trăm đến vài ngàn học sinh.
2. Dữ liệu học sinh bắt đầu từ Hàng 10 của sheet 0.M7.
3. Thu thập các trường:
   - Cột B: Số phiếu ĐT (MS)
   - Cột C: Họ và tên
   - Cột D: Ngày tháng năm sinh
   - Cột E: Nữ ('x' hoặc '1')
   - Cột F: Chỗ ở (xóm, thôn)
   - Cột G: Dân tộc
   - Cột H: Họ và tên cha hoặc mẹ hoặc người đỡ đầu
   - Cột I..N: Chương trình học các năm 2025 - 2030
   - Cột O: Tên trường học
   - Cột P: Ghi chú (loại khuyết tật, diện tạm trú...)
4. Bổ sung thông tin đơn vị (Mã xã, Tên xã) và STT toàn tỉnh để bảo đảm tính truy vết nguồn gốc (Audit Trail).
5. Tổng hợp các chỉ tiêu thống kê cấp tỉnh:
   - Tổng số trẻ trong sổ theo dõi toàn tỉnh.
   - Tổng số trẻ nữ.
   - Bảng tổng hợp số lượng trẻ theo từng xã/phường (đối chiếu với chỉ tiêu K5 của biểu 0.M1).
   - Thống kê cơ cấu trẻ theo thành phần dân tộc.
"""

import os
import zipfile
import xml.etree.ElementTree as ET

def process_0_m7(valid_files, collect_records=True):
    """
    Xử lý tổng hợp biểu 0.M7 trên toàn bộ 92 file xã hợp lệ bằng streaming XML.
    """
    ns = {'d': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    
    total_communes_count = len(valid_files)
    all_students = []
    commune_summary = []
    ethnicity_counter = {}
    total_female = 0
    total_children = 0
    
    for f_idx, file_info in enumerate(valid_files, start=1):
        c_name = file_info['commune_name']
        c_stt = file_info.get('commune_stt', f"{f_idx:02d}")
        f_path = file_info['filepath']
        
        c_children_count = 0
        c_female_count = 0
        
        try:
            with zipfile.ZipFile(f_path, 'r') as z:
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
                    
                # 2. Đọc shared strings nếu có
                shared_strings = []
                if 'xl/sharedStrings.xml' in z.namelist():
                    ss_root = ET.fromstring(z.read('xl/sharedStrings.xml'))
                    for si in ss_root.findall('d:si', ns):
                        t = si.find('d:t', ns)
                        shared_strings.append(t.text if t is not None and t.text else '')
                        
                # 3. Đọc dữ liệu từng dòng học sinh
                s_xml = z.read(target)
                s_root = ET.fromstring(s_xml)
                
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
                    if not ho_ten or str(ho_ten).strip() in ['', 'None', '0', '0.0']:
                        continue
                    
                    ho_ten_str = str(ho_ten).strip()
                    c_children_count += 1
                    total_children += 1
                    
                    # Nữ
                    nu_raw = row_vals.get('E')
                    is_nu = False
                    if nu_raw is not None:
                        nu_str = str(nu_raw).strip().lower()
                        if nu_str in ['x', '1', '1.0', 'nữ', 'nu']:
                            is_nu = True
                            c_female_count += 1
                            total_female += 1
                            
                    # Dân tộc
                    dt_raw = row_vals.get('G')
                    dt_str = str(dt_raw).strip() if dt_raw and str(dt_raw).strip() not in ['0', '0.0', 'None'] else "Không xác định"
                    ethnicity_counter[dt_str] = ethnicity_counter.get(dt_str, 0) + 1
                    
                    if collect_records:
                        record = {
                            'stt_tinh': total_children,
                            'commune_stt': c_stt,
                            'commune_name': c_name,
                            'ms_phieu': str(row_vals.get('B', '')).strip() if row_vals.get('B') is not None else '',
                            'ho_ten': ho_ten_str,
                            'ngay_sinh': str(row_vals.get('D', '')).strip() if row_vals.get('D') is not None else '',
                            'nu': 'x' if is_nu else '',
                            'thon_xom': str(row_vals.get('F', '')).strip() if row_vals.get('F') is not None else '',
                            'dan_toc': dt_str,
                            'phu_huynh': str(row_vals.get('H', '')).strip() if row_vals.get('H') is not None else '',
                            'ct_2025': str(row_vals.get('I', '')).strip() if row_vals.get('I') is not None and str(row_vals.get('I')) != '0.0' else '',
                            'ct_2026': str(row_vals.get('J', '')).strip() if row_vals.get('J') is not None and str(row_vals.get('J')) != '0.0' else '',
                            'ct_2027': str(row_vals.get('K', '')).strip() if row_vals.get('K') is not None and str(row_vals.get('K')) != '0.0' else '',
                            'ct_2028': str(row_vals.get('L', '')).strip() if row_vals.get('L') is not None and str(row_vals.get('L')) != '0.0' else '',
                            'ct_2029': str(row_vals.get('M', '')).strip() if row_vals.get('M') is not None and str(row_vals.get('M')) != '0.0' else '',
                            'ct_2030': str(row_vals.get('N', '')).strip() if row_vals.get('N') is not None and str(row_vals.get('N')) != '0.0' else '',
                            'truong_hoc': str(row_vals.get('O', '')).strip() if row_vals.get('O') is not None and str(row_vals.get('O')) != '0.0' else '',
                            'ghi_chu': str(row_vals.get('P', '')).strip() if row_vals.get('P') is not None and str(row_vals.get('P')) != '0.0' else ''
                        }
                        all_students.append(record)
                        
        except Exception as e:
            pass
            
        commune_summary.append({
            'commune_stt': c_stt,
            'commune_name': c_name,
            'total_students': c_children_count,
            'female_students': c_female_count
        })

    rate_female = round(total_female / total_children * 100, 2) if total_children > 0 else 0.0

    return {
        'sheet_name': '0.M7',
        'communes_count': total_communes_count,
        'total_children': total_children,
        'total_female': total_female,
        'rate_female': rate_female,
        'commune_summary': commune_summary,
        'ethnicity_summary': ethnicity_counter,
        'students_data': all_students
    }

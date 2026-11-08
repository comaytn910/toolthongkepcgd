import openpyxl
from core.fast_excel import fast_read_sheet

COUNT_COLS_2M4 = ['C', 'D', 'E', 'F', 'G', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X']
AREA_COLS_2M4 = ['P', 'R', 'T', 'V', 'X']

def clean_num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v) if v < 50000000 else 0.0
    if isinstance(v, str):
        try:
            num = float(v.strip().replace(',', ''))
            return num if num < 50000000 else 0.0
        except ValueError:
            return 0.0
    return 0.0

def process_2_m4(valid_files):
    """
    Tổng hợp Biểu 2.M4: Thống kê cơ sở vật chất phổ cập giáo dục trung học cơ sở (Mẫu THCS-01-CSVC).
    
    Cấu trúc bảng theo mẫu chuẩn:
    - Bảng gồm các hàng 8..11 (chi tiết các trường) và Hàng 12 (Dòng CỘNG).
    - Cột A: TT
    - Cột B: Trường cấp THCS
    - Cột C: Số điểm trường
    - Cột D: Số lớp
    - Cột E..G: Phòng học (Kiên cố, Bán kiên cố, Tạm)
    - Cột H: Tỉ lệ Ph/Lớp % = (E+F+G) / D * 100
    - Cột I..N: Phòng chức năng (Hiệu trưởng, P.Hiệu trưởng, Văn phòng, Y tế, Phòng họp, Thư viện)
    - Cột O..P: Phòng thí nghiệm (SL, DT)
    - Cột Q..R: Vệ sinh GV (SL, DT)
    - Cột S..T: Vệ sinh HS (SL, DT)
    - Cột U..V: Sân chơi (SL, DT)
    - Cột W..X: Bãi tập (SL, DT)
    
    Quy tắc nghiệp vụ:
    1. Định vị động dòng Cộng của từng xã (tìm text 'Cộng'/'Tổng'/'Tổng cộng' ở cột B hoặc A)
       để xử lý trơn tru các xã co dòng (dòng 9..11) hoặc giãn dòng (dòng 14, 17).
    2. Cấp tỉnh tính lại tỷ lệ chuẩn:
       - H12 = sum(E12+F12+G12) / D12 * 100
       - Định dạng tỷ lệ 0.00"%" theo đúng chuẩn quy ước.
    3. Trích xuất chi tiết từng trường THCS của từng xã để tạo Sheet đối soát chi tiết.
    """
    commune_summaries = []
    schools_details = []
    discrepancies = []
    
    province_counts = {c: 0.0 for c in COUNT_COLS_2M4}
    total_communes = len(valid_files)
    
    for f in valid_files:
        fpath = f['filepath']
        cname = f['commune_name']
        cstt = f.get('commune_stt')
        
        cells = fast_read_sheet(fpath, '2.M4')
        if not cells:
            continue
            
        # Tìm dòng Cộng
        tot_row = None
        for r in range(8, 60):
            b_val = str(cells.get(f"B{r}", {}).get("val") or "").strip().lower()
            a_val = str(cells.get(f"A{r}", {}).get("val") or "").strip().lower()
            if b_val in ["cộng", "tổng", "tổng cộng"] or a_val in ["cộng", "tổng", "tổng cộng"]:
                tot_row = r
                break
                
        if tot_row is None:
            tot_row = 12
            
        # Lấy giá trị dòng tổng của xã
        comm_vals = {}
        for c in COUNT_COLS_2M4:
            comm_vals[c] = clean_num(cells.get(f"{c}{tot_row}", {}).get("val"))
            
        # Tính lại tỷ lệ phòng/lớp của xã
        comm_tot_rooms = comm_vals['E'] + comm_vals['F'] + comm_vals['G']
        comm_rate_h = (comm_tot_rooms / comm_vals['D'] * 100) if comm_vals['D'] > 0 else 0.0
        
        # Đọc chi tiết từng trường của xã từ dòng 8 đến tot_row - 1
        for r in range(8, tot_row):
            sch_name = str(cells.get(f"B{r}", {}).get("val") or "").strip()
            sch_classes = clean_num(cells.get(f"D{r}", {}).get("val"))
            if sch_name and not sch_name.lower().startswith("cộng") and sch_classes > 0:
                sch_dict = {
                    'commune_stt': cstt,
                    'commune_name': cname,
                    'school_name': sch_name,
                    'row_idx': r
                }
                for c in COUNT_COLS_2M4:
                    sch_dict[c] = clean_num(cells.get(f"{c}{r}", {}).get("val"))
                sch_rooms = sch_dict['E'] + sch_dict['F'] + sch_dict['G']
                sch_dict['rate_h'] = (sch_rooms / sch_dict['D'] * 100) if sch_dict['D'] > 0 else 0.0
                schools_details.append(sch_dict)
                
        # Cộng dồn cấp tỉnh
        for c in COUNT_COLS_2M4:
            province_counts[c] += comm_vals[c]
            
        commune_summaries.append({
            'stt': cstt,
            'commune_name': cname,
            'filename': f.get('filename', ''),
            'tot_row': tot_row,
            'counts': comm_vals,
            'total_rooms': comm_tot_rooms,
            'rate_h': comm_rate_h
        })

    # Tính toán cấp tỉnh
    province_tot_rooms = province_counts['E'] + province_counts['F'] + province_counts['G']
    province_rate_h = (province_tot_rooms / province_counts['D'] * 100) if province_counts['D'] > 0 else 0.0
    
    # Tạo aggregated_cells để ghi vào Hàng 12 của template
    aggregated_cells = {}
    for c in COUNT_COLS_2M4:
        aggregated_cells[f"{c}12"] = province_counts[c]
    aggregated_cells["H12"] = round(province_rate_h, 2)
    aggregated_cells["total_rooms"] = province_tot_rooms
    
    return {
        'total_communes_count': total_communes,
        'commune_names': [f['commune_name'] for f in valid_files],
        'commune_summaries': commune_summaries,
        'schools_details': schools_details,
        'total_schools': len(schools_details),
        'discrepancies': discrepancies,
        'province_counts': province_counts,
        'province_tot_rooms': province_tot_rooms,
        'province_rate_h': province_rate_h,
        'aggregated_cells': aggregated_cells
    }

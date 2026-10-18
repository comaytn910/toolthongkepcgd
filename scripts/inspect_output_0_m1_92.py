import os
import sys
import openpyxl

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

out_path = os.path.join(BASE_DIR, 'outputs', 'TONG_HOP_THU_0.M1_PCGD_2026.xlsx')
print(f"Kiểm tra file: {out_path} ({os.path.getsize(out_path):,} bytes)")

wb = openpyxl.load_workbook(out_path, data_only=True)
print("Các sheet có trong file kết quả:", wb.sheetnames)
ws = wb['TONG_HOP_THU']
print("C1 (Tiêu đề):", ws['C1'].value)
print("C2 (Ghi chú số xã):", ws['C2'].value)
print("Hàng 5 D5:K5:", [ws[f"{c}5"].value for c in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']])
print("Hàng 16 D16:K16 (Tỷ lệ huy động):", [ws[f"{c}16"].value for c in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']])
print("Hàng 22 D22:K22 (Tỷ lệ chuyên cần):", [ws[f"{c}22"].value for c in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']])

# Kiểm tra sheet chi tiết
if 'CHI_TIET_THEO_XA' in wb.sheetnames:
    ws_dt = wb['CHI_TIET_THEO_XA']
    print("CHI_TIET_THEO_XA max row:", ws_dt.max_row)
    # Lấy danh sách các xã xuất hiện
    communes_in_dt = set()
    for r in range(5, ws_dt.max_row + 1):
        c_name = ws_dt.cell(row=r, column=2).value
        if c_name:
            communes_in_dt.add(c_name)
    print(f"Tổng số xã trong bảng chi tiết: {len(communes_in_dt)}")
    
if 'DOI_CHIEU_K_NGUON' in wb.sheetnames:
    ws_k = wb['DOI_CHIEU_K_NGUON']
    print("DOI_CHIEU_K_NGUON max row:", ws_k.max_row)
    if ws_k.max_row > 4:
        for r in range(5, min(ws_k.max_row + 1, 15)):
            row_v = [ws_k.cell(row=r, column=ci).value for ci in range(1, 9)]
            print(f"  Lệch K [{r-4}]: Xã={row_v[1]}, Hàng={row_v[2]}, Ô={row_v[3]}, Nguồn={row_v[4]}, Tính lại={row_v[5]}, Lệch={row_v[6]}")

wb.close()

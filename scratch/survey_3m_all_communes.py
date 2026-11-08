import sys
import glob
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

files = sorted(glob.glob('data/*.xlsx'))
print(f"Tổng số file xã tìm thấy: {len(files)}")

sheets = ['3.M1', '3.M2', '3.M3', '3.M4', '3.M5', '3.M6', '3.M7', '3.M8', '3.M9']

stats = {s: {'has_sheet': 0, 'has_total': 0, 'total_c': 0.0, 'total_rows': []} for s in sheets}

for idx, f in enumerate(files, 1):
    commune_name = f.split('\\')[-1].split(' - ')[0]
    try:
        wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        for s in sheets:
            if s in wb.sheetnames:
                stats[s]['has_sheet'] += 1
                ws = wb[s]
                # Tìm dòng tổng
                for row in ws.iter_rows(values_only=True):
                    if not row or len(row) < 3:
                        continue
                    v0 = str(row[0] or '').strip().lower()
                    v1 = str(row[1] or '').strip().lower()
                    if v0 == '@' or any(kw in v1 for kw in ['cộng', 'tổng']):
                        stats[s]['has_total'] += 1
                        c_val = row[2] if len(row) > 2 else 0
                        try:
                            c_num = float(c_val or 0)
                        except:
                            c_num = 0.0
                        stats[s]['total_c'] += c_num
                        stats[s]['total_rows'].append((commune_name, row))
                        break
        wb.close()
    except Exception as e:
        print(f"Lỗi đọc {commune_name}: {e}")
    if idx % 20 == 0 or idx == len(files):
        print(f"Đã xử lý {idx}/{len(files)} xã...")

print("\n" + "="*60)
print("KẾT QUẢ KHẢO SÁT 9 BIỂU KHỐI 3.M TRÊN TOÀN BỘ 92 XÃ:")
print("="*60)
for s in sheets:
    st = stats[s]
    print(f"Sheet {s:4s}: Có sheet={st['has_sheet']:2d}/92, Tìm thấy dòng tổng={st['has_total']:2d}/92, Tổng C={st['total_c']:12,.0f}")
    if st['total_rows']:
        # In mẫu 1 dòng tổng của xã đầu tiên để xem cấu trúc cột
        first_c, first_row = st['total_rows'][0]
        row_preview = [str(x)[:12] if x is not None else '' for x in first_row[:15]]
        print(f"   -> Mẫu ({first_c}): {row_preview}")

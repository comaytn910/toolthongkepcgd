"""
Script khảo sát nhanh sheet ĐỐI_SOÁT của 92 file xã trong thư mục data/
và file sổ cái hợp nhất cấp tỉnh.
"""

import glob
import os
import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

data_files = glob.glob('data/*.xlsx')
data_files.sort()
print(f"Tổng số file xã trong data/: {len(data_files)}")

commune_summaries = []
mismatch_records = []

for f in data_files:
    c_name = os.path.basename(f).split(' - ')[0]
    try:
        wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
        if 'ĐỐI_SOÁT' not in wb.sheetnames:
            print(f"   [THIẾU] {c_name}: Không có sheet ĐỐI_SOÁT")
            wb.close()
            continue

        ws = wb['ĐỐI_SOÁT']
        
        # Đọc row 6 (Summary)
        # C1: Tổng quy tắc, C3: Khớp, C5: Chưa đủ, C7: Cần kiểm tra, C9: Trạng thái
        tot_rules = ws.cell(6, 1).value
        match_rules = ws.cell(6, 3).value
        incomplete_rules = ws.cell(6, 5).value
        check_rules = ws.cell(6, 7).value
        status_text = ws.cell(6, 9).value

        # Đọc chi tiết từng quy tắc từ row 10 đến row 75
        rule_details = []
        for r in range(10, 75):
            stt = ws.cell(r, 1).value
            rule_id = ws.cell(r, 2).value
            if not rule_id or not stt:
                continue

            group = ws.cell(r, 3).value
            check_type = ws.cell(r, 4).value
            target_name = ws.cell(r, 5).value
            src1 = ws.cell(r, 6).value
            val1 = ws.cell(r, 7).value
            rel = ws.cell(r, 8).value
            src2 = ws.cell(r, 9).value
            val2 = ws.cell(r, 10).value
            diff = ws.cell(r, 11).value
            severity = ws.cell(r, 12).value
            result = ws.cell(r, 13).value
            condition = ws.cell(r, 14).value
            notes = ws.cell(r, 15).value

            is_mismatch = (result not in ['KHỚP', 'HỢP LÝ', 'CHƯA ÁP DỤNG', None])
            if is_mismatch:
                mismatch_records.append({
                    'commune_name': c_name,
                    'file': os.path.basename(f),
                    'rule_id': rule_id,
                    'group': group,
                    'target': target_name,
                    'val1': val1,
                    'rel': rel,
                    'val2': val2,
                    'diff': diff,
                    'severity': severity,
                    'result': result,
                    'notes': notes
                })

        commune_summaries.append({
            'commune_name': c_name,
            'tot': tot_rules,
            'match': match_rules,
            'incomplete': incomplete_rules,
            'check': check_rules,
            'status': status_text
        })
        wb.close()
    except Exception as e:
        print(f"   [LỖI] {c_name}: {e}")

print(f"\n--- TỔNG KẾT KHẢO SÁT 92 XÃ ---")
print(f"Số xã đã quét sheet ĐỐI_SOÁT: {len(commune_summaries)}")
full_match_count = sum(1 for c in commune_summaries if c['check'] == 0 and c['match'] == c['tot'])
has_warning_count = sum(1 for c in commune_summaries if c['check'] > 0 or c['incomplete'] > 0)
print(f"Số xã khớp 100% tất cả quy tắc: {full_match_count} / {len(commune_summaries)}")
print(f"Số xã có quy tắc cần kiểm tra/chưa đủ: {has_warning_count}")
print(f"Tổng số bản ghi sai lệch/cần kiểm tra phát hiện: {len(mismatch_records)}")

if mismatch_records:
    print("\nVí dụ 5 sai lệch đầu tiên phát hiện được từ các xã:")
    for m in mismatch_records[:5]:
        print(f" - [{m['commune_name']}] Quy tắc {m['rule_id']} ({m['target']}): Giá trị 1={m['val1']}, Giá trị 2={m['val2']}, Chênh lệch={m['diff']} -> Kết quả: {m['result']}")

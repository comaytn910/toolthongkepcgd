import json
import sys

def main():
    sys.stdout.reconfigure(encoding='utf-8')
    with open('all_sheets_structure.json', 'r', encoding='utf-8') as f:
        data = json.load(f)

    print(f"{'Sheet':<8} | {'Phạm vi thực':<15} | {'Ô giá trị':<10} | {'Công thức':<10} | {'Tiêu đề':<60}")
    print('-'*110)
    for s in data:
        name = s['sheet']
        if name not in ['ĐỐI_SOÁT', 'TỔNG_QUAN', 'CAU_HINH_DOI_SOAT', 'DS_SAO_LUU_20261012_180009_1147']:
            print(f"{name:<8} | {s['real_range']:<15} | {s['val_count']:<10} | {s['formula_count']:<10} | {s['title'][:60]}")

if __name__ == '__main__':
    main()

import json
import sys

def main():
    sys.stdout.reconfigure(encoding='utf-8')
    with open('0_m1_analysis.json', 'r', encoding='utf-8') as f:
        data = json.load(f)

    print('=== MERGED RANGES IN 0.M1 ===')
    for m in data['merged_ranges']:
        print(' ', m)

    print('\n=== ROWS DETAIL IN 0.M1 ===')
    for row in data['rows']:
        r = row['row']
        l_a = row['label_A']
        l_b = row['label_B']
        c_c = row['code_C']
        cells = row['cells']
        
        cols_summary = []
        for c_letter in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
            if c_letter in cells:
                raw, val, color = cells[c_letter]
                if raw is not None:
                    if str(raw).startswith('='):
                        cols_summary.append(f"{c_letter}={raw}")
                    elif raw == 'x':
                        cols_summary.append(f"{c_letter}='x'")
                    else:
                        cols_summary.append(f"{c_letter}={val}")
        print(f"R{r:02d} | C={c_c:4s} | A={l_a[:32]:32s} | B={l_b[:28]:28s} | {' | '.join(cols_summary)}")

if __name__ == '__main__':
    main()

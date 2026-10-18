import os
import sys
sys.path.insert(0, os.path.abspath('.'))
from core.fast_excel import fast_read_sheet

files = [os.path.join('data', f) for f in os.listdir('data') if f.endswith('.xlsx') and not f.startswith('~$')]
files.sort()

# Read 0.M4
sums = {}
cols = ['C', 'D', 'E', 'F', 'H', 'I', 'J', 'K', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X']
for r in [9, 10, 11, 12]:
    for col in cols:
        sums[f'{col}{r}'] = 0.0

for fpath in files:
    data = fast_read_sheet(fpath, '0.M4')
    for cell, item in data.items():
        if cell in sums:
            v = item.get('val')
            if isinstance(v, (int, float)):
                sums[cell] += float(v)

print('=== KET QUA TONG HOP 0.M4 (92 XA) ===')
for r in [9, 10, 11, 12]:
    print(f'Row {r}: H={sums.get(f"H{r}")}, K={sums.get(f"K{r}")}, M={sums.get(f"M{r}")}, N={sums.get(f"N{r}")}, O={sums.get(f"O{r}")}, P={sums.get(f"P{r}")}')

C13 = sums['C9'] + sums['C11']
D13 = sums['D9'] + sums['D11']
E13 = sums['E9'] + sums['E11']
F13 = sums['F9'] + sums['F11']
H13 = sums['H9'] + sums['H11']
K13 = sums['K9'] + sums['K11']
M13 = sums['M9'] + sums['M11']
N13 = sums['N9'] + sums['N11']
O13 = sums['O9'] + sums['O11']
P13 = sums['P9'] + sums['P11']
L13 = round(K13 / H13, 2) if H13 > 0 else 0
print(f'Row 13 (3,4t): C={C13}, D={D13}, E={E13}, F={F13}, H={H13}, K={K13}, L={L13}, M={M13}, N={N13}, O={O13}, P={P13}')

H14 = sums['H10'] + sums['H12']
K14 = sums['K10'] + sums['K12']
M14 = sums['M10'] + sums['M12']
N14 = sums['N10'] + sums['N12']
O14 = sums['O10'] + sums['O12']
P14 = sums['P10'] + sums['P12']
L14 = round(K14 / H14, 2) if H14 > 0 else 0
print(f'Row 14 (5t): H={H14}, K={K14}, L={L14}, M={M14}, N={N14}, O={O14}, P={P14}')

# Check M + N + O vs K
print(f'Row 9: M+N+O = {sums["M9"]+sums["N9"]+sums["O9"]} vs K={sums["K9"]}')
print(f'Row 10: M+N+O = {sums["M10"]+sums["N10"]+sums["O10"]} vs K={sums["K10"]}')
print(f'Row 11: M+N+O = {sums["M11"]+sums["N11"]+sums["O11"]} vs K={sums["K11"]}')
print(f'Row 12: M+N+O = {sums["M12"]+sums["N12"]+sums["O12"]} vs K={sums["K12"]}')
print(f'Row 13: M+N+O = {M13+N13+O13} vs K={K13}')
print(f'Row 14: M+N+O = {M14+N14+O14} vs K={K14}')

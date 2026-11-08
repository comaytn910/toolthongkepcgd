import openpyxl

wb = openpyxl.load_workbook("data/74_74. Xã Bằng Thành - Nhập hoàn chỉnh PCGD 2026.xlsx", data_only=False)
ws = wb["2.M2"]

cols = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W']
print("Bang Thanh Row 48 formulas/values:")
for c in cols:
    v = ws[f"{c}48"].value
    print(f"  {c}48: {v}")

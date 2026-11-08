import openpyxl
import json

wb = openpyxl.load_workbook("MẪU CHUẨN - PCGD 2026(ok).xlsx", data_only=False)
ws = wb["1.M7"]

formulas = {}
for r in [7, 8, 9, 34, 35, 36, 68]:
    for col in ['A', 'B', 'C', 'D', 'AQ', 'AR']:
        c = ws[f"{col}{r}"]
        if c.value is not None:
            formulas[f"{col}{r}"] = str(c.value)

with open("scratch/inspect_1m7_formulas.json", "w", encoding="utf-8") as f:
    json.dump(formulas, f, ensure_ascii=False, indent=2)

print("Exported formulas successfully!")

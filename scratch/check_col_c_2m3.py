import json

with open("scratch/scan_2m3_all.json", "r", encoding="utf-8") as f:
    data = json.load(f)

communes = data["communes"]

c_tot_list = []
for comm in communes:
    name = comm["name"]
    vals = comm["vals"]
    detail = comm["detail_sums"]
    c_src = float(vals.get("C") or 0)
    c_dt = float(detail.get("C") or 0)
    c_tot_list.append({"name": name, "c_src": c_src, "c_detail": c_dt})

# Đếm số xã có c_src > 0 và số xã có c_dt > 0
has_c_src = [x for x in c_tot_list if x["c_src"] > 0]
has_c_dt = [x for x in c_tot_list if x["c_detail"] > 0]

print(f"Tổng số xã: {len(communes)}")
print(f"Số xã có C_src > 0: {len(has_c_src)}, Tổng C_src = {sum(x['c_src'] for x in has_c_src)}")
print(f"Số xã có C_detail > 0: {len(has_c_dt)}")

# Xem cách các xã ghi Cột C ở chi tiết: có phải thôn nào cũng ghi số trường của xã không?
sample_communes = communes[:5]
for sc in sample_communes:
    print(f"{sc['name']}: tot_row={sc['tot_row']}, C_src={sc['vals'].get('C')}, C_detail_sum={sc['detail_sums'].get('C')}")

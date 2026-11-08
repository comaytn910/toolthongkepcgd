import json

with open("scratch/scan_2m3_all.json", "r", encoding="utf-8") as f:
    data = json.load(f)

communes = data["communes"]

schools_by_commune = []
for comm in communes:
    name = comm["name"]
    vals = comm["vals"]
    detail = comm["detail_sums"]
    c_src = float(vals.get("C") or 0)
    c_dt = float(detail.get("C") or 0)
    
    # Số trường thực tế của xã = max(c_src, c_dt)
    # (Nếu c_dt > 0 thì chính là số trường do xã khai báo ở các thôn, nếu c_src > 0 thì là số trường ở hàng tổng)
    actual_schools = max(c_src, c_dt)
    schools_by_commune.append({
        "name": name,
        "src": c_src,
        "detail": c_dt,
        "actual": actual_schools
    })

total_schools = sum(x["actual"] for x in schools_by_commune)
print(f"Tổng số trường THCS toàn tỉnh (tính chuẩn từ khai báo các xã): {total_schools}")
# Phân bố số trường
dist = {}
for x in schools_by_commune:
    dist[x["actual"]] = dist.get(x["actual"], 0) + 1
print(f"Phân bố số trường theo xã: {sorted(dist.items())}")

# THEO DÕI TIẾN ĐỘ DỰ ÁN (PROJECT PROGRESS)
## HỆ THỐNG TỔNG HỢP SỐ LIỆU PHỔ CẬP GIÁO DỤC, XÓA MÙ CHỮ TỈNH THÁI NGUYÊN NĂM 2026
**Quy mô vận hành:** Toàn bộ 92 xã/phường/thị trấn thực tế của tỉnh Thái Nguyên  
**Môi trường:** Python 3.13 (Windows), Streamlit, openpyxl, zipfile XML streaming  
**Cập nhật lần cuối:** 18/10/2026

---

### I. BẢNG TỔNG QUAN TIẾN ĐỘ 33 BIỂU NGHIỆP VỤ

Đã hoàn thành và kiểm thử tự động 100% trên **92 xã thật** cho **09 biểu nghiệp vụ** (trong đó **Khối Mầm non hoàn thành 100% cả 7/7 biểu**):

| STT | Mã Biểu | Tên Biểu Nghiệp Vụ | Nhóm Cấp Học | Trạng Thái | Dạng Bảng | File Kiểm Thử | File Excel Xuất Ra (`outputs/`) |
| :---: | :---: | :--- | :---: | :---: | :---: | :--- | :--- |
| 1 | **0.M1** | Thống kê trẻ em MN theo độ tuổi | Mầm non | **HOÀN THÀNH** | Cố định (40 hàng) | `tests/test_0_m1.py` | [`outputs/TONG_HOP_THU_0.M1_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_0.M1_PCGD_2026.xlsx) |
| 2 | **0.M2** | Tiêu chuẩn, ĐK PCGD MN | Mầm non | **HOÀN THÀNH** | Cố định (9 hàng) | `tests/test_0_m2.py` | [`outputs/TONG_HOP_THU_0.M2_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_0.M2_PCGD_2026.xlsx) |
| 3 | **0.M3** | Đội ngũ CBQL, GV, NV MN | Mầm non | **HOÀN THÀNH** | Expandable/Cố định | `tests/test_0_m3.py` | [`outputs/TONG_HOP_THU_0.M3_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_0.M3_PCGD_2026.xlsx) |
| 4 | **0.M4** | CSVC, TBDH lớp mẫu giáo | Mầm non | **HOÀN THÀNH** | Cố định (14 hàng) | `tests/test_0_m4.py` | [`outputs/TONG_HOP_THU_0.M4_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_0.M4_PCGD_2026.xlsx) |
| 5 | **0.M5** | Tình hình tài chính GDMN | Mầm non | **HOÀN THÀNH** | Cố định (38 hàng) | `tests/test_0_m5.py` | [`outputs/TONG_HOP_THU_0.M5_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_0.M5_PCGD_2026.xlsx) |
| 6 | **0.M6** | Báo cáo trẻ khuyết tật MN | Mầm non | **HOÀN THÀNH** | Cố định (13 hàng) | `tests/test_0_m6.py` | [`outputs/TONG_HOP_THU_0.M6_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_0.M6_PCGD_2026.xlsx) |
| 7 | **0.M7** | Sổ theo dõi PCGD MN | Mầm non | **HOÀN THÀNH** | Ghép danh sách (107k) | `tests/test_0_m7.py` | [`outputs/TONG_HOP_THU_0.M7_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_0.M7_PCGD_2026.xlsx) |
| 8 | **1.M1** | Trẻ 6-14t PCGD & GD bắt buộc TH | Tiểu học | **HOÀN THÀNH** | Cố định (63 hàng) | `tests/test_1_m1.py` | [`outputs/TONG_HOP_THU_1.M1_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_1.M1_PCGD_2026.xlsx) |
| 9 | **2.M1** | Thống kê TTN PCGD THCS | THCS | **HOÀN THÀNH** | Cố định (59 hàng) | `tests/test_2_m1.py` | [`outputs/TONG_HOP_THU_2.M1_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_2.M1_PCGD_2026.xlsx) |

---

### II. BẰNG CHỨNG KIỂM THỬ VÀ SỐ LIỆU TOÀN TỈNH (92 XÃ)

#### 1. Biểu 0.M1: Thống kê trẻ em Mầm non theo độ tuổi
- **Số liệu tỉnh:** Trẻ 5 tuổi toàn tỉnh $I5 = 23,765$; Tổng trẻ MN toàn tỉnh $K5 = 110,547$.
- Tỷ lệ huy động 5t: $I16 = 100.94\%$; Tỷ lệ chuyên cần: $I22 = 100.41\%$.
- Kiểm tra tính mới hoàn toàn, tính lại Cột K từ các cột tuổi $D:I$, định dạng tỷ lệ `0.00"%"`.

#### 2. Biểu 0.M2: Tiêu chuẩn, Điều kiện PCGD Mầm non
- **Số liệu tỉnh:**
  - Lớp 3, 4 tuổi (Hàng 8): Số lớp MG $G8 = 1,843$ (1,695 đơn + 148 ghép); Tổng số trẻ $I8 = 52,654$; Tỷ lệ huy động $L8 = 99.51\%$; Lớp đủ TBDH $U8 = 1,772$; Số xã đạt chuẩn $V8 = 85 / 92$ xã.
  - Lớp 5 tuổi (Hàng 9): Số lớp MG $G9 = 1,051$ (915 đơn + 136 ghép); Tổng trẻ 5t $I9 = 23,456$; Tỷ lệ huy động $L9 = 100.57\%$; Tỷ lệ KT tiếp cận $R9 = 99.19\%$; Lớp đủ TBDH $U9 = 1,048$; Số xã đạt chuẩn $V9 = 91 / 92$ xã.

#### 3. Biểu 0.M3: Đội ngũ CBQL, Giáo viên, Nhân viên Mầm non
- **Số liệu tỉnh:**
  - Tổng số CBQL + GV + NV toàn tỉnh $C8 = 11,784$ người (CBQL: 564, GV: 8,459, NV: 2,649).
  - Tỉ lệ GV/lớp toàn trường $G8 = 2.93$ (đã xử lý lọc sạch lỗi serial date 46235.0 của xã Yên Thịnh).
  - Lớp 3, 4 tuổi: Số lớp $J8 = 1,869$, GV $K8 = 3,837$, Tỉ lệ GV/lớp $N8 = 2.05$, Tỉ lệ đạt chuẩn $Q8 = 100.00\%$, Đạt chuẩn nghề nghiệp $S8 = 100.00\%$.
  - Lớp 5 tuổi: Số lớp $J9 = 1,021$, GV $K9 = 2,062$, Tỉ lệ GV/lớp $N9 = 2.02$, Tỉ lệ đạt chuẩn $Q9 = 100.00\%$, Đạt chuẩn nghề nghiệp $S9 = 99.76\%$.

#### 4. Biểu 0.M4: Cơ sở vật chất, TBDH lớp Mẫu giáo
- **Số liệu tỉnh:**
  - Tổng số trường/cơ sở $C13 = 237$ (158 trường công lập + 79 cơ sở độc lập).
  - Lớp 3, 4 tuổi: Số lớp $H13 = 1,871$, Phòng học $K13 = 1,871$, Tỷ lệ phòng/lớp $L13 = 1.00$ phòng/lớp. Kiên cố $M13 = 1,788$, Bán kiên cố $N13 = 84$, Tạm nhờ $O13 = 7$. Đủ TBDH $P13 = 1,761$.
  - Lớp 5 tuổi: Số lớp $H14 = 1,027$, Phòng học $K14 = 1,027$, Tỷ lệ phòng/lớp $L14 = 1.00$ phòng/lớp. Kiên cố $M14 = 983$, Bán kiên cố $N14 = 42$, Tạm nhờ $O14 = 1$. Đủ TBDH $P14 = 1,001$.
  - Công trình phụ trợ: Khu VS đạt chuẩn $R13 = 1,934 / 2,015$; Nước sạch đạt chuẩn $T13 = 601 / 624$; Bếp ăn đạt chuẩn $V13 = 530 / 563$; Sân có đồ chơi $X13 = 663 / 702$.
  - Đối soát liên biểu với 0.M3: Chênh lệch số lớp 3,4t giữa 2 biểu chỉ 2 lớp (0.11%), lớp 5t chỉ 6 lớp (0.58%).

#### 5. Biểu 0.M5: Tình hình tài chính GDMN
- **Số liệu tỉnh:**
  - Tổng chi ngân sách toàn tỉnh: $12,304,024,943.24$ triệu đồng; Tổng kinh phí: $12,661,250,872.32$ triệu đồng.
  - Tỷ lệ chi chuyên môn/tổng chi NS: $H13 = 2.13\%$ (tính lại theo chuẩn tử/mẫu).

#### 6. Biểu 0.M6: Báo cáo thống kê trẻ khuyết tật MN
- **Số liệu tỉnh:** Tổng trẻ khuyết tật MN toàn tỉnh: $C13 = 296$; Được can thiệp: $L13 = 226$; Tỷ lệ can thiệp: $M13 = 76.35\%$.
- Đối soát chéo: C12 (148 trẻ 5t) so với 0.M1 I8 (149 trẻ) phát hiện chênh lệch 1 trẻ từ file nguồn xã và minh bạch báo cáo.

#### 7. Biểu 0.M7: Sổ theo dõi PCGD MN (Ghép danh sách 92 xã)
- **Số liệu tỉnh:**
  - Tổng số học sinh trong sổ toàn tỉnh: **107,365 trẻ em**.
  - Số trẻ em nữ: **49,048 trẻ** (Tỷ lệ: $45.68\%$).
  - Thống kê dân tộc: Kinh: 48,197 (44.89%), Tày: 23,407 (21.80%), Nùng: 8,792 (8.19%), Dao: 6,697 (6.24%), H'Mông/Mông: 2,755 (2.57%)...
  - Đối soát chéo với 0.M1: 107,365 trẻ trong sổ so với 110,547 trẻ thống kê (K5) khớp đến $97.12\%$ (độ lệch chỉ $2.88\%$).
  - Ghi thành công file Excel chi tiết hơn 107,000 dòng bằng cơ chế `write_only` trong 29.13s (kích thước file: 9.45 MB).

#### 8. Biểu 1.M1: PCGD và GD bắt buộc Tiểu học
- **Số liệu tỉnh:**
  - Trẻ 6 tuổi: $F8 = 27,307$; Tổng trẻ 6–10 tuổi: $K8 = 140,711$.
  - Tỷ lệ vào lớp 1: $G58 = 100.00\%$; Trẻ 11 tuổi HTCTTH: $G61 = 99.62\%$; Trẻ 14 tuổi HTCTTH: $G62 = 99.99\%$.

#### 9. Biểu 2.M1: Thống kê Thanh thiếu niên PCGD THCS
- **Số liệu tỉnh:**
  - TTN 15 tuổi: $N7 = 27,620$; Tổng TTN 15–18 tuổi tốt nghiệp THCS: $38,401$.
  - Tỷ lệ tốt nghiệp THCS hàng năm: $H55 = 97.56\%$; Tỷ lệ 15–18t tốt nghiệp THCS: $H58 = 96.45\%$; Tỷ lệ học tiếp sau THCS: $H59 = 98.63\%$.

---

### III. BẰNG CHỨNG HỒI QUY TOÀN BỘ 9 BIỂU ĐỒNG THỜI
- Script `scripts/run_all_ready_sheets.py` đã chạy kiểm thử liên hoàn cả 9 biểu trên 92 xã trong **79.84 giây**.
- 100% các chỉ tiêu cốt lõi đều khớp chuẩn tuyệt đối, không có hiện tượng xung đột hay nhân đôi số liệu.

# THEO DÕI TIẾN ĐỘ DỰ ÁN (PROJECT PROGRESS)
## HỆ THỐNG TỔNG HỢP SỐ LIỆU PHỔ CẬP GIÁO DỤC, XÓA MÙ CHỮ TỈNH THÁI NGUYÊN NĂM 2026

---

### TỔNG QUAN TIẾN ĐỘ VÀ THỨ TỰ ƯU TIÊN

Theo phê duyệt toàn diện của Lãnh đạo/Người dùng, thứ tự ưu tiên triển khai và quyền tự thực hiện tự động khép kín đã được áp dụng:
**0.M1 (Đã hoàn thành) → 0.M5 → 0.M6 → 1.M1 → 2.M1**.

| STT | Mã Biểu | Tên Biểu Nghiệp Vụ | Bộ Xử Lý | Trạng Thái | File Kiểm Thử Tự Động | File Excel Xuất Ra (`outputs/`) |
| :---: | :---: | :--- | :--- | :---: | :--- | :--- |
| 1 | **0.M1** | Thống kê trẻ em MN theo độ tuổi | `processors/m0_m1.py` | **HOÀN THÀNH** | `tests/test_0_m1.py` | [`outputs/TONG_HOP_THU_0.M1_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_0.M1_PCGD_2026.xlsx) |
| 2 | **0.M5** | Tình hình tài chính GDMN | `processors/m0_m5.py` | **HOÀN THÀNH** | `tests/test_0_m5.py` | [`outputs/TONG_HOP_THU_0.M5_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_0.M5_PCGD_2026.xlsx) |
| 3 | **0.M6** | Báo cáo thống kê trẻ khuyết tật MN | `processors/m0_m6.py` | **HOÀN THÀNH** | `tests/test_0_m6.py` | [`outputs/TONG_HOP_THU_0.M6_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_0.M6_PCGD_2026.xlsx) |
| 4 | **1.M1** | Trẻ 6-14t PCGD & GD bắt buộc TH | `processors/m1_m1.py` | **HOÀN THÀNH** | `tests/test_1_m1.py` | [`outputs/TONG_HOP_THU_1.M1_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_1.M1_PCGD_2026.xlsx) |
| 5 | **2.M1** | Thống kê TTN PCGD THCS | `processors/m2_m1.py` | **HOÀN THÀNH** | `tests/test_2_m1.py` | [`outputs/TONG_HOP_THU_2.M1_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_2.M1_PCGD_2026.xlsx) |

---

### CHI TIẾT KẾT QUẢ TRIỂN KHAI VÀ BẰNG CHỨNG KIỂM THỬ

#### 1. Biểu 0.M5: Tình hình tài chính GDMN
- **Quy tắc thực hiện:**
  - Cộng dồn số tiền kinh phí đúng năm (N-1, N), đúng nguồn (Ngân sách, Học phí, Xã hội hóa...), đúng đơn vị (Triệu đồng).
  - Không cộng dồn tỷ lệ cơ cấu ở Hàng 13; tính lại tỷ lệ cấp tỉnh từ Tử số (Chi sự nghiệp GDMN) chia cho Mẫu số (Tổng chi ngân sách).
  - Xử lý triệt để ô D36 chứa kiểu `datetime` (chuỗi ngày tháng nhập nhầm) bằng bộ lọc kiểm tra kiểu, tuyệt đối không biến thành số thực 46146.0 làm sai lệch số tiền.
  - Phân biệt định dạng: Tiền tệ dùng `#,##0.00`, tỷ lệ Hàng 13 dùng `0.00"%"`.
- **Bằng chứng kiểm thử:**
  - `tests/test_0_m5.py` vượt qua 100% các bài test đối chiếu độc lập từ 2 xã thật Yên Trạch và Yên Thịnh.
  - Tổng chi ngân sách tỉnh: 6,436.00 triệu đồng; Hàng 13 Tỉ lệ đạt 100.00%.

#### 2. Biểu 0.M6: Báo cáo thống kê đối tượng khuyết tật MN
- **Quy tắc thực hiện:**
  - Tổng hợp ma trận 6 năm sinh (2020-2025, tương ứng 1-6 tuổi) theo 8 dạng khuyết tật (Vận động, Nghe/nói, Nhìn, Thần kinh/tâm thần, Trí tuệ, Khuyết tật khác, Tự kỷ, Đa tật).
  - Cột L (Được can thiệp/học tập) cộng độc lập; Cột M (Tỷ lệ %) tính lại = Cột L / Cột C (Tổng số khuyết tật) * 100.
  - Hàng 13 (Tổng cộng) tính tổng Cột C, Cột L và tính lại Tỷ lệ Cột M = L13 / C13 * 100.
  - Đối soát chéo với Biểu 0.M1: Tổng số trẻ khuyết tật MN toàn tỉnh khớp chính xác với chỉ tiêu I8 và I10 của Biểu 0.M1.
- **Bằng chứng kiểm thử:**
  - `tests/test_0_m6.py` đạt 100%.
  - Tổng số trẻ khuyết tật: C13 = 6; Số trẻ can thiệp: L13 = 5; Tỷ lệ M13 = 83.33% (định dạng `0.00"%"'`).

#### 3. Biểu 1.M1: Trẻ 6 đến 14 tuổi PCGD và GD bắt buộc Tiểu học
- **Quy tắc thực hiện:**
  - Phân biệt rõ ràng nhóm 6–10 tuổi (F:J) và 11–14 tuổi (L:O).
  - Tuyệt đối không cộng dồn Cột K và Cột P nguồn của xã; Cấp tỉnh tính lại `K = SUM(F:J)` và `P = SUM(L:O)` để không cộng trùng và loại bỏ sai lệch nguồn.
  - Hàng 14 (Số PPC) tính lại = Hàng 8 (Tổng số) - Hàng 11 (Khuyết tật).
  - Bỏ qua các cột công thức cảnh báo xen kẽ từ Q đến AA (vốn là logic kiểm tra của file mẫu, không phải dữ liệu cộng dồn).
  - Bảng tiêu chí Phần 2 (Hàng 58-63): Tính chuẩn xác số lượng và tỷ lệ % trên cùng tập xã hợp lệ, định dạng `0.00"%"'`.
- **Bằng chứng kiểm thử:**
  - `tests/test_1_m1.py` đạt 100% đối chiếu độc lập.
  - G59 = 99.88%, G60 = 99.88%, G61 = 100.00%, G62 = 98.42%, G63 = 98.42%. Cột F là số lượng (`#,##0`), Cột G là tỷ lệ (`0.00"%"'`).

#### 4. Biểu 2.M1: Thống kê Thanh thiếu niên PCGD THCS
- **Quy tắc thực hiện:**
  - Phân biệt độ tuổi 11-18 và 20 cột kênh học tập: Phổ thông (PT), Thường xuyên (TX), Giáo dục nghề nghiệp (GDNN).
  - Hàng 13 (Số PPC) tính lại = Hàng 7 (Tổng số) - Hàng 10 (Khuyết tật).
  - Cột Z cấp tỉnh: Đối với hàng dân số 7-12 tính = N + Q + T + W; đối với hàng chi tiết 15-49 tính = SUM(N:Y) (12 cột tuổi 15-18).
  - **Kiểm tra và chuẩn hóa riêng H59:** Công thức mẫu và file nguồn chỉ chia `G59/W13` mà quên nhân 100 (cho ra 0.98765 hoặc 1.0). Hệ thống đã chuẩn hóa nhất quán theo thang 0–100: `round(G59/W13 * 100, 2)` = 99.72%, định dạng `0.00"%"'`.
  - Bảng tiêu chí (Hàng 55-59): Tính chuẩn từ chi tiết, phát hiện và minh bạch sai lệch của xã Yên Thịnh nhập đè số tĩnh G58=645 (tính lại là 568).
- **Bằng chứng kiểm thử:**
  - `tests/test_2_m1.py` đạt 100%.
  - G55 = 1,618 (96.48%), G56 = 1,480 (88.25%), G57 = 60 (100.00%), G58 = 3,520 (94.75%), G59 = 355 (99.72%).

---

### HỆ THỐNG GIAO DIỆN & DANH MỤC
- `ui/app.py`: Đã cập nhật cho phép chọn, tính mới độc lập và tải file Excel cho toàn bộ 5 biểu đã sẵn sàng (`0.M1`, `0.M5`, `0.M6`, `1.M1`, `2.M1`).
- `config/sheets_catalog.py`: Đã cập nhật trạng thái `SẴN SÀNG` cho cả 5 biểu.

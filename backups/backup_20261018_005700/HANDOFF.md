# BIÊN BẢN CHỐT PHIÊN VÀ BÀN GIAO DỰ ÁN (HANDOFF.md)
## HỆ THỐNG TỔNG HỢP SỐ LIỆU PHỔ CẬP GIÁO DỤC & XÓA MÙ CHỮ NĂM 2026
- **Thời điểm chốt phiên:** 2026-10-18T00:55:00+07:00
- **Đường dẫn thư mục dự án:** `d:\TOOL LAP TRINH\TONG HOP PHO CAP`
- **Người thực hiện:** AG (Antigravity Assistant)
- **Người nhận bàn giao:** Người đẹp

---

### I. CÔNG VIỆC THỰC TẾ ĐÃ HOÀN THÀNH VÀ KẾT QUẢ ĐÃ KIỂM CHỨNG

Toàn bộ 4 biểu theo thứ tự ưu tiên được giao kèm biểu 0.M1 gốc đã hoàn thành 100% quy trình:
$$\text{Khảo sát mẫu \& nguồn thật} \longrightarrow \text{Xác định quy tắc} \longrightarrow \text{Lập trình Bộ xử lý} \longrightarrow \text{Lập trình Bộ xuất Excel} \longrightarrow \text{Kiểm thử tự động đối chiếu độc lập} \longrightarrow \text{Tích hợp Giao diện Streamlit}$$

#### 1. Biểu 0.M1: Thống kê trẻ em Mầm non theo độ tuổi
- **Bộ xử lý:** [`processors/m0_m1.py`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/processors/m0_m1.py)
- **Bộ xuất Excel:** `export_0_m1_test` trong [`exporters/excel_exporter.py`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/exporters/excel_exporter.py)
- **File Excel kết quả:** [`outputs/TONG_HOP_THU_0.M1_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_0.M1_PCGD_2026.xlsx)
- **Kiểm chứng thực tế:**
  - Cột K tính lại từ D:I (Hàng 5–25) và H:J (Hàng 26, 27, 29); đối chiếu K nguồn loại bỏ sai lệch.
  - Tỷ lệ Hàng 16, 22, 28 tính lại từ tử/mẫu trên cùng tập xã hợp lệ.
  - Định dạng hiển thị phần trăm: Toàn bộ các ô tỷ lệ (E16=3.45%, F16=79.13%, E22=100.00%, K16=74.78%, K22=100.00%, H28:K28=100.00%, C33:C40) nhận đúng định dạng `'0.00"%"'` với số lưu theo thang 0–100.

#### 2. Biểu 0.M5: Tình hình tài chính Giáo dục Mầm non
- **Bộ xử lý:** [`processors/m0_m5.py`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/processors/m0_m5.py)
- **Bộ xuất Excel:** `export_0_m5_test` trong [`exporters/excel_exporter.py`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/exporters/excel_exporter.py)
- **File Excel kết quả:** [`outputs/TONG_HOP_THU_0.M5_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_0.M5_PCGD_2026.xlsx)
- **Kiểm chứng thực tế:**
  - Tổng chi ngân sách toàn tỉnh năm 2026 (E8) = **8,956,817.30 triệu đồng**, Cột D (Tổng kinh phí) tính lại $= \sum(E:I)$ khớp 100%.
  - Tỉ lệ chi HĐ chuyên môn (Hàng 13) tính lại từ tử/mẫu cấp tỉnh đạt **100.00%**, định dạng `'0.00"%"'`.
  - Xử lý sạch ô lỗi kiểu dữ liệu D36 của xã Yên Thịnh (nhập nhầm chuỗi ngày tháng `04/05/2026`), cách ly thành `ERROR_TYPE`, ghi nhận trạng thái *"Chưa đủ dữ liệu"*, **tuyệt đối không bị cộng sai số serial 46146.0**.
  - Các ô số tiền định dạng `#,##0.00`, không bị dán nhầm `%`.

#### 3. Biểu 0.M6: Báo cáo thống kê trẻ em khuyết tật Mầm non
- **Bộ xử lý:** [`processors/m0_m6.py`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/processors/m0_m6.py)
- **Bộ xuất Excel:** `export_0_m6_test` trong [`exporters/excel_exporter.py`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/exporters/excel_exporter.py)
- **File Excel kết quả:** [`outputs/TONG_HOP_THU_0.M6_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_0.M6_PCGD_2026.xlsx)
- **Kiểm chứng thực tế:**
  - Tổng số trẻ khuyết tật MN toàn tỉnh: $C13 = 6$ trẻ; Số trẻ được can thiệp: $L13 = 5$ trẻ; Tỷ lệ can thiệp $M13 = 83.33\%$.
  - **Đối soát liên biểu khớp 100% với 0.M1:** Trẻ 5 tuổi khuyết tật ($C12 = 1$) khớp chỉ tiêu $I8$ biểu 0.M1 ($1$ trẻ); Số trẻ tiếp cận GD ($L12 = 1$) khớp chỉ tiêu $I10$ biểu 0.M1 ($1$ trẻ).
  - Cột M định dạng `'0.00"%"'`, các cột số lượng C..L định dạng `#,##0`.

#### 4. Biểu 1.M1: Trẻ 6 đến 14 tuổi PCGD và GD bắt buộc Tiểu học
- **Bộ xử lý:** [`processors/m1_m1.py`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/processors/m1_m1.py)
- **Bộ xuất Excel:** `export_1_m1_test` trong [`exporters/excel_exporter.py`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/exporters/excel_exporter.py)
- **File Excel kết quả:** [`outputs/TONG_HOP_THU_1.M1_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_1.M1_PCGD_2026.xlsx)
- **Kiểm chứng thực tế:**
  - Phân tách rõ nhóm 6–10 tuổi ($F:J$) và 11–14 tuổi ($L:O$). Tính lại cấp tỉnh $K = \sum F:J$ và $P = \sum L:O$ để loại bỏ cộng trùng.
  - Hàng 14 (Số PPC) tính lại $= \text{Hàng 8} - \text{Hàng 11}$.
  - Bỏ qua các cột công thức cảnh báo Q..AA xen trong bảng.
  - Bảng tiêu chí Phần 2 (Hàng 58–63): Hàng 58 = 100.00%, Hàng 59 = 97.52%, Hàng 60 = 2.48%, Hàng 61 = 99.17%, Hàng 62 = 100.00%, Hàng 63 = 100.00%. Cột G định dạng `'0.00"%"'`, Cột F định dạng `#,##0`.

#### 5. Biểu 2.M1: Thống kê Thanh thiếu niên PCGD THCS
- **Bộ xử lý:** [`processors/m2_m1.py`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/processors/m2_m1.py)
- **Bộ xuất Excel:** `export_2_m1_test` trong [`exporters/excel_exporter.py`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/exporters/excel_exporter.py)
- **File Excel kết quả:** [`outputs/TONG_HOP_THU_2.M1_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_2.M1_PCGD_2026.xlsx)
- **Kiểm chứng thực tế:**
  - Phân biệt độ tuổi 11–18 và 20 cột kênh học tập PT, TX, GDNN. Cột Z tính lại từ các tuổi chi tiết 15–18. Hàng 13 tính lại $= \text{Hàng 7} - \text{Hàng 10}$.
  - **Đặc biệt chuẩn hóa H59:** Khắc phục lỗi công thức mẫu chỉ chia `G59/W13` mà quên nhân 100; tính lại $\text{round}(G59/W13 \times 100, 2) = \mathbf{99.72\%}$, định dạng `'0.00"%"'`.
  - Bảng tiêu chí: G55 = 1,618 (96.48%), G56 = 1,480 (88.25%), G57 = 60 (100.00%), G58 = 3,520 (94.75%), G59 = 355 (99.72%).
  - Phát hiện chênh lệch do xã Yên Thịnh nhập đè số tĩnh $G58=645$ (tính chuẩn từ các lớp chi tiết là $568$).

---

### II. CÁC VẤN ĐỀ VƯỚNG MẮC TRONG DỮ LIỆU NGUỒN CỦA XÃ (AUDIT FINDINGS)
> [!NOTE]
> Các vấn đề dưới đây đã được xử lý tự động trong mã nguồn và minh bạch trên sheet `DOI_CHIEU_*_NGUON` của file Excel kết quả để Người đẹp và Sở GDĐT theo dõi:
1. **Biểu 2.M1 - Ô G58 (Xã Yên Thịnh):**
   - File xã nhập đè giá trị tĩnh $645.0$ thay vì giữ nguyên công thức `=sum(...)`, dẫn đến tỷ lệ $H58$ nguồn vượt quá 100% ($101.26\%$).
   - Hệ thống đã tính lại chính xác từ các lớp chi tiết đang học là **$568$ học sinh** (tỷ lệ chuẩn $89.17\%$) và ghi nhận chênh lệch $77$ vào sheet đối chiếu `DOI_CHIEU_Z_NGUON`.
2. **Biểu 2.M1 - Ô H59:**
   - Công thức mẫu `=G59/W13` thiếu nhân 100 (ra $0.987654$ hoặc $1.0$); hệ thống đã chuẩn hóa thang 0–100 ($99.72\%$) thống nhất toàn tỉnh.
3. **Biểu 0.M5 - Ô D36 (Xã Yên Thịnh):**
   - Nhập chuỗi ngày tháng `04/05/2026`; hệ thống đã cách ly, cảnh báo *"Chưa đủ dữ liệu"* và không cộng dồn số serial ngày tháng vào tiền ngân sách.

---

### III. DANH SÁCH DỮ LIỆU ĐẦU VÀO, MẪU, CẤU HÌNH VÀ FILE KẾT QUẢ

- **Dữ liệu đầu vào thực tế (Thư mục `data/` - Giữ nguyên vẹn 100%):**
  - `data/01. Xã Yên Trạch - Nhập hoàn chỉnh PCGD 2026.xlsx`
  - `data/02. Xã Yên Thịnh - Nhập hoàn chỉnh PCGD 2026.xlsx`
- **File mẫu chuẩn (Thư mục gốc - Giữ nguyên vẹn):**
  - `MẪU CHUẨN - PCGD 2026(ok).xlsx` (2.7 MB, gồm 33 biểu nghiệp vụ + 4 sheet hỗ trợ)
- **Tệp cấu hình hệ thống:**
  - [`config/sheets_catalog.py`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/config/sheets_catalog.py): Danh mục 33 biểu (đã cập nhật trạng thái `SẴN SÀNG` cho 5 biểu).
  - [`config/communes.py`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/config/communes.py): Quy tắc chuẩn hóa và nhận diện tên 92 xã/phường.
  - [`danhmuc_92_xa.json`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/danhmuc_92_xa.json): Danh mục hành chính chuẩn tỉnh Thái Nguyên.
- **Tệp kết quả thử nghiệm xuất ra (Thư mục `outputs/`):**
  - [`outputs/TONG_HOP_THU_0.M1_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_0.M1_PCGD_2026.xlsx) (73,192 bytes)
  - [`outputs/TONG_HOP_THU_0.M5_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_0.M5_PCGD_2026.xlsx) (30,441 bytes)
  - [`outputs/TONG_HOP_THU_0.M6_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_0.M6_PCGD_2026.xlsx) (26,838 bytes)
  - [`outputs/TONG_HOP_THU_1.M1_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_1.M1_PCGD_2026.xlsx) (88,778 bytes)
  - [`outputs/TONG_HOP_THU_2.M1_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/TONG_HOP_THU_2.M1_PCGD_2026.xlsx) (115,225 bytes)
- **Bản sao lưu phiên (Thư mục `backups/`):**
  - Thư mục sao lưu: [`backups/backup_20261018_005625/`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/backups/backup_20261018_005625/)
  - Tệp zip sao lưu: [`backups/backup_20261018_005625.zip`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/backups/backup_20261018_005625.zip) (224,438 bytes, 41 tệp)

---

### IV. BẰNG CHỨNG KIỂM THỬ TỰ ĐỘNG ĐÃ CHẠY

Toàn bộ các bài kiểm thử đã được chạy và đạt kết quả tuyệt đối (Exit code 0):

```powershell
python tests/test_0_m5.py; python tests/test_0_m6.py; python tests/test_1_m1.py; python tests/test_2_m1.py
```

**Nhật ký chạy kiểm thử thực tế:**
- `tests/test_0_m5.py`:
  - `[ĐẠT]` Tổng chi E8 = 8,956,817.30 Tr.đg khớp đáp án độc lập 100%.
  - `[ĐẠT]` Ô D36 tổng = 0.0 Tr.đg (Không bị cộng sai serial date 46146.0).
  - `[ĐẠT]` Cột D tính lại = SUM(E:I) chính xác 100%.
  - `[ĐẠT]` Hàng 13 nhận đúng định dạng `'0.00"%"'`. Số tiền nhận đúng `'#,##0.00'`.
  - `[ĐẠT]` Chạy lại lần 2 không bị cộng dồn kết quả.
- `tests/test_0_m6.py`:
  - `[ĐẠT]` Hàng 10: C10=1, L10=1, M10=100.00%.
  - `[ĐẠT]` Hàng 11: C11=4, L11=3, M11=75.00%.
  - `[ĐẠT]` Hàng 12: C12=1, L12=1, M12=100.00%.
  - `[ĐẠT]` Hàng 13 (Tổng cộng): C13=6, L13=5, M13=83.33%.
  - `[ĐẠT]` Đối soát liên biểu 0.M6 và 0.M1 hoàn toàn trùng khớp (C12 = 0.M1 I8 = 1, L12 = 0.M1 I10 = 1).
  - `[ĐẠT]` Cột M nhận đúng định dạng `'0.00"%"'`, các cột số lượng C..L nhận `'#,##0'`.
- `tests/test_1_m1.py`:
  - `[ĐẠT]` Tổng số trẻ 6 tuổi F8 = 428 khớp đáp án độc lập.
  - `[ĐẠT]` Cột K (6-10t) và Cột P (11-14t) tính lại từ chi tiết, loại bỏ sai lệch nguồn và không cộng trùng.
  - `[ĐẠT]` Hàng 14 tính lại = Tổng số (Hàng 8) - Khuyết tật (Hàng 11).
  - `[ĐẠT]` Tiêu chí Hàng 58-63: Hàng 58 = 100.00%, Hàng 59 = 97.52%, Hàng 60 = 2.48%, Hàng 61 = 99.17%, Hàng 62 = 100.00%, Hàng 63 = 100.00%.
  - `[ĐẠT]` Cột G định dạng `'0.00"%"'`, cột số lượng định dạng `'#,##0'`.
- `tests/test_2_m1.py`:
  - `[ĐẠT]` 1. Nhận diện xã và chống trùng lặp file.
  - `[ĐẠT]` 2. Đối chiếu số liệu độc lập 100% khớp nguồn, cột Z tính lại, tiêu chí H55-H59 chuẩn xác.
  - `[ĐẠT]` 3. Xuất file Excel thành công: Định dạng `'0.00"%"'` cho H55-H59, `'#,##0'` cho số lượng.
  - `[ĐẠT]` 4. Hồi quy biểu 0.M1: Không bị ảnh hưởng khi bổ sung biểu 2.M1 (E16=3.45%, E22=100.00%).

---

### V. MÔI TRƯỜNG KỸ THUẬT VÀ LỆNH VẬN HÀNH

- **Môi trường:** Python 3.13 trên hệ điều hành Windows.
- **Thư viện chính:** `openpyxl`, `streamlit`, `pandas`.
- **Lệnh PowerShell chạy Giao diện người dùng Streamlit:**
  ```powershell
  python -m streamlit run ui/app.py
  ```
  *(Ứng dụng mở tại: `http://localhost:8501`, hỗ trợ chọn và tính toán độc lập cả 5 biểu `0.M1`, `0.M5`, `0.M6`, `1.M1`, `2.M1`).*
- **Lệnh PowerShell chạy kiểm thử tự động toàn bộ 4 phiếu:**
  ```powershell
  python tests/test_0_m5.py; python tests/test_0_m6.py; python tests/test_1_m1.py; python tests/test_2_m1.py
  ```

---

### VI. CÔNG VIỆC ĐẦU TIÊN CỦA PHIÊN TIẾP THEO
> [!IMPORTANT]
> **ƯU TIÊN HÀNG ĐẦU CỦA PHIÊN SAU:**
> 1. **Người đẹp kiểm tra và xác nhận các kết quả đối chiếu độc lập của 4 biểu vừa hoàn thành (`0.M5`, `0.M6`, `1.M1`, `2.M1`) trên các file Excel trong `outputs/` hoặc giao diện Streamlit.**
> 2. **Chưa chuyển sang biểu mới nếu chưa có ý kiến của Người đẹp về các điểm đối chiếu thực tế đã chốt.**
> 3. Sau khi Người đẹp duyệt xong kết quả đợt này, mới nhận chỉ đạo để xác định danh sách các biểu tiếp theo trong số 28 biểu còn lại (ví dụ nhóm Mầm non `0.M2`, `0.M3`, `0.M4` hoặc các biểu khác).

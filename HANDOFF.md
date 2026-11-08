# BIÊN BẢN CHỐT PHIÊN VÀ BÀN GIAO DỰ ÁN (HANDOFF.md)
## HỆ THỐNG TỔNG HỢP SỐ LIỆU PHỔ CẬP GIÁO DỤC & XÓA MÙ CHỮ NĂM 2026
- **Thời điểm chốt phiên:** 2026-10-29T19:35:00+07:00
- **Đơn vị tiếp nhận:** Sở Giáo dục và Đào tạo tỉnh Thái Nguyên
- **Người bàn giao:** Trợ lý AG (Antigravity)
- **Người nhận bàn giao:** Người đẹp
- **Quy mô hệ thống:** **Đầy đủ 92 xã/phường/thị trấn thực tế của tỉnh Thái Nguyên**
- **Trạng thái:** Toàn bộ 33 biểu nghiệp vụ, Sổ cái hợp nhất master và Báo cáo đối soát 63 quy tắc đã được đồng bộ 100% từ dữ liệu nguồn mới nhất.

---

### I. KẾT QUẢ KIỂM CHỨNG BỘ KHỞI ĐỘNG VÀ DỪNG HỆ THỐNG
*Tệp kiểm thử thực tế:* [`tests/test_startup_and_stop_scenarios.py`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/tests/test_startup_and_stop_scenarios.py) (Chạy và đạt 5/5 kịch bản):

1. **Khởi động khi app chưa chạy:**
   - PowerShell `scripts/check_port_service.ps1` kiểm tra trả về `STATUS=NOT_RUNNING`.
   - `KHOI_DONG_HE_THONG.bat` tự động phát hiện Python, kích hoạt Streamlit ngầm và mở trình duyệt tại `http://localhost:8501`.
2. **Nhấp đúp lần 2 khi app đang chạy (Chống trùng lặp tiến trình):**
   - Nhận diện tức thì `STATUS=PCGD_RUNNING;PID=...`.
   - Batch script nhảy thẳng tới nhãn `OPEN_BROWSER`, chỉ mở thêm tab trình duyệt, tuyệt đối không sinh ra thêm tiến trình chạy ngầm trùng lặp.
3. **Phát hiện xung đột cổng 8501:**
   - Giả lập một HTTP Server lạ chiếm cổng 8501 -> `check_port_service.ps1` phân tích và trả về chính xác:
     `STATUS=PORT_CONFLICT;PID=...;NAME=python.exe;CMD=...`
   - Batch script hiển thị cảnh báo xung đột cổng và từ chối chạy đè để bảo vệ an toàn cho ứng dụng kia.
4. **Nút dừng an toàn (`DUNG_HE_THONG.bat` / `scripts/stop_service.ps1`):**
   - Chỉ tìm và dừng đúng PID của tiến trình Streamlit PCGD thuộc dự án.
   - Nếu phát hiện cổng 8501 thuộc ứng dụng khác (`PORT_CONFLICT`), script từ chối dừng (exit code 2), tuyệt đối không dừng toàn bộ python.exe hoặc app khác.
5. **Thư mục có dấu tiếng Việt & khoảng trắng:**
   - Thử nghiệm trong thư mục `scratch/thử nghiệm thư mục tiếng Việt 2026/` -> PowerShell và Batch script thực thi mượt mà, không bị lỗi ký tự hay ngắt quãng lệnh.

---

### II. BỘ THIẾT LẬP KHI CHUYỂN SANG MÁY KHÁC (PORTABILITY)
*Không khẳng định chỉ sao chép `.venv` là chạy được:*
1. **Nguyên nhân kỹ thuật:** Thư mục `.venv` của Python gắn cứng đường dẫn tuyệt đối của máy tính hiện tại, không thể copy-paste máy khác.
2. **Bộ công cụ thiết lập máy mới đã tạo:**
   - [`CAI_DAT_BAN_DAU.bat`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/CAI_DAT_BAN_DAU.bat): Tệp chạy thiết lập 1-click cho máy mới.
   - [`scripts/setup_new_machine.py`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/scripts/setup_new_machine.py): Kịch bản kiểm tra môi trường, tạo `.venv` mới, cài thư viện qua pip.
   - [`requirements-lock.txt`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/requirements-lock.txt): Khóa chính xác 100% phiên bản thư viện đã kiểm chứng (Streamlit 1.66.1, openpyxl 3.1.5, pandas 3.1.0...).
3. **Xử lý ngoại lệ chặt chẽ (Đã kiểm thử qua `tests/test_setup_error_handling.py`):**
   - *Thiếu Python hoặc sai phiên bản (< 3.10 hoặc > 3.13):* Báo lỗi rõ ràng, hiển thị link tải Python 3.13 và ghi nhật ký `setup_error.log`.
   - *Mất mạng / mạng chặn PyPI:* Kiểm tra kết nối socket, thông báo hướng dẫn cắm mạng trực tiếp/4G và thoát có mã lỗi.
   - *Lỗi cài đặt thư viện:* Ghi toàn bộ STDOUT/STDERR vào `setup_error.log`, không bao giờ báo thành công giả tạo.
4. **Khởi động siêu tốc không cài lại:** Lệnh kiểm tra môi trường trong `KHOI_DONG_HE_THONG.bat` được tối ưu bằng `find_spec`, hoàn thành chỉ trong **0.059 giây**, không cài lại thư viện ở các lần mở tiếp theo.
5. **Đã kiểm thử thực tế trên Sandbox:** Tệp `scratch/test_setup_sandbox.py` đã tạo thử nghiệm `.venv` mới trong thư mục sandbox độc lập, cài đặt pip và import thành công, tự dọn dẹp sạch sẽ 100%.

---

### III. XÁC MINH SỰ THAY ĐỔI SỐ LIỆU 0.M1 & ĐỒNG BỘ ĐỘ MỚI
1. **Xác minh số liệu 0.M1:**
   - Báo cáo trước: $I5 = 23,765$; $K5 = 110,547$.
   - Báo cáo hiện tại: $I5 = 23,774$; $K5 = 110,542$.
   - *Nguyên nhân:* So sánh trực tiếp với Git HEAD phát hiện **12 xã trong thư mục `data/` được cán bộ cơ sở nộp cập nhật lại số liệu**:
     + Phường Bắc Kạn tăng 24 trẻ 5 tuổi ($I5$ từ $683 \rightarrow 707$).
     + Phường Vạn Xuân giảm 5 trẻ 5 tuổi ($I5$ từ $836 \rightarrow 831$).
     + Các điều chỉnh bổ sung tại Xã An Khánh, Côn Minh, Trung Thành...
     + Tổng $I5$ tăng đúng $+9$ trẻ (lên $23,774$); Tổng $K5$ giảm $5$ trẻ (xuống $110,542$).
   - Cơ chế Clean Recalculation đã tự động bắt đúng số liệu mới nhất của 92 xã.
2. **Bảo đảm hai file hợp nhất và đối soát cùng từ một bộ nguồn mới nhất:**
   - [`outputs/SO_CAI_TONG_HOP_PCGD_2026_THAI_NGUYEN.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/SO_CAI_TONG_HOP_PCGD_2026_THAI_NGUYEN.xlsx) (2,199.1 KB): Đóng gói từ 33 biểu mới nhất.
   - [`outputs/BAO_CAO_DOI_SOAT_63_QUY_TAC_PCGD_2026.xlsx`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/outputs/BAO_CAO_DOI_SOAT_63_QUY_TAC_PCGD_2026.xlsx) (32.1 KB): Được tạo trực tiếp từ Sổ cái hợp nhất và quét đối chiếu 92 xã mới nhất.
3. **Cơ chế giám sát độ mới dữ liệu (Data Freshness Badge):**
   - Tệp metadata: `outputs/last_sync_metadata.json` lưu thời điểm đồng bộ và hash/mtime của 92 files nguồn.
   - Đã tích hợp vào giao diện Web `ui/app.py`: Hiển thị Badge Xanh (`FRESH`) khi số liệu đồng bộ; hiển thị Cảnh báo Đỏ (`STALE`) kèm danh sách file thay đổi nếu có file nguồn mới hơn thời điểm tổng hợp.

---

### IV. BẰNG CHỨNG KIỂM THỬ ĐỘC LẬP TỪNG BIỂU (33/33 BIỂU ĐẠT 100%)
*Tuyệt đối không dùng kết quả 5 biểu để tuyên bố 33 biểu đạt:*
Hệ thống đã chạy trích xuất độc lập toàn bộ 33 biểu trên 92 xã thật (qua `scratch/collect_33_sheets_evidence.py`):
- **Khối Mầm non (7/7 biểu):** 0.M1 ($K5=110,542$; $I5=23,774$), 0.M2 (Lớp 3-4t: 1,845; Lớp 5t: 1,050), 0.M3 (Tổng CB+GV+NV: 11,784; TL GV/lớp: 2.93), 0.M4 (237 trường; TL phòng/lớp: 1.00), 0.M5 (Tổng chi NS: 12,304,024,943.24; TL chi CM: 2.13%), 0.M6 (296 trẻ KT; 226 can thiệp), 0.M7 (107,556 trẻ trong sổ; 49,129 nữ).
- **Khối Tiểu học (7/7 biểu):** 1.M1 (Trẻ 6t: 27,307; Tổng 6-10t: 140,711; Vào L1: 100%), 1.M2 (8,028 GV; TL GV/lớp: 1.62), 1.M3 (360 điểm trường; 4,984 lớp; TL Ph/lớp: 101.91%), 1.M4 (160 trường; Vào L1: 100%; HTTH: 99.99%), 1.M5 (3,797 trẻ KT; 3,304 tiếp cận GD), 1.M6 (10,390 HS đi học ngoài), 1.M7 (9,395 HS nơi khác đến học).
- **Khối THCS (10/10 biểu):** 2.M1 (TTN 15t: 27,620; TN THCS: 97.56%; Học tiếp: 98.63%), 2.M2 (Vào L1: 27,046; Có bằng THCS: 98,796), 2.M3 (TTN 11-18t: 230,682; TN THCS: 98,765), 2.M4 (3,332 lớp; TL phòng/lớp: 97.66%), 2.M5 (6,379 GV; TL GV/lớp: 1.91), 2.M6 (3,601 đối tượng KT; 2,869 tiếp cận GD), 2.M7 (4,067 HS đi học ngoài), 2.M8 (10,965 HS nơi khác đến học), 2.M9 (166,166 đối tượng 15-21t), 2.M10 (88,368 đối tượng 18-21t; 75,875 có bằng THPT/Nghề).
- **Khối XMC (9/9 biểu):** 3.M1 (Dân số: 244,515; Biết chữ: 244,352 - 99.93%), 3.M2 (Dân số: 479,435; Biết chữ: 478,664 - 99.84%), 3.M3 (Dân số: 1,097,844; Biết chữ: 1,090,373 - 99.32%), 3.M4 (Mù chữ M1 15-25t: 163), 3.M5 (Mù chữ M1 15-35t: 771), 3.M6 (Mù chữ M1 15-60t: 7,471), 3.M7 (Mù chữ M2 15-25t: 285), 3.M8 (Mù chữ M2 15-35t: 1,918), 3.M9 (Mù chữ M2 15-60t: 17,236).

---

### V. VẤN ĐỀ NGHIỆP VỤ CÒN TỒN TẠI VÀ ĐỀ XUẤT HƯỚNG XỬ LÝ
1. **Biểu 0.M5 (Tài chính GDMN):**
   - Đơn vị tính biểu mẫu ghi `Tr.đg`. Dữ liệu các xã nộp lên hỗn hợp cả đơn vị Đồng (hàng tỷ), Nghìn đồng và Triệu đồng.
   - Hệ thống giữ nguyên trạng số liệu nộp thô (Tổng chi NS 2026: $12,304,024,943.24$), không tự chia hệ số $1,000,000$ khi chưa có văn bản hướng dẫn từ Sở GDĐT.
   - Đề xuất Sở GDĐT ban hành văn bản quy định thống nhất đơn vị tính (Triệu đồng hay Đồng) để các trường/xã rà soát lại trước kỳ thanh tra chính thức.
2. **Quy trình tổng hợp lại khi có file mới:**
   - Cán bộ Sở chỉ cần chép file mới vào `data/`, mở trang Web và bấm nút "Tổng hợp lại", hệ thống sẽ kích hoạt Clean Recalculation và tự động chuyển trạng thái từ `STALE` sang `FRESH`.

---

### VI. BIÊN BẢN KHẮC PHỤC LỖI XUẤT EXCEL KHỐI 3.M (3.M1 - 3.M9)
- **Thời điểm:** 30/10/2026 07:50:00
- **Lỗi gốc:** `NameError: name 'export_3_mu_chu_test' is not defined` trên giao diện Web khi tổng hợp nhóm biểu mẫu xóa mù chữ.
- **Xác minh ánh xạ và tính toàn vẹn của 3.M2:**
  + Khẳng định **Biểu 3.M2** (Chuẩn XMC Mức 2 15-35 tuổi) có cấu trúc dữ liệu riêng (tỷ lệ biết chữ M1, M2; số xã đạt chuẩn M1, M2) và được điều hướng độc lập đến hàm xuất `export_3_m2_test`.
  + Nhóm biểu mù chữ **3.M4–3.M9** (Hiện trạng mù chữ Mức 1 và Mức 2 theo 3 độ tuổi) sử dụng chung bộ xử lý linh hoạt `processors/m3_mu_chu.py` và bộ xuất `export_3_mu_chu_test` (cùng các hàm wrapper `export_3_m4_test` đến `export_3_m9_test`).
- **Kết quả kiểm thử thực tế:**
  + Tệp kiểm thử tự động [`tests/test_block_3m_pipeline.py`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/tests/test_block_3m_pipeline.py) đã chạy qua toàn bộ luồng: Đọc 92 xã -> Tổng hợp cấp tỉnh -> Xuất Excel -> Mở đọc kiểm tra ngược 9/9 tệp kết quả.
  + **100% 9/9 biểu đạt chuẩn**, các tệp Excel xuất ra có đầy đủ 2 sheet (`TONG_HOP_THU`, `CHI_TIET_92_XA`), định dạng % chuẩn `0.00"%"`, số liệu toàn tỉnh và chi tiết 92 xã khớp chính xác tuyệt đối.
  + Dịch vụ Streamlit đã nạp mã mới sạch sẽ trên cổng 8501.


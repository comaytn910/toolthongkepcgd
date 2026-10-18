# QUY TẮC PHỐI HỢP VÀ VẬN HÀNH DỰ ÁN (AGENTS.md)
## HỆ THỐNG TỔNG HỢP SỐ LIỆU PHỔ CẬP GIÁO DỤC & XÓA MÙ CHỮ NĂM 2026
**Đơn vị sử dụng:** Sở Giáo dục và Đào tạo tỉnh Thái Nguyên  
**Môi trường:** Python 3.13 (Windows), Streamlit, openpyxl, pandas

---

### I. NGUYÊN TẮC BẮT BUỘC KHI KHỞI ĐỘNG PHIÊN MỚI
> [!IMPORTANT]
> **Mỗi phiên làm việc mới, AG BẮT BUỘC phải đọc tuần tự các tệp tài liệu sau trước khi thực hiện bất kỳ hành động nào:**
> 1. [`AGENTS.md`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/AGENTS.md): Bản quy ước ứng xử, xưng hô, thẩm quyền và các quy tắc cốt lõi đã chốt.
> 2. [`PROJECT_REQUIREMENTS.md`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/PROJECT_REQUIREMENTS.md): Đặc tả yêu cầu kỹ thuật và nghiệp vụ chi tiết của 33 biểu mẫu.
> 3. [`PROGRESS.md`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/PROGRESS.md): Nhật ký tiến độ thực tế, trạng thái từng phiếu và bằng chứng kiểm thử.
> 4. [`HANDOFF.md`](file:///d:/TOOL%20LAP%20TRINH/TONG%20HOP%20PHO%20CAP/HANDOFF.md): Biên bản bàn giao phiên gần nhất, các lưu ý kỹ thuật và công việc đầu tiên cần làm tiếp theo.

---

### II. XƯNG HÔ VÀ DANH XƯNG QUY CHUẨN
- **Người dùng:** Gọi là **Người đẹp**.
- **Trợ lý AI (Antigravity):** Gọi là **AG**.
- **ChatGPT / Mô hình phụ trợ đối chiếu:** Gọi là **GPT**.

---

### III. PHẠM VI VÀ THẨM QUYỀN TRIỂN KHAI
1. **Quy mô dự án:**
   - Tổng hợp 33 biểu nghiệp vụ PCGD - XMC và 4 sheet hỗ trợ từ mẫu chuẩn `MẪU CHUẨN - PCGD 2026(ok).xlsx`.
   - Dự kiến tiếp nhận đủ 92 file Excel từ 92 xã/phường/thị trấn tỉnh Thái Nguyên.
   - Hiện tại kiểm thử bằng các file thật trong thư mục `data/` (ví dụ: `01. Xã Yên Trạch...`, `02. Xã Yên Thịnh...`).
2. **Quyền hạn tự thực hiện đã duyệt trước:**
   - **Người đẹp** đã phê duyệt toàn diện cho **AG** tự khảo sát, lập đặc tả, viết/sửa mã, bổ sung giao diện, chạy ứng dụng, tự động kiểm thử và xuất kết quả.
   - Triển khai tuần tự từng biểu: Phiếu đạt 100% kiểm thử mới tự động chuyển phiếu tiếp theo; không dừng chờ duyệt từng bước thường lệ.
   - Khi hoàn thành toàn bộ đợt, báo cáo tổng hợp một lần kèm đầy đủ bằng chứng đối chiếu độc lập.

---

### IV. CÁC NGUYÊN TẮC KỸ THUẬT & NGHIỆP VỤ BẤT DI BẤT DỊCH
1. **Bảo toàn dữ liệu nguồn:**
   - Tuyệt đối **không sửa đổi, không ghi đè, không di chuyển** các file Excel trong `data/` và file `MẪU CHUẨN - PCGD 2026(ok).xlsx`.
   - Kết quả xuất ra thư mục `outputs/`, ghi rõ *"Tổng hợp thử"* và số lượng xã thực tế được tính.
2. **Tính mới hoàn toàn khi chạy lại (Clean Recalculation):**
   - Khi thay đổi, thêm, bớt file trong `data/` và bấm tổng hợp lại, hệ thống đọc lại toàn bộ và tính mới 100%. Không lưu cache cộng dồn với kết quả cũ.
3. **Chống trùng lặp xã:**
   - Mỗi xã chỉ được tính 1 file duy nhất; loại trừ tự động các file trùng xã.
4. **Nguyên tắc tính tỷ lệ (%) và mức chi bình quân:**
   - Tuyệt đối không cộng trực tiếp các tỷ lệ hoặc lấy trung bình cộng tỷ lệ nguồn của các xã.
   - Tỷ lệ cấp tỉnh phải tính từ: $\frac{\sum \text{Tử số}}{\sum \text{Mẫu số}} \times 100$ trên cùng tập xã hợp lệ có cả tử và mẫu.
   - Mức chi bình quân phải dùng đúng trọng số (số học sinh, số đối tượng thực tế).
5. **Quy chuẩn lưu trữ và hiển thị phần trăm (Yêu cầu cốt lõi):**
   - Giữ tỷ lệ dưới dạng **số thực theo thang 0–100** (ví dụ: `3.45` nghĩa là $3.45\%$; `100` nghĩa là $100.00\%$).
   - Định dạng hiển thị ô Excel: `cell.number_format = '0.00"%"'`.
   - Tuyệt đối không dùng `0%` hoặc `0.00%` với số đã nhân 100 (sẽ bị Excel nhân tiếp thành 345% hoặc 10000%).
   - Tuyệt đối không áp dụng định dạng `%` cho số tiền hoặc số lượng.
6. **Minh bạch kiểu dữ liệu & Không che giấu thiếu hụt:**
   - Phân biệt rõ số 0 thực tế, ô trống (`None`/`""`), ký hiệu `"x"` (không áp dụng), lỗi sai kiểu (`ERROR_TYPE`), lỗi công thức Excel và thiếu kết quả lưu sẵn.
   - Khi thiếu dữ liệu hoặc có lỗi, hiển thị rõ số xã đóng góp hợp lệ (ví dụ: *Đã cộng 1/2 xã*), không biến thiếu dữ liệu thành 0 hay 0%.
7. **Bảo toàn và hồi quy 0.M1:**
   - Biểu `0.M1` đã được nghiệm thu hoàn chỉnh. Khi sửa bất kỳ thành phần dùng chung nào (`excel_reader.py`, `excel_exporter.py`, `app.py`), bắt buộc phải chạy lại kiểm thử hồi quy `0.M1` để đảm bảo không bị ảnh hưởng.
8. **Chống xung đột khóa file Windows:**
   - Khi xuất file Excel, luôn bao bọc trong `try ... except PermissionError` để lưu dự phòng ra file `_capnhat.xlsx` nếu người dùng đang mở file xem.

---

### V. ĐẶC TẢ NGHIỆP VỤ RIÊNG CỦA 5 BIỂU ĐÃ HOÀN THÀNH
- **0.M1 (Trẻ em MN theo độ tuổi):**
  - Tính lại Cột K từ các cột tuổi ($D:I$ cho hàng 5–25, $H:J$ cho hàng 26, 27, 29). Đối chiếu K nguồn và ghi sai lệch.
  - Tỷ lệ Hàng 16, 22, 28 tính lại từ tử/mẫu cấp tỉnh, định dạng `0.00"%"'`.
- **0.M5 (Tình hình tài chính GDMN):**
  - Cộng đúng số tiền theo chỉ tiêu, năm ($N-1$, $N$) và nguồn (Ngân sách, Học phí...).
  - Không cộng dồn tỷ lệ Hàng 13; tính lại từ tổng chi chuyên môn / tổng chi ngân sách.
  - Lọc sạch ô D36 (Xã Yên Thịnh) chứa giá trị ngày tháng `04/05/2026`, không bị cộng nhầm serial 46146.0.
  - Định dạng tiền tệ `#,##0.00`, tỷ lệ `0.00"%"'`.
- **0.M6 (Đối tượng khuyết tật MN):**
  - Tổng hợp 6 năm sinh và 8 dạng khuyết tật.
  - Cột L (can thiệp/học tập) cộng độc lập; Cột M (tỷ lệ %) tính lại $= L / C \times 100$.
  - Đối soát chéo khớp 100% với $I8$ (khuyết tật) và $I10$ (tiếp cận GD) của biểu `0.M1`.
- **1.M1 (Trẻ 6–14t PCGD & GD bắt buộc Tiểu học):**
  - Tách bạch nhóm 6–10 tuổi ($F:J$) và 11–14 tuổi ($L:O$). Tính lại $K = \sum F:J$ và $P = \sum L:O$ cấp tỉnh, loại bỏ sai lệch nguồn và tránh cộng trùng.
  - Hàng 14 (Số PPC) tính lại $= \text{Hàng 8} - \text{Hàng 11}$.
  - Bỏ qua các cột công thức cảnh báo Q..AA xen trong bảng.
  - Bảng tiêu chí Phần 2 (Hàng 58–63) tính chuẩn tử/mẫu trên cùng tập xã, tỷ lệ định dạng `0.00"%"'`.
- **2.M1 (Thanh thiếu niên PCGD THCS):**
  - Phân biệt độ tuổi 11–18 và 20 cột kênh học tập: Phổ thông (PT), Thường xuyên (TX), Giáo dục nghề nghiệp (GDNN).
  - Cột Z cấp tỉnh tính lại: Hàng dân số 7–12 tính $= N + Q + T + W$; Hàng chi tiết 15–49 tính $= \sum N:Y$ (12 cột tuổi 15–18).
  - **Chuẩn hóa riêng H59 theo thang 0–100:** Khắc phục lỗi công thức mẫu chỉ chia `G59/W13` mà quên nhân 100; tính lại $\text{round}(G59/W13 \times 100, 2)$, định dạng `0.00"%"'`.
  - Bảng tiêu chí (Hàng 55–59) tính chuẩn từ chi tiết, phát hiện và ghi nhận chênh lệch của xã Yên Thịnh nhập đè số tĩnh $G58=645$ (tính chuẩn là $568$).

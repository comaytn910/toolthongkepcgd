# TÀI LIỆU YÊU CẦU DỰ ÁN (PROJECT REQUIREMENTS)
## HỆ THỐNG TỔNG HỢP SỐ LIỆU PHỔ CẬP GIÁO DỤC & XÓA MÙ CHỮ NĂM 2026
**Đơn vị sử dụng:** Sở Giáo dục và Đào tạo tỉnh Thái Nguyên  
**Quy mô triển khai:** 92 xã/phường/thị trấn toàn tỉnh Thái Nguyên  
**Môi trường vận hành:** Máy tính Windows, Python 3.13, Giao diện Web cục bộ (Streamlit)  

---

### I. MỤC TIÊU VÀ NGUYÊN TẮC CỐT LÕI

1. **Phát triển theo giai đoạn và từng biểu:**
   - Triển khai theo từng giai đoạn rõ ràng: Khảo sát -> Xây dựng khung nền tảng -> Triển khai & nghiệm thu từng biểu nghiệp vụ (bắt đầu với biểu `0.M1`).
   - Các biểu chưa xây dựng phải hiển thị trạng thái **"Chưa triển khai"** trên giao diện; tuyệt đối không làm gộp cả 33 biểu trong một lượt khi chưa kiểm thử.

2. **Nguyên tắc bảo toàn dữ liệu nguồn:**
   - Không được sửa đổi, ghi đè hoặc làm biến dạng các file Excel nguồn nộp từ xã/phường cũng như file mẫu chuẩn.
   - Toàn bộ kết quả tổng hợp phải truy xuất được nguồn gốc (Audit Trail): Chỉ tiêu nào, từ xã nào, file nào, sheet nào và tọa độ ô nào.

3. **Nguyên tắc tính toán số học & tỷ lệ chuẩn nghiệp vụ:**
   - **Chỉ tiêu số lượng (đếm/người/lớp/phòng):** Được cộng dồn từ các xã nộp hợp lệ.
   - **Chỉ tiêu tỷ lệ (%):** **TUYỆT ĐỐI KHÔNG ĐƯỢC CỘNG HOẶC LẤY TRUNG BÌNH CỘNG ĐƠN GIẢN** các tỷ lệ của các xã. Tỷ lệ cấp tỉnh phải được tính lại dựa trên:
     $$\text{Tỷ lệ toàn tỉnh} = \frac{\sum \text{Tử số các xã}}{\sum \text{Mẫu số các xã}} \times 100$$
   - Phân biệt minh bạch giữa:
     - Giá trị `0` (số 0 thực tế)
     - Ô trống `None`/`""` (chưa nhập liệu)
     - Ký hiệu `'x'` (không áp dụng)
     - Dữ liệu sai kiểu (chuỗi văn bản nằm trong ô số)
     - Ô công thức lỗi (`#VALUE!`, `#REF!`, `#DIV/0!`, ...) hoặc công thức chưa chạy kết quả.
     - **Không tự ý chuyển dữ liệu thiếu thành số 0.**

4. **Nguyên tắc kiểm soát đơn vị và phát hiện trùng lặp:**
   - Mỗi xã/phường chỉ được chọn duy nhất **01 file** để đưa vào tổng hợp trong mỗi đợt.
   - Nhận diện đơn vị dựa trên danh mục chuẩn 92 xã/phường của tỉnh Thái Nguyên kết hợp thông tin trong file (`TỔNG_QUAN!A2`, `ĐỐI_SOÁT!A2`, tên file).
   - Phát hiện và cảnh báo: Xã bị nộp nhiều file (trùng lặp), file trùng tên, file không đúng mẫu, file thiếu biểu nghiệp vụ.
   - Không chỉ dựa vào số lượng file (ví dụ có 92 file) để kết luận đủ 92 xã (có thể có 2 file cùng 1 xã và thiếu xã khác).

5. **Phân biệt mẫu xã và mẫu tổng hợp tỉnh:**
   - File mẫu hiện tại là biểu nhập liệu cơ sở dành cho cấp xã/phường.
   - Mẫu kết quả tổng hợp cấp tỉnh phải được thống nhất về bố cục (bảng tổng hợp toàn tỉnh hoặc bảng chi tiết 92 xã + dòng tổng toàn tỉnh) trước khi lập trình xuất kết quả.

---

### II. KIỂM KÊ VÀ PHÂN LOẠI DANH MỤC BIỂU

Bộ file chuẩn `MẪU CHUẨN - PCGD 2026(ok).xlsx` gồm **37 sheet**:
- **33 biểu nghiệp vụ chuyên môn:**
  - **Khối Mầm non (0.M): 7 biểu** (`0.M1` đến `0.M7`)
  - **Khối Tiểu học (1.M): 7 biểu** (`1.M1` đến `1.M7`)
  - **Khối THCS (2.M): 10 biểu** (`2.M1` đến `2.M10`)
  - **Khối Xóa mù chữ (3.M): 9 biểu** (`3.M1` đến `3.M9`)
- **04 sheet hỗ trợ & điều hành hệ thống:**
  - `ĐỐI_SOÁT`: Báo cáo đối soát 63 quy tắc kiểm tra logic dữ liệu.
  - `TỔNG_QUAN`: Dashboard kiểm tra tiến độ, số biểu hoàn thành, đang nhập, có lỗi.
  - `CAU_HINH_DOI_SOAT`: Bảng cấu hình mã quy tắc và công thức đối soát chéo giữa các sheet.
  - `DS_SAO_LUU_20261012_180009_1147`: Bản chụp sao lưu cấu hình quy tắc.

*Lưu ý hệ thống:* Không được coi file là lỗi/thừa chỉ vì có 37 sheet; 4 sheet này là thành phần hỗ trợ nghiệp vụ đã được Sở GDĐT thiết kế sẵn.

---

### III. BẢNG PHÂN LOẠI CẤU TRÚC 33 BIỂU NGHIỆP VỤ

| Nhóm biểu | Mã biểu | Tên biểu | Phạm vi dữ liệu thực | Dạng bảng | Tính chất hàng |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Mầm non** | `0.M1` | Thống kê trẻ em MN theo độ tuổi | A1:K40 | Ma trận độ tuổi & Tiêu chí | Cố định (40 hàng) |
| | `0.M2` | Tiêu chuẩn, ĐK PCGD MN | A1:V9 | Ma trận chỉ tiêu xã | Cố định (9 hàng) |
| | `0.M3` | Đội ngũ CBQL, GV, NV | A1:S73 | Danh sách cơ sở / nhóm trẻ | **Có thể thêm hàng** |
| | `0.M4` | CSVC Mầm non | A1:X14 | Ma trận phòng học, thiết bị | Cố định (14 hàng) |
| | `0.M5` | Tình hình tài chính MN | A1:J38 | Ma trận nguồn kinh phí | Cố định (38 hàng) |
| | `0.M6` | Trẻ khuyết tật MN | A2:M13 | Ma trận dạng khuyết tật | Cố định (13 hàng) |
| | `0.M7` | Sổ theo dõi PCGD MN | A3:P50 | Danh sách chi tiết từng trẻ | **Có thể thêm hàng** |
| **Tiểu học** | `1.M1` | Trẻ 6-14t PCGD & GD bắt buộc | A2:AA63 | Ma trận độ tuổi & lớp học | Cố định (63 hàng) |
| | `1.M2` | Đội ngũ giáo viên Tiểu học | A2:AG35 | Ma trận trình độ giáo viên | Cố định (35 hàng) |
| | `1.M3` | CSVC cấp Tiểu học | A1:AC101 | Bảng theo trường / điểm trường | **Có thể thêm hàng** |
| | `1.M4` | Kết quả PCGD Tiểu học | A1:U116 | Ma trận kết quả các chuẩn | Cố định (116 hàng) |
| | `1.M5` | Trẻ khuyết tật Tiểu học | A3:N17 | Ma trận dạng khuyết tật | Cố định (17 hàng) |
| | `1.M6` | Học sinh học xã/tỉnh ngoài | A2:AR49 | Danh sách trường ngoài | **Có thể thêm hàng** |
| | `1.M7` | Học sinh ngoài học tại xã | A2:AR68 | Danh sách trường ngoài | **Có thể thêm hàng** |
| **THCS** | `2.M1` | Thống kê TTN PCGD THCS | A1:Z59 | Ma trận thanh thiếu niên | Cố định (59 hàng) |
| | `2.M2` | Tổng hợp huy động PCGD THCS | A1:W48 | Ma trận huy động | Cố định (48 hàng) |
| | `2.M3` | Thống kê PCGD THCS | A1:S46 | Ma trận chuẩn PCGD | Cố định (46 hàng) |
| | `2.M4` | CSVC cấp THCS | A1:X12 | Ma trận CSVC THCS | Cố định (12 hàng) |
| | `2.M5` | Đội ngũ giáo viên THCS | A1:AQ17 | Danh sách trường THCS | **Có thể thêm hàng** |
| | `2.M6` | Trẻ khuyết tật THCS | A2:M14 | Ma trận dạng khuyết tật | Cố định (14 hàng) |
| | `2.M7` | HS THCS học ngoài địa bàn | A1:AM40 | Danh sách trường ngoài | **Có thể thêm hàng** |
| | `2.M8` | HS THCS ngoài học tại xã | A2:AM39 | Danh sách trường ngoài | **Có thể thêm hàng** |
| | `2.M9` | Kết quả huy động GD Trung học | A2:V48 | Ma trận huy động THPT/GDTX | Cố định (48 hàng) |
| | `2.M10` | Hiệu quả GD Trung học | A2:T47 | Ma trận phân luồng, tốt nghiệp | Cố định (47 hàng) |
| **Xóa mù chữ** | `3.M1` | Chuẩn XMC Mức 1 (15-25t, 15-35t, 15-60t) | A2:G48 | Ma trận độ tuổi XMC Mức 1 | Cố định (48 hàng) |
| | `3.M2` | Chuẩn XMC Mức 2 (15-25t, 15-35t, 15-60t) | A2:G49 | Ma trận độ tuổi XMC Mức 2 | Cố định (49 hàng) |
| | `3.M3` | Chuẩn XMC Chung toàn xã | A2:G49 | Ma trận tổng hợp XMC | Cố định (49 hàng) |
| | `3.M4` | Hiện trạng mù chữ Mức 1 (15-25t) | A2:O48 | Ma trận chi tiết xóm/độ tuổi | Cố định (48 hàng) |
| | `3.M5` | Hiện trạng mù chữ Mức 1 (15-35t) | A2:O47 | Ma trận chi tiết xóm/độ tuổi | Cố định (47 hàng) |
| | `3.M6` | Hiện trạng mù chữ Mức 1 (15-60t) | A2:O47 | Ma trận chi tiết xóm/độ tuổi | Cố định (47 hàng) |
| | `3.M7` | Hiện trạng mù chữ Mức 2 (15-25t) | A3:N48 | Ma trận chi tiết xóm/độ tuổi | Cố định (48 hàng) |
| | `3.M8` | Hiện trạng mù chữ Mức 2 (15-35t) | A2:N47 | Ma trận chi tiết xóm/độ tuổi | Cố định (47 hàng) |
| | `3.M9` | Hiện trạng mù chữ Mức 2 (15-60t) | A2:N47 | Ma trận chi tiết xóm/độ tuổi | Cố định (47 hàng) |


---

### IV. ĐIỀU CHỈNH PHẠM VI KIỂM THỬ THỰC TẾ HIỆN TẠI

1. **Cơ chế nạp file tự động từ thư mục:**
   - Người dùng đặt 1–2 file Excel thật của xã/phường vào thư mục `data/`.
   - Hệ thống quét và đọc trực tiếp từ thư mục `data/`, không yêu cầu người dùng phải tải lên (upload) thủ công từng file.
   - Thiết kế mô-đun đọc thư mục theo chuẩn tổng quát để sau này khi có đủ 92 file thì tự động quét và xử lý toàn bộ 92 file mà không cần sửa code.

2. **Nguyên tắc dữ liệu thực:**
   - Không yêu cầu bắt buộc đủ 92 file ở giai đoạn này.
   - **Tuyệt đối không tự nhân bản file hoặc tạo dữ liệu giả mạo** để lấp đầy 92 xã. Chỉ tính toán trên đúng các file thật hiện có trong `data/`.

3. **Quy trình 2 bước rõ ràng trên ứng dụng:**
   - **Bước 1: Kiểm tra khả năng đọc file (File Reader & Inspector):**
     Quét thư mục `data/`, hiển thị danh sách: Tên file, Tên xã/phường nhận diện được, Số lượng biểu tìm thấy, Trạng thái (Hợp lệ / Lỗi) và chi tiết lỗi nếu có (ví dụ: thiếu sheet 0.M1, file sai định dạng, xã bị trùng lặp).
   - **Bước 2: Tổng hợp và Xem trước biểu 0.M1:**
     Chỉ tổng hợp biểu `0.M1` khi người dùng bấm nút tổng hợp. Các biểu khác hiển thị rõ nhãn **"Chưa triển khai"**.

4. **Xuất kết quả thử nghiệm:**
   - Kết quả xuất ra thư mục `outputs/`, giữ nguyên vẹn 100% các file gốc trong thư mục `data/`.
   - Tên file xuất và tiêu đề bảng ghi rõ: **"Tổng hợp thử - Biểu 0.M1 (Số xã thực tế được tính: N)"**.

5. **Nguyên tắc "Tổng hợp lại" (Clean Recalculation):**
   - Khi người dùng thay thế, thêm, bớt file trong thư mục `data/` rồi bấm "Tổng hợp lại", hệ thống phải giải phóng bộ nhớ, đọc lại từ đầu các file hiện tại và tính toán mới hoàn toàn. **Không được lưu cache cộng dồn với kết quả lần chạy trước.**

---

### V. PHÊ DUYỆT TRIỂN KHAI TỰ ĐỘNG & THỨ TỰ ƯU TIÊN TIẾP THEO

Người dùng đã duyệt trước toàn bộ các bước triển khai tự động khép kín theo thứ tự:
$$\text{0.M1 (Đã đạt)} \longrightarrow \text{0.M5} \longrightarrow \text{0.M6} \longrightarrow \text{1.M1} \longrightarrow \text{2.M1}$$

1. **Quyền hạn tự thực hiện:**
   - AG tự khảo sát, lập đặc tả, viết/sửa mã, bổ sung giao diện, chạy ứng dụng, tự động kiểm thử và xuất kết quả.
   - Triển khai tuần tự từng phiếu: Phiếu đạt 100% kiểm thử mới chuyển sang phiếu tiếp theo. Không dừng chờ duyệt giữa chừng.

2. **Quy tắc riêng từng phiếu:**
   - **0.M5:** Cộng đúng số tiền theo chỉ tiêu, năm (N-1, N) và nguồn kinh phí. Không cộng trực tiếp tỷ lệ; tính lại tỷ lệ Hàng 13 = Tử số / Mẫu số. Lọc sạch ô D36 chứa `datetime` để không biến thành số 46146.0 làm sai số tiền. Định dạng tiền tệ `#,##0.00`, tỷ lệ `0.00"%"`.
   - **0.M6:** Tổng hợp ma trận 6 năm sinh và 8 dạng khuyết tật. Tính lại cột L (can thiệp) và cột M (tỷ lệ %) = L / C * 100. Đối soát chéo với I8 và I10 của Biểu 0.M1.
   - **1.M1:** Phân biệt rõ nhóm 6–10 tuổi (F:J, Tổng K) và 11–14 tuổi (L:O, Tổng P). Tính lại cấp tỉnh K = SUM(F:J) và P = SUM(L:O) để không cộng trùng và loại bỏ sai lệch nguồn. Bỏ qua các cột công thức cảnh báo Q..AA. Tính chuẩn xác Bảng tiêu chí Phần 2 (Hàng 58-63) với tỷ lệ theo thang 0–100 và định dạng `0.00"%"`.
   - **2.M1:** Phân biệt tuổi (11-18) và 20 cột kênh học tập: PT, TX, GDNN. Tính lại Cột Z cấp tỉnh từ chi tiết các tuổi 15-18. **Chuẩn hóa riêng H59 theo thang 0–100:** Khắc phục lỗi công thức mẫu chỉ chia `G59/W13` mà quên nhân 100; tính lại `round(G59/W13 * 100, 2)` và định dạng `0.00"%"`. Bảng tiêu chí tính chuẩn từ chi tiết, phát hiện chênh lệch do xã nhập đè số tĩnh.

3. **Quy chuẩn hiển thị tỷ lệ phần trăm:**
   - Giữ nguyên số thực theo thang 0–100 (ví dụ: 3.45 là 3.45%, 100 là 100.00%).
   - Sử dụng định dạng Excel: `0.00"%"`.
   - Tuyệt đối không dùng `0%` hay `0.00%` với số đã nhân 100 (tránh bị Excel nhân 100 thành 345% hay 10000%).
   - Tuyệt đối không áp dụng định dạng phần trăm cho số lượng hoặc số tiền.



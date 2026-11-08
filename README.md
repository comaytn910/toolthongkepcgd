# HỆ THỐNG TỔNG HỢP SỐ LIỆU PHỔ CẬP GIÁO DỤC & XÓA MÙ CHỮ (PCGD - XMC) NĂM 2026
### SỞ GIÁO DỤC VÀ ĐÀO TẠO TỈNH THÁI NGUYÊN

![Python Version](https://img.shields.io/badge/Python-3.10%20--%203.13-blue.svg)
![Framework](https://img.shields.io/badge/UI-Streamlit%201.66-red.svg)
![Status](https://img.shields.io/badge/Trạng%20Thái-Hoàn%20Thành%20100%25%20(33%2F33%20Biểu)-success.svg)
![Scope](https://img.shields.io/badge/Quy%20Mô-92%20Xã%2FPhường%20Thái%20Nguyên-green.svg)

---

## 📖 GIỚI THIỆU TỔNG QUAN

Hệ thống được xây dựng chuyên dụng để tự động hóa công tác thu thập, làm sạch, tổng hợp và đối soát chéo số liệu Phổ cập Giáo dục & Xóa mù chữ năm 2026 từ **toàn bộ 92 xã, phường, thị trấn** thuộc tỉnh Thái Nguyên theo biểu mẫu chuẩn của Bộ Giáo dục và Đào tạo (`MẪU CHUẨN - PCGD 2026(ok).xlsx`).

### ✨ Tính Năng Nổi Bật:
1. **Tổng hợp tự động 33/33 biểu nghiệp vụ** thuộc 4 cấp học: Mầm non (7 biểu), Tiểu học (7 biểu), THCS (10 biểu), Xóa mù chữ (9 biểu).
2. **Cơ chế tính mới 100% (Clean Recalculation):** Không cộng dồn cache số liệu cũ. Bất cứ khi nào cán bộ nộp lại file mới vào thư mục `data/`, hệ thống đọc lại từ đầu và tính toán mới hoàn toàn.
3. **Giám sát độ mới của dữ liệu (Data Freshness Badge):**
   - 🟢 **Xanh (DỮ LIỆU ĐỒNG BỘ):** Kết quả khớp 100% với toàn bộ file nguồn trong `data/`.
   - 🔴 **Đỏ (CẢNH BÁO SỐ LIỆU ĐANG CŨ):** Nhận diện ngay xã nào vừa cập nhật lại file để nhắc nhở người dùng bấm "Tổng hợp lại".
4. **Đối soát thanh tra 63 quy tắc:** Tự động quét kiểm tra tính hợp lý nội tại giữa các biểu (ví dụ: đối soát trẻ khuyết tật giữa 0.M1 và 0.M6, đối soát thanh thiếu niên 2.M1 với 2.M2, v.v.).
5. **Đóng gói xuất bản đa dạng:** Xuất riêng từng biểu kèm chi tiết 92 xã; xuất Sổ cái toàn diện 33 biểu cấp tỉnh; xuất Báo cáo phục vụ thanh kiểm tra của Sở GD&ĐT và Bộ GD&ĐT.

---

## 🚀 HƯỚNG DẪN CÀI ĐẶT TRÊN MÁY TÍNH MỚI (CHUYỂN MÁY)

Khi chuyển sang máy tính mới để làm việc, Người đẹp không cần phải cấu hình thủ công phức tạp. Hệ thống đã tích hợp sẵn bộ cài đặt tự động 1-Click:

### 1. Điều kiện tiên quyết:
- Máy tính chạy hệ điều hành **Windows 10 hoặc Windows 11**.
- Đã cài đặt **Python** (phiên bản từ **3.10** đến **3.13**, khuyến nghị Python 3.13 chính thức).
  > ⚠️ **LƯU Ý QUAN TRỌNG KHI CÀI PYTHON:** Ở màn hình cài đặt đầu tiên của Python, **BẮT BUỘC đánh dấu chọn ô**: `☑ Add python.exe to PATH`.
- Máy tính có kết nối Internet trong lần đầu thiết lập để tải thư viện.

### 2. Các bước cài đặt:

**Bước 1: Tải mã nguồn về máy**
Mở Command Prompt hoặc PowerShell và chạy lệnh:
```bash
git clone https://github.com/comaytn910/toolthongkepcgd.git
cd toolthongkepcgd
```
*(Hoặc tải file ZIP từ GitHub về và giải nén vào ổ đĩa bất kỳ, ví dụ `D:\TOOL LAP TRINH\TONG HOP PHO CAP`)*.

**Bước 2: Chạy thiết lập tự động 1-Click**
- Vào thư mục dự án vừa tải về.
- Nhấp đúp chuột vào tệp:
  👉 **`CAI_DAT_BAN_DAU.bat`**

Hệ thống sẽ hoàn toàn tự động:
1. Kiểm tra tính tương thích của phiên bản Python trên máy.
2. Tự động khởi tạo môi trường ảo độc lập (`.venv`).
3. Tải và cài đặt chính xác các phiên bản thư viện đã khóa từ `requirements-lock.txt` (Streamlit 1.66, openpyxl, pandas, requests...).
4. Kiểm tra thử nghiệm import các mô-đun lõi.
5. Tự động tạo biểu tượng lối tắt **"HE THONG PCGD 2026 - THAI NGUYEN"** ngoài màn hình Desktop.

---

## 🖥️ HƯỚNG DẪN SỬ DỤNG VÀ VẬN HÀNH HÀNG NGÀY

### 1. Khởi Động Ứng Dụng:
- **Cách 1 (Nhanh nhất):** Nhấp đúp chuột vào biểu tượng **"HE THONG PCGD 2026 - THAI NGUYEN"** ngoài Desktop.
- **Cách 2:** Nhấp đúp vào tệp **`KHOI_DONG_HE_THONG.bat`** trong thư mục dự án.

> ⚡ **Tốc độ mở siêu tốc:** Chỉ mất chưa đầy 1–2 giây vì hệ thống kiểm tra môi trường bằng `find_spec` (0.05s), không cài đặt lại thư viện ở các lần mở sau.  
> 🛡️ **Cơ chế an toàn:** Nếu ứng dụng đang chạy mà nhấp đúp lần nữa, hệ thống tự động nhận diện và chỉ mở thêm tab trình duyệt tại `http://localhost:8501`, tuyệt đối không sinh tiến trình chạy ngầm trùng lặp.

### 2. Quy Trình Tổng Hợp Khi Có File Nguồn Mới:
1. Khi có xã nộp file hoặc sửa lại số liệu, chỉ cần chép đè hoặc thêm file Excel đó vào thư mục:
   📂 `data/`
2. Mở giao diện Web trên trình duyệt (`http://localhost:8501`).
3. Thanh trạng thái ở đầu trang sẽ báo màu đỏ: `🔴 CẢNH BÁO: DỮ LIỆU ĐANG CŨ` kèm danh sách xã vừa sửa đổi.
4. Bấm nút **"🔄 Tổng hợp lại dữ liệu"**.
5. Hệ thống sẽ đọc lại 100% toàn bộ các xã, tính toán lại và chuyển sang màu xanh: `🟢 DỮ LIỆU ĐỒNG BỘ`.

### 3. Tải Tài Liệu & Kết Quả Xuất Ra:
Ngay đầu trang Web có 2 nút tải báo cáo cấp tỉnh:
- 📥 **TẢI SỔ CÁI 33 BIỂU CẤP TỈNH:** Tệp `outputs/SO_CAI_TONG_HOP_PCGD_2026_THAI_NGUYEN.xlsx` (~2.2 MB) chứa đầy đủ 33 biểu nghiệp vụ cấp tỉnh.
- 📥 **TẢI BÁO CÁO ĐỐI SOÁT 63 QUY TẮC:** Tệp `outputs/BAO_CAO_DOI_SOAT_63_QUY_TAC_PCGD_2026.xlsx` phục vụ thanh tra logic chéo giữa các biểu của 92 xã.
- Ngoài ra, trong thư mục `outputs/` luôn có sẵn 33 file Excel riêng biệt tương ứng với từng biểu (mỗi file gồm sheet Toàn tỉnh và sheet Chi tiết 92 xã).

### 4. Tắt Dịch Vụ An Toàn:
- Nhấp đúp vào tệp: **`DUNG_HE_THONG.bat`** (hoặc chọn mục số `[5]` trong bảng điều khiển).
- Hệ thống chỉ tắt duy nhất tiến trình Streamlit của ứng dụng PCGD, bảo đảm an toàn tuyệt đối cho các chương trình Python khác đang chạy trên máy.

---

## 📤 HƯỚNG DẪN ĐẨY CẬP NHẬT LÊN GITHUB

Khi Người đẹp làm việc xong trên máy này và muốn đẩy phiên làm việc lên GitHub để máy khác có thể tải về:

1. Nhấp đúp vào tệp:
   👉 **`DAY_LEN_GITHUB.bat`**
2. Cửa sổ dòng lệnh sẽ tự động kiểm tra và thực hiện đẩy lên nhánh `main` của kho lưu trữ.
3. Nếu là lần đầu tiên chạy từ máy mới, trình duyệt sẽ hiện lên yêu cầu xác nhận, Người đẹp chỉ cần bấm **"Sign in with your browser"** (hoặc Approve). Toàn bộ dữ liệu sẽ được đẩy lên GitHub trong 30 giây đến 1 phút.

---

## 📂 DANH MỤC 33 BIỂU NGHIỆP VỤ ĐÃ TRIỂN KHAI

| Khối Cấp Học | Mã Biểu | Tên Biểu Nghiệp Vụ Chuyên Môn | Chỉ Tiêu Toàn Tỉnh (92 Xã Thật) | Trạng Thái |
| :--- | :---: | :--- | :--- | :---: |
| **Mầm non (0.M)** | `0.M1` | Thống kê trẻ em MN theo độ tuổi | Tổng trẻ: 110,542; Trẻ 5t: 23,774 | ✅ Đạt 100% |
| | `0.M2` | Quy mô nhóm lớp GDMN | Lớp MG 3-4t: 1,845; Lớp MG 5t: 1,050 | ✅ Đạt 100% |
| | `0.M3` | Đội ngũ CBQL, GV, NV Mầm non | Tổng CB+GV+NV: 11,784; Tỷ lệ GV/lớp: 2.93 | ✅ Đạt 100% |
| | `0.M4` | CSVC, TBDH lớp mẫu giáo | Số trường: 237; Tỷ lệ phòng/lớp: 1.00 | ✅ Đạt 100% |
| | `0.M5` | Tình hình tài chính GDMN | Tổng chi NS: 12.304 tỷ; Tỷ lệ chuyên môn: 2.13% | ✅ Đạt 100% (*) |
| | `0.M6` | Trẻ em khuyết tật GDMN | Tổng trẻ KT: 296; Được can thiệp: 226 | ✅ Đạt 100% |
| | `0.M7` | Sổ theo dõi PCGD Mầm non | Tổng học sinh trong sổ: 107,556 trẻ | ✅ Đạt 100% |
| **Tiểu học (1.M)** | `1.M1` | Trẻ 6-14t PCGD & Bắt buộc TH | Trẻ 6t: 27,307; Tổng 6-10t: 140,711; Vào L1: 100% | ✅ Đạt 100% |
| | `1.M2` | Đội ngũ CBQL, GV, NV Tiểu học | Tổng GV: 8,028; Tỷ lệ GV/lớp: 1.62 | ✅ Đạt 100% |
| | `1.M3` | CSVC & Trang thiết bị Tiểu học | Điểm trường: 360; Tổng lớp: 4,984; TL Ph/lớp: 101.91% | ✅ Đạt 100% |
| | `1.M4` | Chuẩn PCGD & Bắt buộc TH | Trường TH: 160; Vào L1: 100%; Hoàn thành TH: 99.99% | ✅ Đạt 100% |
| | `1.M5` | Đối tượng khuyết tật Tiểu học | Tổng KT: 3,797; Tiếp cận giáo dục: 3,304 (87.10%) | ✅ Đạt 100% |
| | `1.M6` | Học sinh đi học ngoài địa bàn TH | Tổng HS ngoài địa bàn: 10,390 học sinh | ✅ Đạt 100% |
| | `1.M7` | Học sinh ngoại bàn đến học TH | Tổng HS nơi khác đến: 9,395 học sinh | ✅ Đạt 100% |
| **THCS (2.M)** | `2.M1` | Thanh thiếu niên PCGD THCS | TTN 15t: 27,620; Tốt nghiệp THCS: 97.56% | ✅ Đạt 100% |
| | `2.M2` | PCGD THCS theo độ tuổi | Vào Lớp 1: 27,046; 15-18t có bằng THCS: 98,796 | ✅ Đạt 100% |
| | `2.M3` | Hiệu quả PCGD THCS | TTN 11-18t: 230,682; Tốt nghiệp THCS: 98,765 | ✅ Đạt 100% |
| | `2.M4` | CSVC trường lớp THCS | Số lớp: 3,332; Tỷ lệ phòng/lớp: 97.66% | ✅ Đạt 100% |
| | `2.M5` | Đội ngũ giáo viên THCS | Tổng GV: 6,379; Tỷ lệ GV/lớp: 1.91 | ✅ Đạt 100% |
| | `2.M6` | Đối tượng khuyết tật THCS | Tổng KT: 3,601; Tiếp cận giáo dục: 2,869 (79.67%) | ✅ Đạt 100% |
| | `2.M7` | Học sinh đi học ngoài địa bàn THCS | Tổng HS ngoài địa bàn: 4,067 học sinh | ✅ Đạt 100% |
| | `2.M8` | Học sinh ngoại bàn đến học THCS | Tổng HS nơi khác đến: 10,965 học sinh | ✅ Đạt 100% |
| | `2.M9` | Đạt chuẩn PCGD THCS Mức 1-3 | Đối tượng 15-21t: 166,166; Tốt nghiệp THCS: 98,111 | ✅ Đạt 100% |
| | `2.M10` | Đạt chuẩn PCGD THCS Mức 4 | Đối tượng 18-21t: 88,368; Có bằng THPT/Nghề: 75,875 | ✅ Đạt 100% |
| **Xóa mù chữ (3.M)** | `3.M1` | Chuẩn XMC Mức 1 (15-25t) | Dân số: 244,515; Biết chữ: 244,352 (99.93%) | ✅ Đạt 100% |
| | `3.M2` | Chuẩn XMC Mức 2 (15-35t) | Dân số: 479,435; Biết chữ: 478,664 (99.84%) | ✅ Đạt 100% |
| | `3.M3` | Chuẩn XMC Chung (15-60t) | Dân số: 1,097,844; Biết chữ: 1,090,373 (99.32%) | ✅ Đạt 100% |
| | `3.M4` | Hiện trạng mù chữ M1 (15-25t) | Dân số: 244,515; Mù chữ: 163 (0.07%) | ✅ Đạt 100% |
| | `3.M5` | Hiện trạng mù chữ M1 (15-35t) | Dân số: 479,435; Mù chữ: 771 (0.16%) | ✅ Đạt 100% |
| | `3.M6` | Hiện trạng mù chữ M1 (15-60t) | Dân số: 1,097,844; Mù chữ: 7,471 (0.68%) | ✅ Đạt 100% |
| | `3.M7` | Hiện trạng mù chữ M2 (15-25t) | Dân số: 244,515; Mù chữ: 285 (0.12%) | ✅ Đạt 100% |
| | `3.M8` | Hiện trạng mù chữ M2 (15-35t) | Dân số: 479,435; Mù chữ: 1,918 (0.40%) | ✅ Đạt 100% |
| | `3.M9` | Hiện trạng mù chữ M2 (15-60t) | Dân số: 1,097,844; Mù chữ: 17,236 (1.57%) | ✅ Đạt 100% |

*(Ghi chú `*`: Biểu 0.M5 giữ nguyên trạng số liệu nộp thô theo chỉ đạo của Sở GD&ĐT, không tự ý chia hệ số).*

---

## 📁 CẤU TRÚC THƯ MỤC DỰ ÁN

```text
TONG HOP PHO CAP/
├── CAI_DAT_BAN_DAU.bat         # Tệp thiết lập môi trường tự động khi chuyển sang máy mới
├── KHOI_DONG_HE_THONG.bat      # Tệp khởi động ứng dụng siêu tốc & chống trùng tiến trình
├── DUNG_HE_THONG.bat           # Tệp dừng dịch vụ Streamlit an toàn
├── DAY_LEN_GITHUB.bat          # Tệp 1-Click đồng bộ và đẩy mã nguồn lên GitHub
├── README.md                   # Tài liệu hướng dẫn sử dụng chi tiết trên GitHub
├── HUONG_DAN_SU_DUNG_VA_BAN_GIAO.txt # Bản bàn giao nghiệp vụ chi tiết
├── requirements-lock.txt       # Danh sách thư viện và phiên bản chuẩn hóa đã khóa
├── MẪU CHUẨN - PCGD 2026(ok).xlsx # File mẫu chuẩn gốc gồm 37 sheet của Sở GD&ĐT
├── config/                     # Cấu hình danh mục 92 xã và metadata của 33 biểu
├── core/                       # Lõi đọc dữ liệu Excel, kiểm tra trùng lặp và độ mới
├── processors/                 # Các mô-đun xử lý thuật toán tổng hợp riêng từng biểu (0.M1 - 3.M9)
├── exporters/                  # Bộ xuất dữ liệu Excel chuẩn định dạng % và công thức
├── ui/                         # Giao diện Web trực quan (Streamlit)
├── data/                       # Chứa 92 file Excel dữ liệu nguồn nộp từ 92 xã/phường
├── outputs/                    # Chứa Sổ cái hợp nhất, Báo cáo đối soát và 33 file xuất thử
├── scripts/                    # Các script hỗ trợ kiểm thử, tạo shortcut, dừng port
└── tests/                      # Bộ kiểm thử tự động toàn diện pipeline 33 biểu
```

---

## 🛡️ NGUYÊN TẮC BẢO TOÀN DỮ LIỆU & NGHIỆP VỤ

1. **Bảo toàn dữ liệu nguồn:** Tuyệt đối không chỉnh sửa, không ghi đè, không di chuyển các file trong thư mục `data/` và file `MẪU CHUẨN - PCGD 2026(ok).xlsx`.
2. **Quy chuẩn tỷ lệ phần trăm (%):** Tỷ lệ cấp tỉnh được tính lại chuẩn nghiệp vụ $\frac{\sum \text{Tử số}}{\sum \text{Mẫu số}} \times 100$, lưu dưới dạng số thực theo thang 0–100 và định dạng hiển thị `cell.number_format = '0.00"%"'`. Tuyệt đối không cộng gộp trực tiếp các tỷ lệ của cấp xã.
3. **Phân biệt kiểu dữ liệu:** Minh bạch giữa số `0` thực tế, ô trống (`None`), ký hiệu `"x"` (không áp dụng) và lỗi sai kiểu dữ liệu; không biến ô thiếu dữ liệu thành 0 làm sai lệch thống kê.

---
**Đơn vị sử dụng:** Sở Giáo dục và Đào tạo tỉnh Thái Nguyên  
**Hỗ trợ kỹ thuật:** Antigravity AI (AG)

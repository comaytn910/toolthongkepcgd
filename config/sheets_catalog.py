# DANH MỤC 33 BIỂU NGHIỆP VỤ VÀ 4 SHEET HỖ TRỢ

SHEET_CATALOG = {
    # Khối Mầm non (0.M)
    '0.M1': {'name': 'Thống kê trẻ em MN theo độ tuổi', 'group': 'Mầm non', 'status': 'SẴN SÀNG', 'real_range': 'A1:K40', 'type': 'fixed'},
    '0.M2': {'name': 'Tiêu chuẩn, ĐK PCGD MN', 'group': 'Mầm non', 'status': 'SẴN SÀNG', 'real_range': 'A1:V9', 'type': 'fixed'},
    '0.M3': {'name': 'Đội ngũ CBQL, GV, NV MN', 'group': 'Mầm non', 'status': 'SẴN SÀNG', 'real_range': 'A1:S73', 'type': 'expandable'},
    '0.M4': {'name': 'Cơ sở vật chất MN', 'group': 'Mầm non', 'status': 'SẴN SÀNG', 'real_range': 'A1:X14', 'type': 'fixed'},
    '0.M5': {'name': 'Tình hình tài chính MN', 'group': 'Mầm non', 'status': 'SẴN SÀNG', 'real_range': 'A1:J38', 'type': 'fixed'},
    '0.M6': {'name': 'Trẻ khuyết tật MN', 'group': 'Mầm non', 'status': 'SẴN SÀNG', 'real_range': 'A2:M13', 'type': 'fixed'},
    '0.M7': {'name': 'Sổ theo dõi PCGD MN', 'group': 'Mầm non', 'status': 'SẴN SÀNG', 'real_range': 'A3:P50', 'type': 'expandable'},

    # Khối Tiểu học (1.M)
    '1.M1': {'name': 'Trẻ 6-14t PCGD & GD bắt buộc', 'group': 'Tiểu học', 'status': 'SẴN SÀNG', 'real_range': 'A2:AA63', 'type': 'fixed'},
    '1.M2': {'name': 'Đội ngũ cán bộ, GV Tiểu học', 'group': 'Tiểu học', 'status': 'SẴN SÀNG', 'real_range': 'A2:AG35', 'type': 'fixed'},
    '1.M3': {'name': 'CSVC cấp Tiểu học', 'group': 'Tiểu học', 'status': 'SẴN SÀNG', 'real_range': 'A1:AC101', 'type': 'expandable'},
    '1.M4': {'name': 'Kết quả PCGD Tiểu học', 'group': 'Tiểu học', 'status': 'SẴN SÀNG', 'real_range': 'A1:U116', 'type': 'fixed'},
    '1.M5': {'name': 'Trẻ khuyết tật Tiểu học', 'group': 'Tiểu học', 'status': 'SẴN SÀNG', 'real_range': 'A3:N17', 'type': 'fixed'},
    '1.M6': {'name': 'HS học xã/tỉnh ngoài (Tiểu học)', 'group': 'Tiểu học', 'status': 'SẴN SÀNG', 'real_range': 'A2:AR49', 'type': 'expandable'},
    '1.M7': {'name': 'HS ngoài học tại xã (Tiểu học)', 'group': 'Tiểu học', 'status': 'SẴN SÀNG', 'real_range': 'A2:AR68', 'type': 'expandable'},

    # Khối THCS (2.M)
    '2.M1': {'name': 'Thống kê TTN PCGD THCS', 'group': 'THCS', 'status': 'SẴN SÀNG', 'real_range': 'A1:Z59', 'type': 'fixed'},
    '2.M2': {'name': 'Tổng hợp huy động PCGD THCS', 'group': 'THCS', 'status': 'SẴN SÀNG', 'real_range': 'A1:W48', 'type': 'fixed'},
    '2.M3': {'name': 'Thống kê PCGD THCS', 'group': 'THCS', 'status': 'SẴN SÀNG', 'real_range': 'A1:S46', 'type': 'fixed'},
    '2.M4': {'name': 'CSVC cấp THCS', 'group': 'THCS', 'status': 'SẴN SÀNG', 'real_range': 'A1:X12', 'type': 'fixed'},
    '2.M5': {'name': 'Đội ngũ giáo viên THCS', 'group': 'THCS', 'status': 'SẴN SÀNG', 'real_range': 'A1:AQ17', 'type': 'expandable'},
    '2.M6': {'name': 'Trẻ khuyết tật THCS', 'group': 'THCS', 'status': 'SẴN SÀNG', 'real_range': 'A2:M14', 'type': 'fixed'},
    '2.M7': {'name': 'HS học ngoài địa bàn (THCS)', 'group': 'THCS', 'status': 'SẴN SÀNG', 'real_range': 'A1:AM40', 'type': 'expandable'},
    '2.M8': {'name': 'HS ngoài học tại xã (THCS)', 'group': 'THCS', 'status': 'SẴN SÀNG', 'real_range': 'A2:AM39', 'type': 'expandable'},
    '2.M9': {'name': 'Huy động GD Trung học', 'group': 'THCS', 'status': 'SẴN SÀNG', 'real_range': 'A2:V48', 'type': 'fixed'},
    '2.M10': {'name': 'Hiệu quả GD Trung học', 'group': 'THCS', 'status': 'SẴN SÀNG', 'real_range': 'A2:T47', 'type': 'fixed'},

    # Khối Xóa mù chữ (3.M)
    '3.M1': {'name': 'Chuẩn XMC Mức 1', 'group': 'Xóa mù chữ', 'status': 'SẴN SÀNG', 'real_range': 'A2:G48', 'type': 'fixed'},
    '3.M2': {'name': 'Chuẩn XMC Mức 2', 'group': 'Xóa mù chữ', 'status': 'SẴN SÀNG', 'real_range': 'A2:G49', 'type': 'fixed'},
    '3.M3': {'name': 'Chuẩn XMC Chung toàn xã', 'group': 'Xóa mù chữ', 'status': 'SẴN SÀNG', 'real_range': 'A2:G49', 'type': 'fixed'},
    '3.M4': {'name': 'Hiện trạng mù chữ M1 (15-25t)', 'group': 'Xóa mù chữ', 'status': 'SẴN SÀNG', 'real_range': 'A2:O48', 'type': 'fixed'},
    '3.M5': {'name': 'Hiện trạng mù chữ M1 (15-35t)', 'group': 'Xóa mù chữ', 'status': 'SẴN SÀNG', 'real_range': 'A2:O47', 'type': 'fixed'},
    '3.M6': {'name': 'Hiện trạng mù chữ M1 (15-60t)', 'group': 'Xóa mù chữ', 'status': 'SẴN SÀNG', 'real_range': 'A2:O47', 'type': 'fixed'},
    '3.M7': {'name': 'Hiện trạng mù chữ M2 (15-25t)', 'group': 'Xóa mù chữ', 'status': 'SẴN SÀNG', 'real_range': 'A3:N48', 'type': 'fixed'},
    '3.M8': {'name': 'Hiện trạng mù chữ M2 (15-35t)', 'group': 'Xóa mù chữ', 'status': 'SẴN SÀNG', 'real_range': 'A2:N47', 'type': 'fixed'},
    '3.M9': {'name': 'Hiện trạng mù chữ M2 (15-60t)', 'group': 'Xóa mù chữ', 'status': 'SẴN SÀNG', 'real_range': 'A2:N47', 'type': 'fixed'},
}

SUPPORTING_SHEETS = [
    'ĐỐI_SOÁT',
    'TỔNG_QUAN',
    'CAU_HINH_DOI_SOAT',
    'DS_SAO_LUU_20261012_180009_1147'
]

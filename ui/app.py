import streamlit as st
import os
import pandas as pd
from core.excel_reader import scan_data_directory
from processors.m0_m1 import process_0_m1
from processors.m0_m2 import process_0_m2
from processors.m0_m3 import process_0_m3
from processors.m0_m4 import process_0_m4
from processors.m0_m5 import process_0_m5
from processors.m0_m6 import process_0_m6
from processors.m0_m7 import process_0_m7
from processors.m1_m1 import process_1_m1
from processors.m1_m2 import process_1_m2
from processors.m1_m3 import process_1_m3
from processors.m1_m4 import process_1_m4
from processors.m1_m5 import process_1_m5
from processors.m1_m6 import process_1_m6
from processors.m1_m7 import process_1_m7
from processors.m2_m1 import process_2_m1
from processors.m2_m2 import process_2_m2
from processors.m2_m3 import process_2_m3
from processors.m2_m4 import process_2_m4
from processors.m2_m5 import process_2_m5
from processors.m2_m6 import process_2_m6

from exporters.excel_exporter import (
    export_0_m1_test,
    export_0_m2_test,
    export_0_m3_test,
    export_0_m4_test,
    export_0_m5_test,
    export_0_m6_test,
    export_0_m7_test,
    export_1_m1_test,
    export_1_m2_test,
    export_1_m3_test,
    export_1_m4_test,
    export_1_m5_test,
    export_1_m6_test,
    export_1_m7_test,
    export_2_m1_test,
    export_2_m2_test,
    export_2_m3_test,
    export_2_m4_test,
    export_2_m5_test,
    export_2_m6_test,
    export_2_m7_test,
    export_2_m8_test,
    export_2_m9_test,
    export_2_m10_test,
    export_3_m1_test,
    export_3_m2_test,
    export_3_m3_test,
    export_3_m4_test,
    export_3_m5_test,
    export_3_m6_test,
    export_3_m7_test,
    export_3_m8_test,
    export_3_m9_test,
    export_3_mu_chu_test
)
from processors.m2_m7 import process_2_m7
from processors.m2_m8 import process_2_m8
from processors.m2_m9 import process_2_m9
from processors.m2_m10 import process_2_m10
from processors.m3_m1 import process_3_m1_single_file, aggregate_3_m1
from processors.m3_m2 import process_3_m2_single_file, aggregate_3_m2
from processors.m3_m3 import process_3_m3_single_file, aggregate_3_m3
from processors.m3_mu_chu import process_3_mu_chu_single_file, aggregate_3_mu_chu
from config.sheets_catalog import SHEET_CATALOG

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Tổng hợp PCGD 2026 - Thái Nguyên",
    page_icon="📊",
    layout="wide"
)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, 'data')
OUTPUTS_DIR = os.path.join(BASE_DIR, 'outputs')
TEMPLATE_PATH = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')

# CSS tùy biến giao diện tiếng Việt
st.markdown("""
<style>
    .main-title {
        color: #003366;
        font-size: 26px;
        font-weight: bold;
        text-align: center;
        margin-bottom: 5px;
    }
    .sub-title {
        color: #555555;
        font-size: 16px;
        text-align: center;
        margin-bottom: 25px;
    }
    .status-badge-ok {
        background-color: #d4edda;
        color: #155724;
        padding: 4px 10px;
        border-radius: 4px;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

from core.data_freshness import check_data_freshness, save_sync_metadata
METADATA_PATH = os.path.join(OUTPUTS_DIR, 'last_sync_metadata.json')

st.markdown('<div class="main-title">SỞ GIÁO DỤC VÀ ĐÀO TẠO TỈNH THÁI NGUYÊN</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">HỆ THỐNG TỔNG HỢP SỐ LIỆU PHỔ CẬP GIÁO DỤC & XÓA MÙ CHỮ NĂM 2026</div>', unsafe_allow_html=True)

# KIỂM TRA ĐỘ MỚI CỦA SỐ LIỆU TỔNG HỢP SO VỚI NGUỒN DATA/
freshness = check_data_freshness(METADATA_PATH, DATA_DIR)
if freshness['status'] == 'FRESH':
    st.success(f"🟢 **DỮ LIỆU ĐỒNG BỘ**: {freshness['message']}")
elif freshness['status'] == 'STALE':
    st.error(f"🔴 **CẢNH BÁO SỐ LIỆU ĐANG CŨ**: {freshness['message']}")
    with st.expander(f"🔍 Xem chi tiết {len(freshness['changed_files'])} tệp nguồn đã thay đổi sau lần tổng hợp gần nhất ({freshness['last_sync_time']})"):
        for cf in freshness['changed_files']:
            st.write(f"- 📄 `{cf}`")
        st.info("💡 **Lưu ý**: Dữ liệu nguồn trong `data/` đã được chỉnh sửa. Các file Sổ cái và Báo cáo đối soát hiện tại chưa phản ánh những thay đổi mới này.")
else:
    st.warning("🟡 **CHƯA GHI NHẬN ĐỒNG BỘ**: Chưa tìm thấy nhật ký tổng hợp gần nhất. Vui lòng thực hiện tổng hợp dữ liệu toàn tỉnh.")

# Sidebar điều khiển
with st.sidebar:
    st.header("⚙️ ĐIỀU KHIỂN HỆ THỐNG")
    st.info(f"📁 **Thư mục nguồn:** `data/`\n\n📁 **Thư mục xuất:** `outputs/`")
    
    btn_rescan = st.button("🔄 Quét lại thư mục data", use_container_width=True)
    if btn_rescan:
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

# 1. QUÉT VÀ KIỂM TRA THƯ MỤC DATA
st.subheader("1. Kiểm tra khả năng đọc file trong thư mục `data/`")
scan_res = scan_data_directory(DATA_DIR)

if 'error' in scan_res and not scan_res['files']:
    st.warning(f"⚠️ {scan_res['error']}")
    st.stop()

files_data = []
for f in scan_res['files']:
    status_label = "✅ HỢP LỆ" if f['is_valid'] else "❌ CÓ LỖI"
    notes = "; ".join(f['errors']) if f['errors'] else ("Không lỗi" if not f['warnings'] else "; ".join(f['warnings']))
    files_data.append({
        'STT Xã': f['commune_stt'] if f['commune_stt'] else "-",
        'Tên Xã / Phường nhận diện': f['commune_name'],
        'Tên file nguồn': f['filename'],
        'Số biểu tìm thấy': f"{f['biz_sheets_count']} / 33",
        'Biểu 0.M1': "Có" if f['has_0_m1'] else "Thiếu",
        'Trạng thái': status_label,
        'Ghi chú': notes
    })

df_files = pd.DataFrame(files_data)
st.dataframe(df_files, use_container_width=True, hide_index=True)

valid_files = [f for f in scan_res['files'] if f['is_valid']]
valid_count = len(valid_files)

col_stat1, col_stat2, col_stat3 = st.columns(3)
col_stat1.metric("Tổng số file phát hiện", len(scan_res['files']))
col_stat2.metric("Số file hợp lệ", valid_count)
col_stat3.metric("Xã bị trùng lặp", len(scan_res['duplicate_communes']))

st.divider()

# KHU VỰC TẢI TÀI LIỆU CẤP TỈNH & BÁO CÁO THANH TRA
st.markdown("### 🏆 TÀI LIỆU CẤP TỈNH & BÁO CÁO ĐỐI SOÁT THANH TRA BỘ GDĐT")
col_exp1, col_exp2 = st.columns(2)

master_book_path = os.path.join(OUTPUTS_DIR, "SO_CAI_TONG_HOP_PCGD_2026_THAI_NGUYEN.xlsx")
audit_report_path = os.path.join(OUTPUTS_DIR, "BAO_CAO_DOI_SOAT_63_QUY_TAC_PCGD_2026.xlsx")

with col_exp1:
    st.info("📕 **SỔ CÁI TỔNG HỢP HỢP NHẤT TOÀN TỈNH (33 BIỂU)**")
    st.caption("Đóng gói toàn bộ 33 biểu nghiệp vụ chuẩn (Mầm non, Tiểu học, THCS, Xóa mù chữ) vào 01 file Excel duy nhất cấp tỉnh.")
    if os.path.exists(master_book_path):
        with open(master_book_path, 'rb') as f_m:
            st.download_button(
                label="📥 Tải Sổ cái tổng hợp 33 biểu (.xlsx)",
                data=f_m.read(),
                file_name="SO_CAI_TONG_HOP_PCGD_2026_THAI_NGUYEN.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
    else:
        st.warning("Chưa tìm thấy tệp Sổ cái hợp nhất.")

with col_exp2:
    st.info("📋 **BÁO CÁO ĐỐI SOÁT 63 QUY TẮC PHỤC VỤ THANH TRA**")
    st.caption("Báo cáo thẩm định chất lượng dữ liệu: Đánh giá 62 quy tắc cấp tỉnh, ma trận 92 xã và chi tiết 198 sai lệch nguồn cơ sở.")
    if os.path.exists(audit_report_path):
        with open(audit_report_path, 'rb') as f_a:
            st.download_button(
                label="📥 Tải Báo cáo đối soát 63 quy tắc (.xlsx)",
                data=f_a.read(),
                file_name="BAO_CAO_DOI_SOAT_63_QUY_TAC_PCGD_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="secondary",
                use_container_width=True
            )
    else:
        st.warning("Chưa tìm thấy tệp Báo cáo đối soát.")

st.divider()

# 2. CHỌN BIỂU TỔNG HỢP
st.subheader("2. Lựa chọn biểu tổng hợp")
READY_SHEETS = list(SHEET_CATALOG.keys())

selected_sheet = st.selectbox(
    "Chọn biểu nghiệp vụ cần xử lý:",
    options=list(SHEET_CATALOG.keys()),
    format_func=lambda k: f"{k} - {SHEET_CATALOG[k]['name']} [{SHEET_CATALOG[k]['status']}]"
)

if selected_sheet not in READY_SHEETS:
    st.warning(f"⚠️ Biểu **{selected_sheet} - {SHEET_CATALOG[selected_sheet]['name']}** hiện đang ở trạng thái **CHƯA TRIỂN KHAI**.")
    st.stop()

# 3. TỔNG HỢP VÀ XEM TRƯỚC
sheet_name_title = SHEET_CATALOG[selected_sheet]['name']
st.subheader(f"3. Thử nghiệm tổng hợp biểu `{selected_sheet}: {sheet_name_title}`")

col_btn1, col_btn2 = st.columns([1, 3])
with col_btn1:
    btn_calc = st.button("🚀 Tổng hợp lại (Tính mới hoàn toàn)", type="primary", use_container_width=True)

state_key = f"result_{selected_sheet}"
out_key = f"out_path_{selected_sheet}"

if btn_calc:
    with st.spinner(f"Đang tính toán độc lập từ các file trong data/ cho biểu {selected_sheet}..."):
        if selected_sheet == '0.M1':
            res = process_0_m1(valid_files)
            out_path = export_0_m1_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '0.M2':
            res = process_0_m2(valid_files)
            out_path = export_0_m2_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '0.M3':
            res = process_0_m3(valid_files)
            out_path = export_0_m3_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '0.M4':
            res = process_0_m4(valid_files)
            out_path = export_0_m4_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '0.M5':
            res = process_0_m5(valid_files)
            out_path = export_0_m5_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '0.M6':
            res = process_0_m6(valid_files)
            out_path = export_0_m6_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '0.M7':
            res = process_0_m7(valid_files, collect_records=True)
            out_path = export_0_m7_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '1.M1':
            res = process_1_m1(valid_files)
            out_path = export_1_m1_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '1.M2':
            res = process_1_m2(valid_files)
            out_path = export_1_m2_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '1.M3':
            res = process_1_m3(valid_files)
            out_path = export_1_m3_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '1.M4':
            res = process_1_m4(valid_files)
            out_path = export_1_m4_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '1.M5':
            res = process_1_m5(valid_files)
            out_path = export_1_m5_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '1.M6':
            res = process_1_m6(valid_files)
            out_path = export_1_m6_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '1.M7':
            res = process_1_m7(valid_files)
            out_path = export_1_m7_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '2.M1':
            res = process_2_m1(valid_files)
            out_path = export_2_m1_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '2.M2':
            res = process_2_m2(valid_files)
            out_path = export_2_m2_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '2.M3':
            res = process_2_m3(valid_files)
            out_path = export_2_m3_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '2.M4':
            res = process_2_m4(valid_files)
            out_path = export_2_m4_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '2.M5':
            res = process_2_m5(valid_files)
            out_path = export_2_m5_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '2.M6':
            res = process_2_m6(valid_files)
            out_path = export_2_m6_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '2.M7':
            res = process_2_m7(valid_files)
            out_path = export_2_m7_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '2.M8':
            res = process_2_m8(valid_files)
            out_path = export_2_m8_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '2.M9':
            res = process_2_m9(valid_files)
            out_path = export_2_m9_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '2.M10':
            res = process_2_m10(valid_files)
            out_path = export_2_m10_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
        elif selected_sheet == '3.M1':
            c_data = [process_3_m1_single_file(f['filepath'], f['commune_name']) for f in valid_files]
            res = aggregate_3_m1(c_data)
            out_path = export_3_m1_test(res, TEMPLATE_PATH, os.path.join(OUTPUTS_DIR, f"TONG_HOP_THU_{selected_sheet}_PCGD_2026.xlsx"))
        elif selected_sheet == '3.M2':
            c_data = [process_3_m2_single_file(f['filepath'], f['commune_name']) for f in valid_files]
            res = aggregate_3_m2(c_data)
            out_path = export_3_m2_test(res, TEMPLATE_PATH, os.path.join(OUTPUTS_DIR, f"TONG_HOP_THU_{selected_sheet}_PCGD_2026.xlsx"))
        elif selected_sheet == '3.M3':
            c_data = [process_3_m3_single_file(f['filepath'], f['commune_name']) for f in valid_files]
            res = aggregate_3_m3(c_data)
            out_path = export_3_m3_test(res, TEMPLATE_PATH, os.path.join(OUTPUTS_DIR, f"TONG_HOP_THU_{selected_sheet}_PCGD_2026.xlsx"))
        elif selected_sheet in ['3.M4', '3.M5', '3.M6', '3.M7', '3.M8', '3.M9']:
            c_data = [process_3_mu_chu_single_file(f['filepath'], f['commune_name'], selected_sheet) for f in valid_files]
            res = aggregate_3_mu_chu(c_data, selected_sheet)
            export_map_mu_chu = {
                '3.M4': export_3_m4_test,
                '3.M5': export_3_m5_test,
                '3.M6': export_3_m6_test,
                '3.M7': export_3_m7_test,
                '3.M8': export_3_m8_test,
                '3.M9': export_3_m9_test,
            }
            exporter = export_map_mu_chu.get(selected_sheet, export_3_mu_chu_test)
            out_path = exporter(res, TEMPLATE_PATH, os.path.join(OUTPUTS_DIR, f"TONG_HOP_THU_{selected_sheet}_PCGD_2026.xlsx"))
            
        st.session_state[state_key] = res
        st.session_state[out_key] = out_path
        st.success(f"✅ Đã tổng hợp xong biểu {selected_sheet} và xuất file kết quả: `{os.path.basename(out_path)}`")


if state_key in st.session_state:
    res = st.session_state[state_key]
    out_p = st.session_state.get(out_key, "")
    
    tab1, tab2 = st.tabs(["📋 Bảng số liệu tổng hợp", "📄 Xuất file Excel kết quả"])
    
    with tab1:
        st.info(f"📊 Đã tổng hợp thành công từ **{len(valid_files)} xã thực tế** trong thư mục `data/`.")
        if selected_sheet == '0.M7':
            m1, m2, m3 = st.columns(3)
            m1.metric("Tổng số trẻ trong sổ toàn tỉnh", f"{res['total_children']:,} trẻ")
            m2.metric("Số trẻ em nữ", f"{res['total_female']:,} trẻ")
            m3.metric("Tỷ lệ nữ", f"{res['rate_female']:.2f}%")
            
            st.markdown("##### 🏫 Bảng tổng hợp số lượng trẻ từng xã/phường (92 xã)")
            df_communes = pd.DataFrame(res['commune_summary'])
            st.dataframe(df_communes, use_container_width=True, hide_index=True)
            
            st.markdown("##### 👥 Thống kê cơ cấu theo dân tộc")
            eth_df = pd.DataFrame([{'Dân tộc': k, 'Số trẻ': v, 'Tỷ lệ': f"{v/res['total_children']*100:.2f}%"} for k, v in sorted(res['ethnicity_summary'].items(), key=lambda x: x[1], reverse=True)])
            st.dataframe(eth_df, use_container_width=True, hide_index=True)
        elif selected_sheet == '1.M2':
            agg = res['aggregated_cells']
            crit = res['criteria']
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Tổng số GV Tiểu học (tính chuẩn)", f"{agg['I8']:,.0f} GV")
            m2.metric("Biên chế", f"{agg['J8']:,.0f} GV")
            m3.metric("Hợp đồng", f"{agg['K8']:,.0f} GV")
            m4.metric("Tỉ lệ GV/lớp chuẩn toàn tỉnh", f"{agg['N8']:.2f}")

            c1, c2, c3 = st.columns(3)
            c1.metric("Đạt trình độ Đại học", f"{crit['dai_hoc']['sl']:,.0f} GV", f"{crit['dai_hoc']['tl']:.2f}%")
            c2.metric("Thạc sĩ trở lên", f"{crit['thac_si']['sl']:,.0f} GV", f"{crit['thac_si']['tl']:.2f}%")
            c3.metric("Chuẩn nghề nghiệp đạt trở lên", f"{crit['chuan_nghe_nghiep']['sl']:,.0f} GV", f"{crit['chuan_nghe_nghiep']['tl']:.2f}%")

            if res.get('discrepancies'):
                st.warning(f"⚠️ Phát hiện {len(res['discrepancies'])} xã có sai lệch giữa Tổng GV nguồn (I8) và Biên chế + Hợp đồng (J8+K8):")
                st.dataframe(pd.DataFrame(res['discrepancies']), use_container_width=True, hide_index=True)
        elif selected_sheet == '1.M3':
            agg = res['aggregated_cells']
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Tổng số Điểm trường (C8)", f"{agg['C8']:,.0f} điểm")
            m2.metric("Tổng số Lớp học TH (D8)", f"{agg['D8']:,.0f} lớp")
            m3.metric("Tổng số Phòng học", f"{res['total_rooms']:,.0f} phòng")
            m4.metric("Tỉ lệ Phòng/Lớp cấp tỉnh (J8)", f"{agg['J8']:.2f}%")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Phòng kiên cố (F8)", f"{agg['F8']:,.0f} phòng", f"{agg['F8']/res['total_rooms']*100:.1f}%")
            c2.metric("Bán kiên cố (G8)", f"{agg['G8']:,.0f} phòng", f"{agg['G8']/res['total_rooms']*100:.1f}%")
            c3.metric("Phòng tạm (H8)", f"{agg['H8']:,.0f} phòng")
            c4.metric("Phòng thuê/mượn (I8)", f"{agg['I8']:,.0f} phòng")

            if res.get('discrepancies'):
                st.warning(f"⚠️ Phát hiện {len(res['discrepancies'])} xã chèn dòng cộng phụ làm nhân đôi số liệu hàng 8 (đã được tự động chuẩn hóa):")
                st.dataframe(pd.DataFrame(res['discrepancies']), use_container_width=True, hide_index=True)
        elif selected_sheet == '1.M4':
            agg = res['aggregated_cells']
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Tổng trường Tiểu học (C8)", f"{agg['C8']:,.0f} trường")
            m2.metric("Tổng HS Tiểu học (D8)", f"{agg['D8']:,.0f} hs")
            m3.metric("Trẻ khuyết tật TH (E8)", f"{agg['E8']:,.0f} trẻ")
            m4.metric("Chuẩn PCGD Tiểu học", f"{res['provincial_standard']}")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Trẻ 6t vào L1 (F8)", f"{agg['F8']:,.0f} trẻ", f"{agg['G8']:.2f}%")
            c2.metric("Trẻ 11t HTCTTH (H8)", f"{agg['H8']:,.0f} trẻ", f"{agg['I8']:.2f}%")
            c3.metric("Trẻ đến 14t HTCTTH (J8)", f"{agg['J8']:,.0f} trẻ", f"{agg['K8']:.2f}%")
            c4.metric("Trẻ KT tiếp cận GD (P8)", f"{agg['P8']:,.0f} trẻ", f"{agg['Q8']:.2f}%")

            st.markdown("##### 🏆 Thống kê kết quả công nhận đạt chuẩn PCGDTH 92 xã")
            std_counts = res.get('standard_counts', {})
            df_std = pd.DataFrame([{'Mức độ đạt chuẩn PCGDTH': k, 'Số lượng xã': v, 'Tỷ lệ': f"{v/len(valid_files)*100:.1f}%"} for k, v in std_counts.items()])
            st.dataframe(df_std, use_container_width=True, hide_index=True)
        elif selected_sheet == '1.M5':
            agg = res['aggregated_cells']
            m1, m2, m3 = st.columns(3)
            m1.metric("Tổng trẻ khuyết tật TH (C17)", f"{agg['C17']:,.0f} trẻ")
            m2.metric("Số trẻ tiếp cận giáo dục (L17)", f"{agg['L17']:,.0f} trẻ")
            m3.metric("Tỉ lệ tiếp cận giáo dục (M17)", f"{agg['M17']:.2f}%")

            st.markdown("##### ♿ Cơ cấu 8 dạng khuyết tật toàn tỉnh (Cột D..K hàng 17)")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("1. Trí tuệ (H17)", f"{agg['H17']:,.0f} trẻ", f"{agg['H17']/agg['C17']*100:.1f}%")
            c2.metric("2. Vận động (D17)", f"{agg['D17']:,.0f} trẻ", f"{agg['D17']/agg['C17']*100:.1f}%")
            c3.metric("3. Thần kinh, tâm thần (G17)", f"{agg['G17']:,.0f} trẻ", f"{agg['G17']/agg['C17']*100:.1f}%")
            c4.metric("4. Khác (K17)", f"{agg['K17']:,.0f} trẻ", f"{agg['K17']/agg['C17']*100:.1f}%")

            c5, c6, c7, c8 = st.columns(4)
            c5.metric("5. Nghe, nói (E17)", f"{agg['E17']:,.0f} trẻ", f"{agg['E17']/agg['C17']*100:.1f}%")
            c6.metric("6. Nhìn (F17)", f"{agg['F17']:,.0f} trẻ", f"{agg['F17']/agg['C17']*100:.1f}%")
            c7.metric("7. Học tập (J17)", f"{agg['J17']:,.0f} trẻ", f"{agg['J17']/agg['C17']*100:.1f}%")
            c8.metric("8. Tự kỷ (I17)", f"{agg['I17']:,.0f} trẻ", f"{agg['I17']/agg['C17']*100:.1f}%")

            if res.get('h17_discrepancies'):
                st.warning(f"⚠️ Phát hiện {len(res['h17_discrepancies'])} xã có dòng tổng 17 nguồn lệch với chi tiết (Đã tự động tính chuẩn từ các dòng chi tiết):")
                st.dataframe(pd.DataFrame(res['h17_discrepancies']), use_container_width=True, hide_index=True)
        elif selected_sheet == '1.M6':
            m1, m2, m3 = st.columns(3)
            m1.metric("Tổng học sinh đi học ngoài địa bàn (AR49)", f"{res['total_grand']:,.0f} hs")
            m2.metric("Học ngoài xã trong tỉnh (AR27)", f"{res['total_part1']:,.0f} hs", f"{res['total_part1']/res['total_grand']*100:.1f}%")
            m3.metric("Học ngoài tỉnh (AR48)", f"{res['total_part2']:,.0f} hs", f"{res['total_part2']/res['total_grand']*100:.1f}%")

            st.markdown("##### 🎒 Phân bố học sinh đi học ngoài theo độ tuổi cấp Tiểu học (6–10 tuổi)")
            c1, c2, c3, c4, c5 = st.columns(5)
            by_age = res['province_by_age']
            c1.metric("6 Tuổi", f"{by_age['6 Tuổi']['total']:,.0f} hs", f"{by_age['6 Tuổi']['pct']:.1f}%")
            c2.metric("7 Tuổi", f"{by_age['7 Tuổi']['total']:,.0f} hs", f"{by_age['7 Tuổi']['pct']:.1f}%")
            c3.metric("8 Tuổi", f"{by_age['8 Tuổi']['total']:,.0f} hs", f"{by_age['8 Tuổi']['pct']:.1f}%")
            c4.metric("9 Tuổi", f"{by_age['9 Tuổi']['total']:,.0f} hs", f"{by_age['9 Tuổi']['pct']:.1f}%")
            c5.metric("10 Tuổi", f"{by_age['10 Tuổi']['total']:,.0f} hs", f"{by_age['10 Tuổi']['pct']:.1f}%")

            st.markdown("##### 🏫 Top xã/phường có nhiều học sinh đi học ngoài địa bàn nhất")
            top_communes = sorted(res['commune_summaries'], key=lambda x: x['total_sum'], reverse=True)[:10]
            df_top = pd.DataFrame([{
                'STT': item['stt'],
                'Xã / Phường': item['commune_name'],
                'Học ngoài xã (trong tỉnh)': f"{item['part1_sum']:,.0f}",
                'Học ngoài tỉnh': f"{item['part2_sum']:,.0f}",
                'Tổng cộng': f"{item['total_sum']:,.0f}"
            } for item in top_communes])
            st.dataframe(df_top, use_container_width=True, hide_index=True)
        elif selected_sheet == '1.M7':
            m1, m2, m3 = st.columns(3)
            m1.metric("Tổng HS ngoại địa bàn đến học (AR7)", f"{res['total_grand']:,.0f} hs")
            m2.metric("Ngoài xã trong tỉnh đến học (AR34)", f"{res['total_part1']:,.0f} hs", f"{res['total_part1']/res['total_grand']*100:.1f}%")
            m3.metric("Ngoài tỉnh đến học (AR68)", f"{res['total_part2']:,.0f} hs", f"{res['total_part2']/res['total_grand']*100:.1f}%")

            st.markdown("##### 🎒 Phân bố học sinh ngoại địa bàn theo độ tuổi cấp Tiểu học (6–10 tuổi)")
            c1, c2, c3, c4, c5 = st.columns(5)
            by_age = res['province_by_age']
            c1.metric("6 Tuổi", f"{by_age['6 Tuổi']['total']:,.0f} hs", f"{by_age['6 Tuổi']['total']/res['total_grand']*100:.1f}%")
            c2.metric("7 Tuổi", f"{by_age['7 Tuổi']['total']:,.0f} hs", f"{by_age['7 Tuổi']['total']/res['total_grand']*100:.1f}%")
            c3.metric("8 Tuổi", f"{by_age['8 Tuổi']['total']:,.0f} hs", f"{by_age['8 Tuổi']['total']/res['total_grand']*100:.1f}%")
            c4.metric("9 Tuổi", f"{by_age['9 Tuổi']['total']:,.0f} hs", f"{by_age['9 Tuổi']['total']/res['total_grand']*100:.1f}%")
            c5.metric("10 Tuổi", f"{by_age['10 Tuổi']['total']:,.0f} hs", f"{by_age['10 Tuổi']['total']/res['total_grand']*100:.1f}%")

            st.markdown("##### 🏫 Top xã/phường tiếp nhận nhiều học sinh ngoại địa bàn nhất")
            top_communes = sorted(res['commune_summaries'], key=lambda x: x['total_sum'], reverse=True)[:10]
            df_top = pd.DataFrame([{
                'STT': item['stt'],
                'Xã / Phường': item['commune_name'],
                'Học ngoài xã (trong tỉnh)': f"{item['part1_sum']:,.0f}",
                'Học ngoài tỉnh': f"{item['part2_sum']:,.0f}",
                'Tổng cộng': f"{item['total_sum']:,.0f}"
            } for item in top_communes])
            st.dataframe(df_top, use_container_width=True, hide_index=True)

            if res.get('discrepancies'):
                st.warning(f"⚠️ Phát hiện {len(res['discrepancies'])} xã có sai lệch hàng tổng nguồn (Đã tự động tính chuẩn sạch sẽ bằng Clean Recalculation):")
                st.dataframe(pd.DataFrame(res['discrepancies']), use_container_width=True, hide_index=True)
        elif selected_sheet == '2.M2':
            p_cnt = res['province_counts']
            p_rt = res['province_rates']

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Trẻ 6 tuổi toàn tỉnh (C48)", f"{p_cnt['C']:,.0f} trẻ")
            m2.metric("Trẻ 6t vào Lớp 1 (D48)", f"{p_cnt['D']:,.0f} trẻ", f"{p_rt['E']:.2f}%")
            m3.metric("HS HTCTTH năm qua (F48)", f"{p_cnt['F']:,.0f} hs")
            m4.metric("Vào học Lớp 6 (I48)", f"{p_cnt['I']:,.0f} hs", f"{p_rt['J']:.2f}%")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Trẻ 11-14t toàn tỉnh (K48)", f"{p_cnt['K']:,.0f} trẻ")
            c2.metric("Trẻ 11-14t HTCTTH (L48)", f"{p_cnt['L']:,.0f} trẻ", f"{p_rt['M']:.2f}%")
            c3.metric("HS lớp 9 năm qua (N48)", f"{p_cnt['N']:,.0f} hs")
            c4.metric("Tốt nghiệp THCS L9 (O48)", f"{p_cnt['O']:,.0f} hs", f"{p_rt['Q']:.2f}%")

            d1, d2, d3 = st.columns(3)
            d1.metric("Đối tượng 15-18 tuổi (R48)", f"{p_cnt['R']:,.0f} người")
            d2.metric("Có bằng tốt nghiệp THCS (U48)", f"{p_cnt['U']:,.0f} người", f"{p_rt['V']:.2f}%")
            d3.metric("Đánh giá PCGD THCS toàn tỉnh", f"{res['provincial_standard']}")

            if res.get('discrepancies'):
                st.warning(f"⚠️ Phát hiện {len(res['discrepancies'])} sai lệch nguồn tại các xã (Đã tự động chuẩn hóa bằng Clean Recalculation):")
                st.dataframe(pd.DataFrame(res['discrepancies']), use_container_width=True, hide_index=True)
        elif selected_sheet == '2.M3':
            p_cnt = res['province_counts']
            p_rt = res['province_rates']

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("TTN 11-18 tuổi (D46)", f"{p_cnt['D']:,.0f} người")
            m2.metric("Khuyết tật 11-18t (E46)", f"{p_cnt['E']:,.0f} người")
            m3.metric("Tổng số PPC 15-18t (H46)", f"{p_cnt['H']:,.0f} người")
            m4.metric("Tốt nghiệp THCS 15-18t (I46)", f"{p_cnt['I']:,.0f} người", f"{p_rt['J']:.2f}%")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Học tiếp THPT/GDTX/GDNN (K46)", f"{p_cnt['K']:,.0f} người", f"{p_rt['L']:.2f}%")
            c2.metric("Trẻ KT tổng số (M46)", f"{p_cnt['M']:,.0f} trẻ")
            c3.metric("KT có khả năng học (N46)", f"{p_cnt['N']:,.0f} trẻ")
            c4.metric("KT tiếp cận GD (O46)", f"{p_cnt['O']:,.0f} trẻ", f"{p_rt['P']:.2f}%")

            d1, d2, d3 = st.columns(3)
            d1.metric("Chuẩn GD Tiểu học / XMC", "Mức độ 3 / Mức độ 2")
            d2.metric("Điều kiện GV-NV & CSVC", "Đạt / Đạt")
            d3.metric("Đánh giá PCGD THCS toàn tỉnh", f"{res['provincial_standard']}")

            if res.get('discrepancies'):
                st.info(f"ℹ️ Ghi nhận {len(res['discrepancies'])} xã có số trường THCS chỉ khai báo ở dòng thôn mà ô dòng tổng mẫu để số tĩnh 0 (Đã ghi nhận trong sheet DOI_CHIEU_SAI_LECH):")
                st.dataframe(pd.DataFrame(res['discrepancies']), use_container_width=True, hide_index=True)
        elif selected_sheet == '2.M4':
            p_cnt = res['province_counts']

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Số trường THCS phát hiện", f"{res['total_schools']} trường")
            m2.metric("Số điểm trường (C12)", f"{p_cnt['C']:,.0f} điểm")
            m3.metric("Số lớp học THCS (D12)", f"{p_cnt['D']:,.0f} lớp")
            m4.metric("Tỉ lệ Phòng/Lớp THCS (H12)", f"{res['province_rate_h']:.2f}%", f"{res['province_tot_rooms']:,.0f} phòng")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Phòng Kiên cố (E12)", f"{p_cnt['E']:,.0f} phòng")
            c2.metric("Phòng Bán KC (F12)", f"{p_cnt['F']:,.0f} phòng")
            c3.metric("Phòng Tạm (G12)", f"{p_cnt['G']:,.0f} phòng")
            c4.metric("Phòng Thí nghiệm (O12)", f"{p_cnt['O']:,.0f} phòng", f"{p_cnt['P']:,.1f} m²")

            d1, d2, d3, d4 = st.columns(4)
            d1.metric("Hiệu trưởng / P.Hiệu trưởng", f"{p_cnt['I']:,.0f} / {p_cnt['J']:,.0f} phòng")
            d2.metric("Văn phòng / Y tế", f"{p_cnt['K']:,.0f} / {p_cnt['L']:,.0f} phòng")
            d3.metric("Khu vệ sinh GV / HS", f"{p_cnt['Q']:,.0f} / {p_cnt['S']:,.0f} khu")
            d4.metric("Sân chơi / Bãi tập", f"{p_cnt['U']:,.0f} / {p_cnt['W']:,.0f} khu")
        elif selected_sheet == '2.M5':
            p_cnt = res['province_counts']
            crit = res['criteria']

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Tổng số GV THCS (H12)", f"{p_cnt['H']:,.0f} GV")
            m2.metric("Biên chế (I12)", f"{p_cnt['I']:,.0f} GV")
            m3.metric("Hợp đồng (J12)", f"{p_cnt['J']:,.0f} GV")
            m4.metric("Tỉ lệ GV/lớp THCS (M12)", f"{res['prov_ratio_gv_class']:.2f} GV/lớp")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("CBQL Hiệu trưởng / Phó HT", f"{p_cnt['F']:,.0f} / {p_cnt['G']:,.0f} người")
            c2.metric("Đạt chuẩn trình độ (ĐH trở lên)", f"{crit['dat_chuan']['sl']:,.0f} GV", f"{crit['dat_chuan']['tl']:.2f}%")
            c3.metric("Trên chuẩn (Thạc sĩ/Trên ĐH)", f"{crit['tren_chuan']['sl']:,.0f} GV", f"{crit['tren_chuan']['tl']:.2f}%")
            c4.metric("Đạt chuẩn nghề nghiệp", f"{crit['chuan_nghe_nghiep']['sl']:,.0f} GV", f"{crit['chuan_nghe_nghiep']['tl']:.2f}%")

            if res.get('discrepancies'):
                st.warning(f"⚠️ Phát hiện {len(res['discrepancies'])} xã có sai lệch nguồn giữa Tổng số GV và Biên chế + Hợp đồng (Đã ghi nhận đối chiếu chi tiết trong Sheet DOI_CHIEU_SAI_LECH):")
                st.dataframe(pd.DataFrame(res['discrepancies']), use_container_width=True, hide_index=True)
        elif selected_sheet == '2.M6':
            agg = res['aggregated_cells']

            m1, m2, m3 = st.columns(3)
            m1.metric("Tổng đối tượng KT THCS (C14)", f"{res['total_kt']:,.0f} người")
            m2.metric("Tiếp cận giáo dục (L14)", f"{res['total_access']:,.0f} người")
            m3.metric("Tỷ lệ tiếp cận GD toàn tỉnh (M14)", f"{res['rate_access']:.2f}%")

            st.markdown("##### 📊 Thống kê chi tiết đối tượng khuyết tật THCS theo từng độ tuổi (11 - 18 tuổi):")
            age_data = []
            birth_years = {
                6: ("2015", "11"), 7: ("2014", "12"), 8: ("2013", "13"), 9: ("2012", "14"),
                10: ("2011", "15"), 11: ("2010", "16"), 12: ("2009", "17"), 13: ("2008", "18")
            }
            for r, (by, age) in birth_years.items():
                age_data.append({
                    'Năm sinh': by,
                    'Độ tuổi': f"{age} tuổi",
                    'Tổng số KT (C)': f"{agg[f'C{r}']:,.0f}",
                    'Vận động (D)': f"{agg[f'D{r}']:,.0f}",
                    'Nghe nói (E)': f"{agg[f'E{r}']:,.0f}",
                    'Nhìn (F)': f"{agg[f'F{r}']:,.0f}",
                    'Thần kinh (G)': f"{agg[f'G{r}']:,.0f}",
                    'Trí tuệ (H)': f"{agg[f'H{r}']:,.0f}",
                    'Tự kỷ (I)': f"{agg[f'I{r}']:,.0f}",
                    'Học tập (J)': f"{agg[f'J{r}']:,.0f}",
                    'Khác (K)': f"{agg[f'K{r}']:,.0f}",
                    'Tiếp cận GD (L)': f"{agg[f'L{r}']:,.0f}",
                    'Tỷ lệ % (M)': f"{agg[f'M{r}']:.2f}%"
                })
            st.dataframe(pd.DataFrame(age_data), use_container_width=True, hide_index=True)
        elif selected_sheet == '2.M7':
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Học ngoài xã (trong tỉnh)", f"{res['total_part1']:,.0f} hs")
            m2.metric("Học ngoài tỉnh", f"{res['total_part2']:,.0f} hs")
            m3.metric("Tổng cộng đi học ngoài", f"{res['total_grand']:,.0f} hs")
            m4.metric("Lượt trường ghi nhận", f"{res['total_schools_count']:,} lượt")

            st.markdown("##### 📊 Thống kê số lượng học sinh học ngoài địa bàn theo từng độ tuổi (10 - 18 tuổi):")
            age_data = []
            for age_lbl, d in res['age_breakdown'].items():
                age_data.append({
                    'Độ tuổi': age_lbl,
                    'Ngoài xã (trong tỉnh)': f"{d['part1']:,.0f}",
                    'Ngoài tỉnh': f"{d['part2']:,.0f}",
                    'Tổng cộng (I + II)': f"{d['total']:,.0f}",
                    'Tỷ trọng (%)': f"{d['total']/res['total_grand']*100:.2f}%" if res['total_grand'] > 0 else "0.00%"
                })
            st.dataframe(pd.DataFrame(age_data), use_container_width=True, hide_index=True)

            st.markdown(f"##### 🏫 Danh sách các trường học ngoài địa bàn được khai báo ({len(res['school_records'])} lượt):")
            df_sch = pd.DataFrame(res['school_records'])
            df_sch.columns = ['STT Xã', 'Xã/Phường khai báo', 'Tên trường ngoài địa bàn', 'Phân loại', 'Tổng số học sinh']
            st.dataframe(df_sch[['STT Xã', 'Xã/Phường khai báo', 'Phân loại', 'Tên trường ngoài địa bàn', 'Tổng số học sinh']], use_container_width=True, hide_index=True)
        elif selected_sheet == '2.M8':
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Ngoài xã trong tỉnh đến học", f"{res['total_part1']:,.0f} hs")
            m2.metric("Ngoài tỉnh đến học", f"{res['total_part2']:,.0f} hs")
            m3.metric("Tổng cộng ngoại địa bàn", f"{res['total_grand']:,.0f} hs")
            m4.metric("Lượt trường ghi nhận", f"{res['total_schools_count']:,} lượt")

            st.markdown("##### 📊 Thống kê học sinh ngoại địa bàn đến học theo từng độ tuổi (10 - 18 tuổi):")
            age_data = []
            for age_lbl, d in res['age_breakdown'].items():
                age_data.append({
                    'Độ tuổi': age_lbl,
                    'Trong tỉnh đến học': f"{d['part1']:,.0f}",
                    'Ngoài tỉnh đến học': f"{d['part2']:,.0f}",
                    'Tổng cộng (I + II)': f"{d['total']:,.0f}",
                    'Tỷ trọng (%)': f"{d['total']/res['total_grand']*100:.2f}%" if res['total_grand'] > 0 else "0.00%"
                })
            st.dataframe(pd.DataFrame(age_data), use_container_width=True, hide_index=True)

            if res.get('discrepancies'):
                st.warning(f"⚠️ Phát hiện {len(res['discrepancies'])} xã có sai lệch nguồn ô AM (chứa số 39 tĩnh hoặc gõ thiếu). Hệ thống đã Clean Recalculation bảo toàn chính xác từng học sinh:")
                st.dataframe(pd.DataFrame(res['discrepancies'])[['commune_stt', 'commune_name', 'part', 'calc_am', 'src_am', 'diff', 'reason']], use_container_width=True, hide_index=True)

            st.markdown(f"##### 🏫 Danh sách các trường học được khai báo ({len(res['school_records'])} lượt):")
            df_sch = pd.DataFrame(res['school_records'])
            st.dataframe(df_sch, use_container_width=True, hide_index=True)
        elif selected_sheet == '2.M9':
            p_cnt = res['province_counts']
            p_rt = res['province_rates']

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Đối tượng PC 15-21 tuổi (C48)", f"{p_cnt['C']:,.0f} người")
            m2.metric("TN THCS 15-18 tuổi (E48/D48)", f"{p_cnt['E']:,.0f} / {p_cnt['D']:,.0f}", f"{p_rt['F']:.2f}%")
            m3.metric("TN THCS năm qua (J48/I48)", f"{p_cnt['J']:,.0f} / {p_cnt['I']:,.0f}", f"{p_rt['K']:.2f}%")
            m4.metric("TN THPT, TCCN, Dạy nghề (V48)", f"{p_rt['V']:.2f}%")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Vào lớp 10 THPT (L48)", f"{p_cnt['L']:,.0f} hs", f"{p_cnt['L']/p_cnt['I']*100:.2f}%")
            c2.metric("Học GDTX, Nghề (M48)", f"{p_cnt['M']:,.0f} hs", f"{p_cnt['M']/p_cnt['I']*100:.2f}%")
            c3.metric("Tỷ lệ vào TCCN (T48)", f"{p_cnt['R']:,.0f} hs", f"{p_rt['T']:.2f}%")
            c4.metric("Tỷ lệ vào Dạy nghề (U48)", f"{p_cnt['S']:,.0f} hs", f"{p_rt['U']:.2f}%")

            st.markdown("##### 🏫 Bảng số liệu chi tiết kết quả huy động phổ cập giáo dục trung học từng xã/phường (92 xã):")
            df_92 = []
            for item in res['commune_summaries']:
                c_cnt = item['counts']
                c_rt = item['rates']
                df_92.append({
                    'STT': item['stt'],
                    'Xã/Phường': item['commune_name'],
                    'Dòng nguồn': item['row_total'],
                    'Tổng PC 15-21 (C)': f"{c_cnt['C']:,.0f}",
                    'Phải PC 15-18 (D)': f"{c_cnt['D']:,.0f}",
                    'TN THCS 15-18 (E)': f"{c_cnt['E']:,.0f}",
                    'Tỷ lệ 15-18 (F)': f"{c_rt['F']:.2f}%",
                    'Lớp 9 năm qua (I)': f"{c_cnt['I']:,.0f}",
                    'TN năm qua (J)': f"{c_cnt['J']:,.0f}",
                    'Tỷ lệ TN năm qua (K)': f"{c_rt['K']:.2f}%",
                    'Vào lớp 10 (L)': f"{c_cnt['L']:,.0f}",
                    'GDTX/Nghề (M)': f"{c_cnt['M']:,.0f}",
                    'Bỏ học THPT (N)': f"{c_cnt['N']:,.0f}",
                    'TCCN (R)': f"{c_cnt['R']:,.0f}",
                    'Dạy nghề (S)': f"{c_cnt['S']:,.0f}",
                    'Tỷ lệ TN THPT/Nghề (V)': f"{c_rt['V']:.2f}%"
                })
            st.dataframe(pd.DataFrame(df_92), use_container_width=True, hide_index=True)
        elif selected_sheet == '2.M10':
            p_cnt = res['province_counts']
            p_rt = res['province_rates']

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Đối tượng PC 18-21 tuổi (C47)", f"{p_cnt['C']:,.0f} người", f"Phải PC: {p_cnt['D']:,.0f}")
            tot_bang = p_cnt['E'] + p_cnt['F'] + p_cnt['G']
            m2.metric("Có bằng TN THPT/BT/Nghề", f"{tot_bang:,.0f} người", f"{p_rt['H']:.2f}%")
            m3.metric("TN THPT năm qua (J47/I47)", f"{p_cnt['J']:,.0f} / {p_cnt['I']:,.0f}", f"{p_rt['K']:.2f}%")
            m4.metric("Có bằng đào tạo nghề (S47/R47)", f"{p_cnt['S']:,.0f} / {p_cnt['R']:,.0f}", f"{p_rt['T']:.2f}%")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Đi học Đại học (O47)", f"{p_cnt['O']:,.0f} người")
            c2.metric("Đi học Cao đẳng (P47)", f"{p_cnt['P']:,.0f} người")
            c3.metric("Đi học Trung cấp (Q47)", f"{p_cnt['Q']:,.0f} người")
            c4.metric("TN GDTX/TCCN (M47/L47)", f"{p_cnt['M']:,.0f} / {p_cnt['L']:,.0f}", f"{p_rt['N']:.2f}%")

            st.markdown("##### 🏫 Bảng số liệu chi tiết hiệu quả phổ cập giáo dục trung học từng xã/phường (92 xã):")
            df_92 = []
            for item in res['commune_summaries']:
                c_cnt = item['counts']
                c_rt = item['rates']
                tot_b_c = c_cnt['E'] + c_cnt['F'] + c_cnt['G']
                df_92.append({
                    'STT': item['stt'],
                    'Xã/Phường': item['commune_name'],
                    'Dòng nguồn': item['row_total'],
                    'Tổng PC 18-21 (C)': f"{c_cnt['C']:,.0f}",
                    'Phải PC 18-21 (D)': f"{c_cnt['D']:,.0f}",
                    'TN THPT (E)': f"{c_cnt['E']:,.0f}",
                    'TN BT THPT (F)': f"{c_cnt['F']:,.0f}",
                    'TN THCN/TDN (G)': f"{c_cnt['G']:,.0f}",
                    'Tỷ lệ có bằng (H)': f"{c_rt['H']:.2f}%",
                    'Lớp 12 năm qua (I)': f"{c_cnt['I']:,.0f}",
                    'TN THPT năm qua (J)': f"{c_cnt['J']:,.0f}",
                    'Tỷ lệ TN THPT (K)': f"{c_rt['K']:.2f}%",
                    'Đại học (O)': f"{c_cnt['O']:,.0f}",
                    'Cao đẳng (P)': f"{c_cnt['P']:,.0f}",
                    'Trung cấp (Q)': f"{c_cnt['Q']:,.0f}",
                    'Học nghề >=3 năm (R)': f"{c_cnt['R']:,.0f}",
                    'Có bằng nghề (S)': f"{c_cnt['S']:,.0f}",
                    'Tỷ lệ bằng nghề (T)': f"{c_rt['T']:.2f}%"
                })
            st.dataframe(pd.DataFrame(df_92), use_container_width=True, hide_index=True)
        elif selected_sheet in ['3.M1', '3.M2', '3.M3']:
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Tổng dân số trong độ tuổi (C)", f"{res['sum_c_dan_so']:,.0f} người")
            m2.metric("Biết chữ Mức 1 (D)", f"{res['sum_d_biet_chu_m1']:,.0f} người", f"{res['rate_m1_prov']:.2f}%")
            m3.metric("Biết chữ Mức 2 (F)", f"{res['sum_f_biet_chu_m2']:,.0f} người", f"{res['rate_m2_prov']:.2f}%")
            m4.metric("Xã đạt chuẩn M1 / M2", f"{res['count_dat_m1']} / {res['count_dat_m2']} xã")

            st.markdown(f"##### 🏫 Bảng số liệu chi tiết 92 xã/phường ({selected_sheet}):")
            df_92 = []
            for item in res['communes_detail']:
                df_92.append({
                    'Xã/Phường': item['commune_name'],
                    'Dân số (C)': f"{item['c_dan_so']:,.0f}",
                    'Biết chữ M1 (D)': f"{item['d_biet_chu_m1']:,.0f}",
                    'Tỷ lệ M1 (%)': f"{item['e_rate_m1_calc']:.2f}%",
                    'Biết chữ M2 (F)': f"{item['f_biet_chu_m2']:,.0f}",
                    'Tỷ lệ M2 (%)': f"{item['g_rate_m2_calc']:.2f}%",
                    'Đạt chuẩn M1': "Đạt" if item['dat_m1'] else "Chưa đạt",
                    'Đạt chuẩn M2': "Đạt" if item['dat_m2'] else "Chưa đạt",
                })
            st.dataframe(pd.DataFrame(df_92), use_container_width=True, hide_index=True)
        elif selected_sheet in ['3.M4', '3.M5', '3.M6', '3.M7', '3.M8', '3.M9']:
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Tổng dân số (C)", f"{res['sum_c_dan_so']:,.0f} người", f"Nữ: {res['sum_d_nu']:,.0f}")
            m2.metric(f"Mù chữ {res['muc_do']} (G)", f"{res['sum_g_mu_chu']:,.0f} người", f"{res['rate_h_prov']:.2f}%")
            m3.metric("Mù chữ Nữ (I)", f"{res['sum_i_mu_chu_nu']:,.0f} người", f"{res['rate_j_prov']:.2f}%")
            m4.metric("Mù chữ Dân tộc (K)", f"{res['sum_k_mu_chu_dt']:,.0f} người", f"{res['rate_l_prov']:.2f}%")

            st.markdown(f"##### 🏫 Bảng chi tiết hiện trạng mù chữ {res['muc_do']} ({res['age_label']}) 92 xã:")
            df_92 = []
            for item in res['communes_detail']:
                df_92.append({
                    'Xã/Phường': item['commune_name'],
                    'Tổng DS (C)': f"{item['c_dan_so']:,.0f}",
                    'Nữ (D)': f"{item['d_nu']:,.0f}",
                    'Dân tộc (E)': f"{item['e_dan_toc']:,.0f}",
                    'Mù chữ (G)': f"{item['g_mu_chu']:,.0f}",
                    'Tỷ lệ (%) (H)': f"{item['h_rate_calc']:.2f}%",
                    'Mù chữ Nữ (I)': f"{item['i_mu_chu_nu']:,.0f}",
                    'Tỷ lệ Nữ (%) (J)': f"{item['j_rate_nu_calc']:.2f}%",
                    'Mù chữ DT (K)': f"{item['k_mu_chu_dt']:,.0f}",
                    'Tỷ lệ DT (%) (L)': f"{item['l_rate_dt_calc']:.2f}%",
                })
            st.dataframe(pd.DataFrame(df_92), use_container_width=True, hide_index=True)
        elif 'aggregated_cells' in res:
            agg = res['aggregated_cells']
            # Hiển thị tóm tắt một số chỉ tiêu nổi bật
            preview_items = []
            for k, v in list(agg.items())[:50]:
                if isinstance(v, (int, float)):
                    fmt_v = f"{v:,.2f}%" if ('H' in k and selected_sheet in ['2.M1', '0.M1']) or ('G' in k and selected_sheet == '1.M1' and int(k[1:]) >= 58) or ('M' in k and selected_sheet == '0.M6') else (f"{v:,.2f}" if isinstance(v, float) else f"{v:,}")
                    preview_items.append({'Ô / Chỉ tiêu': k, 'Giá trị': fmt_v})
            st.dataframe(pd.DataFrame(preview_items), use_container_width=True, hide_index=True)
            
    with tab2:
        if out_p and os.path.exists(out_p):
            st.write(f"📁 **File kết quả đã được ghi vào:** `{out_p}` ({os.path.getsize(out_p):,} bytes)")
            with open(out_p, "rb") as f:
                st.download_button(
                    label=f"📥 Tải xuống file Excel kết quả biểu {selected_sheet}",
                    data=f.read(),
                    file_name=os.path.basename(out_p),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

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
from processors.m2_m1 import process_2_m1

from exporters.excel_exporter import (
    export_0_m1_test,
    export_0_m2_test,
    export_0_m3_test,
    export_0_m4_test,
    export_0_m5_test,
    export_0_m6_test,
    export_0_m7_test,
    export_1_m1_test,
    export_2_m1_test
)
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

st.markdown('<div class="main-title">SỞ GIÁO DỤC VÀ ĐÀO TẠO TỈNH THÁI NGUYÊN</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">HỆ THỐNG TỔNG HỢP SỐ LIỆU PHỔ CẬP GIÁO DỤC & XÓA MÙ CHỮ NĂM 2026</div>', unsafe_allow_html=True)

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

# 2. CHỌN BIỂU TỔNG HỢP
st.subheader("2. Lựa chọn biểu tổng hợp")
READY_SHEETS = ['0.M1', '0.M2', '0.M3', '0.M4', '0.M5', '0.M6', '0.M7', '1.M1', '2.M1']

selected_sheet = st.selectbox(
    "Chọn biểu nghiệp vụ cần xử lý:",
    options=list(SHEET_CATALOG.keys()),
    format_func=lambda k: f"{k} - {SHEET_CATALOG[k]['name']} [{SHEET_CATALOG[k]['status']}]"
)

if selected_sheet not in READY_SHEETS:
    st.warning(f"⚠️ Biểu **{selected_sheet} - {SHEET_CATALOG[selected_sheet]['name']}** hiện đang ở trạng thái **CHƯA TRIỂN KHAI**. Các biểu đã hoàn thành và sẵn sàng xử lý: **{', '.join(READY_SHEETS)}**.")
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
        elif selected_sheet == '2.M1':
            res = process_2_m1(valid_files)
            out_path = export_2_m1_test(res, TEMPLATE_PATH, OUTPUTS_DIR)
            
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

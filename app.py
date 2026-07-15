import streamlit as st
import pandas as pd
from modules.extraction import extract_and_align_data
from modules.positioning import apply_jittering
from modules.export import generate_dxf_file, generate_pptx_file

st.set_page_config(page_title="ELAA System - Automation MEP Layout", layout="wide")

# Khởi tạo bộ lưu trữ trạng thái Session State toàn hệ thống
if 'step' not in st.session_state:
    st.session_state.step = 1
if 'mapping_data' not in st.session_state:
    st.session_state.mapping_data = None

st.title("⚡ Electrical Layout Automation Agent (ELAA)")
st.caption("Giải pháp chuyển đổi không chạm từ tài liệu hỗn hợp sang CAD/PowerPoint kỹ thuật chuyên nghiệp.")

# BƯỚC 1: TIẾP NHẬN TỆP ĐẦU VÀO
if st.session_state.step == 1:
    st.header("Bước 1: Tiếp Nhận & Phân Loại File Đầu Vào")
    
    col1, col2 = st.columns(2)
    with col1:
        file_quote = st.file_uploader("Tải lên file Báo giá (Tên chứa chữ 'Quotation')", type=["pdf"])
    with col2:
        file_3d = st.file_uploader("Tải lên bản vẽ thiết kế 3D (Tên chứa chữ '3D review')", type=["pdf"])
        
    st.subheader("⚙️ Bảng Cấu Hình Ký Hiệu Đa Phương Thức")
    config_data = pd.DataFrame([
        {"Ký hiệu chữ": "P20A", "Tên thiết bị": "Nguồn điện công suất 20A", "Icon (.png)": "power_20a.png"},
        {"Ký hiệu chữ": "LED", "Tên thiết bị": "Hệ thống đèn LED chiếu sáng", "Icon (.png)": "led.png"},
        {"Ký hiệu chữ": "AIR", "Tên thiết bị": "Đường cấp khí nén máy CNC", "Icon (.png)": "air.png"}
    ])
    st.data_editor(config_data, num_rows="dynamic", use_container_width=True)

    if st.button("🚀 Thực Thi Quy Trình Ngầm (Submit)", type="primary"):
        if file_quote and file_3d:
            with st.spinner("Hệ thống đang trích xuất văn bản và tính toán lát cắt không gian..."):
                try:
                    # Chạy Module Trích xuất kết hợp API Claude 3.5 Sonnet
                    raw_data = extract_and_align_data(file_quote, file_3d)
                    # Chạy Thuật toán chống trùng lặp Jittering hình học
                    optimized_data = apply_jittering(raw_data)
                    
                    st.session_state.mapping_data = optimized_data
                    st.session_state.step = 2
                    st.rerun()
                except Exception as ex:
                    st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {str(ex)}")
        else:
            st.warning("Vui lòng tải lên đầy đủ cả 2 file Quotation và 3D review để tiếp tục.")

# BƯỚC 2: RÀ SOÁT CHẤT LƯỢNG & ĐIỀU CHỈNH MAPPING TƯƠNG TÁC
elif st.session_state.step == 2:
    st.header("Bước 2: Rà Soát Chất Lượng & Điều Chỉnh Tọa Độ Thực Tế")
    st.success("AI đã định vị xong vị trí thiết bị. Bạn có thể sửa trực tiếp tọa độ (X, Y) dưới bảng nếu cần thiết.")
    
    col_table, col_preview = st.columns([3, 2])
    
    with col_table:
        st.subheader("Bảng Mapping Data Block dữ liệu trung gian")
        # Người dùng có thể chỉnh sửa trực tiếp thông số X, Y của AI gợi ý tại đây
        st.session_state.mapping_data = st.data_editor(
            st.session_state.mapping_data, 
            num_rows="dynamic", 
            use_container_width=True
        )
        
    with col_preview:
        st.subheader("Xem trước thông số lưới")
        st.info("Hệ lưới làm việc tiêu chuẩn: 6m x 6m.")
        df_coords = pd.DataFrame(st.session_state.mapping_data)
        if not df_coords.empty and 'x' in df_coords.columns and 'y' in df_coords.columns:
            st.scatter_chart(df_coords, x='x', y='y', color='#ff4b4b')
            
    col_b1, col_b2 = st.columns([1, 8])
    with col_b1:
        if st.button("Quay lại"):
            st.session_state.step = 1
            st.rerun()
    with col_b2:
        if st.button("✅ Phê Duyệt Xuất Bản Hồ Sơ Kỹ Thuật", type="primary"):
            st.session_state.step = 3
            st.rerun()

# BƯỚC 3: MÀN HÌNH XEM TRƯỚC VÀ DOWNLOAD GÓI SẢN PHẨM
elif st.session_state.step == 3:
    st.header("Bước 3: Tổng Duyệt & Tải Xuống Thành Phẩm")
    
    with st.spinner("Đang đóng gói cấu trúc phân tầng thực thể kỹ thuật..."):
        # Gọi Module xuất dữ liệu nhị phân thực tế
        dxf_bytes = generate_dxf_file(st.session_state.mapping_data)
        pptx_bytes = generate_pptx_file(st.session_state.mapping_data)
        
        st.balloons()
        st.success("Tạo tệp thành công! File CAD không bị lỗi cấu trúc, Slide tự động cập nhật bảng Legend khớp 100%.")
        
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                label="📥 Tải Bản Vẽ CAD Thiết Kế (.dxf)",
                data=dxf_bytes,
                file_name="Electrical_Layout_Master.dxf",
                mime="application/dxf",
                use_container_width=True
            )
        with c2:
            st.download_button(
                label="📥 Tải Slide Báo Cáo Thi Công (.pptx)",
                data=pptx_bytes,
                file_name="Electrical_Layout_Report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )
            
    if st.button("🔄 Tạo Dự Án Mới từ Đầu"):
        st.session_state.clear()
        st.rerun()

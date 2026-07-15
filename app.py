import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from modules.extraction import extract_and_align_data
from modules.export import generate_pptx_file

st.set_page_config(page_title="ELAA System", layout="wide")

if 'step' not in st.session_state: st.session_state.step = 1
if 'mapping_data' not in st.session_state: st.session_state.mapping_data = []

st.title("⚡ Electrical Layout Automation Agent (ELAA)")

# BƯỚC 1: UPLOAD TÀI LIỆU & ICON
if st.session_state.step == 1:
    st.header("1. Upload Tài Liệu & Cấu Hình Icon")
    
    col1, col2 = st.columns(2)
    with col1:
        # Đã hỗ trợ Excel (.xlsx)
        file_quote = st.file_uploader("Upload Báo giá (Quotation - PDF/Excel)", type=["pdf", "xlsx", "xls"])
    with col2:
        file_3d = st.file_uploader("Upload Bản vẽ mặt bằng 3D (PDF)", type=["pdf"])
        
    st.subheader("🖼️ Upload bộ Icon / Ký hiệu (Tùy chọn)")
    uploaded_icons = st.file_uploader("Tải lên các file hình ảnh icon (.png, .jpg)", type=["png", "jpg"], accept_multiple_files=True)
    if uploaded_icons:
        st.success(f"Đã tải lên {len(uploaded_icons)} icon. Sẵn sàng tích hợp!")

    if st.button("🚀 Thực Thi Phân Tích (Submit)", type="primary"):
        if file_quote and file_3d:
            with st.spinner("Đang trích xuất dữ liệu Excel/PDF và phân tích tọa độ không gian..."):
                st.session_state.mapping_data = extract_and_align_data(file_quote, file_3d)
                st.session_state.step = 2
                st.rerun()
        else:
            st.warning("Vui lòng tải lên đủ Báo giá và Bản vẽ 3D.")

# BƯỚC 2: BẢN ĐỒ KÉO THẢ & RÀ SOÁT
elif st.session_state.step == 2:
    st.header("2. Rà Soát Tọa Độ & Kéo Thả Trực Quan")
    st.info("💡 MẸO: Bạn có thể dùng chuột KÉO THẢ trực tiếp các chữ/ký hiệu trên bản đồ bên dưới. Hoặc sửa số liệu X, Y trong bảng.")
    
    col_map, col_table = st.columns([2, 1])
    
    with col_table:
        st.write("Bảng dữ liệu (Có thể sửa trực tiếp số):")
        edited_data = st.data_editor(st.session_state.mapping_data, num_rows="dynamic")
        st.session_state.mapping_data = edited_data
        
    with col_map:
        # VẼ BẢN ĐỒ PLOTLY (CHỈ HIỆN X, Y DƯƠNG)
        fig = go.Figure()
        
        # Thêm các thiết bị lên bản đồ dưới dạng Text (hoặc Icon)
        for item in edited_data:
            fig.add_trace(go.Scatter(
                x=[item.get('x', 0)], 
                y=[item.get('y', 0)],
                mode="markers+text",
                name=item.get('item_name', 'Device'),
                text=[item.get('item_name', 'Device')[:10]], # Hiển thị tên rút gọn
                textposition="top center",
                marker=dict(size=15, color="red")
            ))

        # Ép khung bản đồ chỉ hiển thị 1/4 góc dương (Quadrant 1)
        fig.update_layout(
            xaxis=dict(range=[0, 10], title="Trục X (mét)"), # Bắt đầu từ 0
            yaxis=dict(range=[0, 10], title="Trục Y (mét)"), # Bắt đầu từ 0
            width=700, height=500,
            showlegend=False,
            # Nếu bạn có file ảnh nền, cấu hình images=[dict(source="link_anh", ...)] ở đây
        )
        
        # BẬT TÍNH NĂNG KÉO THẢ (Editable) CỦA PLOTLY CHO NGƯỜI DÙNG
        st.plotly_chart(fig, use_container_width=True, config={
            'editable': True, 
            'edits': {
                'shapePosition': True, 
                'annotationPosition': True
            }
        })

    col_back, col_export = st.columns([1, 5])
    with col_back:
        if st.button("Quay lại"):
            st.session_state.step = 1
            st.rerun()
    with col_export:
        if st.button("✅ Chốt Tọa Độ & Tạo File PowerPoint", type="primary"):
            st.session_state.step = 3
            st.rerun()

# BƯỚC 3: XUẤT FILE POWERPOINT
elif st.session_state.step == 3:
    st.header("3. Hoàn Tất Xuất Bản")
    with st.spinner("Đang tạo Slide PowerPoint..."):
        pptx_bytes = generate_pptx_file(st.session_state.mapping_data)
        st.success("Tạo file PowerPoint thành công!")
        
        st.download_button(
            label="📥 Tải File PowerPoint (.pptx)",
            data=pptx_bytes,
            file_name="Electrical_Layout.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )
            
    if st.button("🔄 Làm mới quy trình"):
        st.session_state.clear()
        st.rerun()

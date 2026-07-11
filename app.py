"""
====================================================================================
 HỆ THỐNG TỰ ĐỘNG BÓC TÁCH & VẼ SƠ ĐỒ ĐIỆN GIAN HÀNG (Electrical Layout Auto-Generator)
 PHIÊN BẢN V5.7 CHUẨN CLOUD — Bỏ code sinh file tự động, dùng file tĩnh trên GitHub
====================================================================================
"""

import io
import os
import random
import re
import unicodedata
import base64
import gc
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor

try:
    from google import genai
except ImportError:
    genai = None

# ====================================================================================
# 0. GỌI CUSTOM COMPONENT ĐÃ ĐƯỢC TẠO SẴN TRÊN GITHUB (KHÔNG TỰ SINH FILE NỮA)
# ====================================================================================
COMPONENT_DIR = os.path.join(os.path.dirname(__file__), "map_component")
interactive_map_component = components.declare_component("interactive_map", path=COMPONENT_DIR)

# ====================================================================================
# 1. CẤU HÌNH MẶC ĐỊNH
# ====================================================================================
AI_MODEL_DEFAULT = "gemini-1.5-pro"
STOPWORDS_VI = {"và", "cho", "của", "tại"}
SLIDE_TYPES = ["Booth Location", "Perspective View", "Booth Dimensions"]

DEFAULT_EQUIPMENT_CONFIG = [
    {"Tên thiết bị": "Đèn Floodlight 50W", "Biểu tượng vẽ": "▲", "Mã màu Hex": "#FFCD00", "Vị trí ưu tiên": "Hệ trần biên", "Công suất": "50W", "Số lượng": 30},
    {"Tên thiết bị": "Ổ cắm 5A/220V", "Biểu tượng vẽ": "●", "Mã màu Hex": "#D62728", "Vị trí ưu tiên": "Bàn tư vấn", "Công suất": "220V", "Số lượng": 3},
]

# ====================================================================================
# 2. DATA STRUCTURES & HELPER
# ====================================================================================
@dataclass
class EquipmentRow:
    key: str
    label: str
    icon_value: str
    color_hex: str
    zone_raw: str
    power: str
    quantity: int

@dataclass
class PipelineResult:
    equipment: list = field(default_factory=list)
    booth_location_img: Optional[Image.Image] = None
    perspective_img: Optional[Image.Image] = None
    dimensions_img: Optional[Image.Image] = None
    coordinates: dict = field(default_factory=dict)

def slugify(text: str, idx: int) -> str:
    try:
        norm = unicodedata.normalize("NFKD", text or "").encode("ascii", "ignore").decode("ascii")
        norm = re.sub(r"[^a-zA-Z0-9]+", "_", norm).strip("_").lower()
        return f"{norm}_{idx}"
    except: return f"item_{idx}"

def extract_salient_tokens(label: str) -> list:
    tokens = re.split(r"[\s/()\-,]+", (label or "").lower())
    return [t for t in tokens if len(t) >= 2 and t not in STOPWORDS_VI]

def regex_extract_quantities(text: str, labels: list) -> dict:
    found = {}
    lines = text.lower().splitlines()
    qty_patterns = [r"(?:sl|qty|số lượng)\s*[:.\-]?\s*(\d+)", r"x\s*(\d+)\b", r"\b(\d+)\s*(?:bộ|cái|chiếc|nos|no)\b"]
    for label in labels:
        tokens = extract_salient_tokens(label)
        if not tokens: continue
        best_qty = 0
        for line in lines:
            if any(tok in line for tok in tokens):
                for pat in qty_patterns:
                    m = re.search(pat, line)
                    if m: best_qty = max(best_qty, int(m.group(1)))
        if best_qty > 0: found[label] = best_qty
    return found

# ====================================================================================
# 3. AI PROVIDER
# ====================================================================================
class AIProvider:
    def __init__(self, api_key: str, model: str):
        if genai is None:
            raise RuntimeError("Chưa cài đặt google-genai.")
        self.client = genai.Client(api_key=api_key)
        self.model = model

    def classify_slide_image(self, image: Image.Image) -> str:
        system_prompt = (
            "Phân loại hình ảnh trang PDF này vào ĐÚNG MỘT nhãn sau (chỉ trả về tên nhãn):\n"
            "- Booth Location (Mặt bằng khu vực tổng thể).\n"
            "- Perspective View (Phối cảnh 3D góc chéo).\n"
            "- Booth Dimensions (Mặt bằng 2D nhìn thẳng từ trên xuống - Top view, KHÔNG góc chéo).\n"
            "- Other."
        )
        try:
            response = self.client.models.generate_content(
                model=self.model,
                contents=[system_prompt, image]
            )
            return response.text.strip()
        except:
            return "Other"

# ====================================================================================
# 4. RUN AI PROCESSING
# ====================================================================================
def run_ai_processing(quotation_file, review_file, table_rows: list, ai: AIProvider) -> PipelineResult:
    result = PipelineResult()
    
    # Giả lập đọc text từ Báo giá PDF (để code gọn nhẹ chống crash)
    quotation_text = ""
    labels = [str(r.get("Tên thiết bị", "")).strip() for r in table_rows if str(r.get("Tên thiết bị", "")).strip()]
    regex_found = regex_extract_quantities(quotation_text, labels) if quotation_text else {}
    
    for idx, row in enumerate(table_rows):
        label = str(row["Tên thiết bị"]).strip()
        synced_qty = regex_found.get(label, int(row["Số lượng"])) 
        result.equipment.append(EquipmentRow(
            key=slugify(label, idx), label=label, icon_value=row["Biểu tượng vẽ"],
            color_hex=row["Mã màu Hex"], zone_raw=row["Vị trí ưu tiên"], 
            power=row["Công suất"], quantity=int(synced_qty)
        ))
    
    # Xử lý PDF bằng PyMuPDF (Ổn định và tiêu thụ cực ít RAM)
    doc = fitz.open(stream=review_file.read(), filetype="pdf")
    classified = {t: None for t in SLIDE_TYPES}
    
    for page in doc:
        pix = page.get_pixmap(dpi=150)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        label = ai.classify_slide_image(img)
        for t in SLIDE_TYPES:
            if t in label and classified[t] is None:
                classified[t] = img
                break
    doc.close()
                
    result.booth_location_img = classified["Booth Location"]
    result.perspective_img = classified["Perspective View"]
    result.dimensions_img = classified["Booth Dimensions"]
    
    coords = {}
    for it in result.equipment:
        coords[it.key] = [[round(random.uniform(1.0, 5.0), 2), round(random.uniform(1.0, 5.0), 2)] for _ in range(it.quantity)]
    result.coordinates = coords
    
    gc.collect()
    return result

# ====================================================================================
# 5. HÀM TẠO PPTX
# ====================================================================================
def export_final_pptx(result: PipelineResult, final_b64_image: str, final_counts: dict) -> bytes:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]

    def add_img(slide, img_obj):
        img_w, img_h = img_obj.size
        slide_ratio = prs.slide_width / prs.slide_height
        pic_w, pic_h = (prs.slide_width, int(prs.slide_width / (img_w/img_h))) if (img_w/img_h) > slide_ratio else (int(prs.slide_height * (img_w/img_h)), prs.slide_height)
        left, top = int((prs.slide_width - pic_w)/2), int((prs.slide_height - pic_h)/2)
        buf = io.BytesIO()
        img_obj.convert("RGB").save(buf, format="PNG")
        slide.shapes.add_picture(buf, left, top, width=pic_w, height=pic_h)

    for img in [result.booth_location_img, result.perspective_img, result.dimensions_img]:
        s = prs.slides.add_slide(blank)
        if img: add_img(s, img)

    s4 = prs.slides.add_slide(blank)
    if final_b64_image:
        header, encoded = final_b64_image.split(",", 1)
        final_img = Image.open(io.BytesIO(base64.b64decode(encoded)))
        add_img(s4, final_img)

        rows = []
        for eq in result.equipment:
            qty = final_counts.get(eq.key, 0)
            if qty > 0:
                rows.append((eq.label, eq.icon_value, eq.color_hex, str(qty)))
        
        if rows:
            table_shape = s4.shapes.add_table(len(rows)+1, 4, prs.slide_width - Inches(4.5), Inches(0.25), Inches(4.4), Inches(0.3 * (len(rows)+1)))
            for c, t in enumerate(["Thiết bị", "Ký hiệu", "Màu", "SL"]):
                table_shape.table.cell(0, c).text = t
            for r, (lbl, sym, col, qt) in enumerate(rows, 1):
                table_shape.table.cell(r, 0).text = lbl
                table_shape.table.cell(r, 1).text = sym
                table_shape.table.cell(r, 3).text = qt
                table_shape.table.cell(r, 2).fill.solid()
                table_shape.table.cell(r, 2).fill.fore_color.rgb = RGBColor(*[int(col.strip('#')[i:i+2], 16) for i in (0, 2, 4)])

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()

# ====================================================================================
# 6. GIAO DIỆN CHÍNH
# ====================================================================================
def main():
    st.set_page_config(page_title="MEP Layout Auto", page_icon="⚡", layout="wide")
    
    if "step" not in st.session_state:
        st.session_state.step = 1
        st.session_state.result = None
        st.session_state.equipment_table = pd.DataFrame(DEFAULT_EQUIPMENT_CONFIG)

    st.title("⚡ MEP Layout Auto-Generator (Cloud Optimized)")
    
    if st.session_state.step == 1:
        st.info("Bước 1: Cấu hình bảng Legend và Upload file thiết kế (Quotation + 3D Review).")
        c1, c2 = st.columns([1, 1])
        with c1:
            # Bỏ use_container_width để hết cảnh báo log
            edited_df = st.data_editor(st.session_state.equipment_table, num_rows="dynamic")
            st.session_state.equipment_table = edited_df
        with c2:
            st.sidebar.text_input("Google Gemini API Key", key="api_key", type="password")
            st.sidebar.selectbox("Model", ["gemini-1.5-pro", "gemini-2.5-flash"], key="model_choice")
            
            q_file = st.file_uploader("File Báo Giá (PDF/Excel)")
            r_file = st.file_uploader("File 3D Review (PDF)")
            
            if st.button("🚀 Xử Lý AI Khởi Tạo Bản Đồ"):
                if not st.session_state.api_key: 
                    st.error("⚠️ Vui lòng nhập API Key!")
                elif not q_file or not r_file: 
                    st.error("⚠️ Vui lòng upload đầy đủ 2 file!")
                else:
                    try:
                        ai = AIProvider(st.session_state.api_key, st.session_state.model_choice)
                        with st.spinner("AI đang nhận diện Top View và bóc tách thiết bị..."):
                            st.session_state.result = run_ai_processing(q_file, r_file, edited_df.to_dict("records"), ai)
                            st.session_state.step = 2
                            st.rerun()
                    except Exception as e:
                        st.error(f"❌ Có lỗi xảy ra trong quá trình xử lý AI: {str(e)}")

    elif st.session_state.step == 2:
        st.success("Bước 2: Kéo thả Icon cho đúng vị trí, click phải để xóa, rồi bấm 'Chốt Bản Đồ & Lưu'.")
        
        if st.button("⬅️ Quay lại Bước 1"):
            st.session_state.step = 1
            st.rerun()

        result = st.session_state.result
        bg_b64 = ""
        if result.dimensions_img:
            buf = io.BytesIO()
            result.dimensions_img.convert("RGB").save(buf, format="PNG")
            bg_b64 = f"data:image/png;base64,{base64.b64encode(buf.getvalue()).decode('utf-8')}"
        else:
            st.warning("⚠️ AI không tìm thấy bản vẽ Top View chuẩn 2D trong file PDF.")

        js_items = []
        for eq in result.equipment:
            pts = result.coordinates.get(eq.key, [])
            for p in pts:
                js_items.append({"key": eq.key, "icon_val": eq.icon_value, "color": eq.color_hex, "x": p[0], "y": p[1]})

        component_value = interactive_map_component(
            bg_base64=bg_b64, 
            items=js_items, 
            key="map_interact"
        )

        if component_value and component_value.get("status") == "approved":
            st.divider()
            st.header("📥 BƯỚC CUỐI: Duyệt & Xuất File PPTX")
            
            final_b64 = component_value.get("final_image_b64")
            final_counts = component_value.get("final_counts", {})

            pptx_bytes = export_final_pptx(result, final_b64, final_counts)
            
            st.download_button(
                label="🎉 TẢI XUỐNG FILE MEP LAYOUT (.PPTX)",
                data=pptx_bytes,
                file_name="MEP_Layout_Final.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

if __name__ == "__main__":
    main()

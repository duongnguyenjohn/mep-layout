"""
====================================================================================
 HỆ THỐNG TỰ ĐỘNG BÓC TÁCH & VẼ SƠ ĐỒ ĐIỆN GIAN HÀNG (Electrical Layout Auto-Generator)
 PHIÊN BẢN V7.0 PURE PYTHON — Bỏ JS/HTML, Render Tọa độ Trực tiếp bằng Python
====================================================================================
"""

import io
import os
import random
import re
import json
import unicodedata
import gc
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from pypdf import PdfReader
from pdf2image import convert_from_bytes
from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor

try:
    from google import genai
except ImportError:
    genai = None

# ====================================================================================
# 1. CẤU HÌNH MẶC ĐỊNH
# ====================================================================================
AI_MODEL_DEFAULT = "gemini-1.5-pro"
STOPWORDS_VI = {"và", "cho", "của", "tại"}
SLIDE_TYPES = ["Booth Location", "Perspective View", "Booth Dimensions"]
GRID_METERS = 6.0

DEFAULT_EQUIPMENT_CONFIG = [
    {"Tên thiết bị": "Đèn Floodlight 50W", "Biểu tượng vẽ": "▲", "Mã màu Hex": "#FFCD00", "Vị trí ưu tiên": "Hệ trần biên", "Công suất": "50W", "Số lượng": 10},
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

def hex_to_rgb(hex_color: str) -> tuple:
    try:
        h = (hex_color or "").strip().lstrip("#")
        return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))
    except: return (128, 128, 128)

def slugify(text: str, idx: int) -> str:
    try:
        norm = unicodedata.normalize("NFKD", text or "").encode("ascii", "ignore").decode("ascii")
        norm = re.sub(r"[^a-zA-Z0-9]+", "_", norm).strip("_").lower()
        return f"{norm}_{idx}"
    except: return f"item_{idx}"

def read_quotation_as_text(file_bytes: bytes, filename: str) -> str:
    ext = os.path.splitext(filename.lower())[1]
    try:
        if ext == ".pdf":
            reader = PdfReader(io.BytesIO(file_bytes))
            return "\n".join(page.extract_text() or "" for page in reader.pages)
        elif ext in (".xlsx", ".xls"):
            sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)
            return "\n".join(df.to_string(index=False) for df in sheets.values())
        else: return ""
    except: return ""

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
# 3. AI PROVIDER (SMART CROPPING)
# ====================================================================================
class AIProvider:
    def __init__(self, api_key: str, model: str):
        if genai is None:
            raise RuntimeError("Chưa cài đặt google-genai.")
        self.client = genai.Client(api_key=api_key)
        self.model = model

    def analyze_slide_image(self, image: Image.Image) -> dict:
        system_prompt = """Trang PDF này có thể chứa một hoặc nhiều góc nhìn.
Trả về JSON:
{
    "label": "Booth Location" | "Perspective View" | "Booth Dimensions" | "Other",
    "crop_box": [ymin, xmin, ymax, xmax]
}
Quy tắc:
- Tổng mặt bằng khu vực: "Booth Location".
- Phối cảnh 3D góc chéo: "Perspective View".
- Nếu có "Top view" (Mặt bằng 2D nhìn thẳng từ trên xuống có lưới), label là "Booth Dimensions". Khi đó, trả về crop_box là mảng 4 số nguyên 0-1000 để cắt vùng Top View. Nếu không cắt được thì để null."""
        try:
            response = self.client.models.generate_content(model=self.model, contents=[system_prompt, image])
            match = re.search(r'\{.*\}', response.text.strip(), re.DOTALL)
            if match: return json.loads(match.group(0))
            return {"label": "Other", "crop_box": None}
        except: return {"label": "Other", "crop_box": None}

# ====================================================================================
# 4. RUN AI PROCESSING
# ====================================================================================
def run_ai_processing(quotation_file, review_file, table_rows: list, ai: AIProvider) -> PipelineResult:
    result = PipelineResult()
    
    quotation_text = read_quotation_as_text(quotation_file.read(), quotation_file.name)
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
    
    images = convert_from_bytes(review_file.read(), dpi=100)
    classified = {t: None for t in SLIDE_TYPES}
    
    for img in images:
        data = ai.analyze_slide_image(img)
        label = data.get("label", "Other")
        
        if label == "Booth Dimensions" and classified["Booth Dimensions"] is None:
            crop_box = data.get("crop_box")
            if crop_box and len(crop_box) == 4:
                try:
                    ymin, xmin, ymax, xmax = crop_box
                    w, h = img.size
                    left, top = (xmin / 1000.0) * w, (ymin / 1000.0) * h
                    right, bottom = (xmax / 1000.0) * w, (ymax / 1000.0) * h
                    classified["Booth Dimensions"] = img.crop((left, top, right, bottom))
                except: classified["Booth Dimensions"] = img
            else: classified["Booth Dimensions"] = img
                
        elif label == "Booth Location" and classified["Booth Location"] is None:
            classified["Booth Location"] = img
        elif label == "Perspective View" and classified["Perspective View"] is None:
            classified["Perspective View"] = img
            
    if classified["Booth Dimensions"] is None and len(images) > 0:
        classified["Booth Dimensions"] = images[0]

    result.booth_location_img = classified["Booth Location"]
    result.perspective_img = classified["Perspective View"]
    result.dimensions_img = classified["Booth Dimensions"]
    
    del images
    gc.collect()
    return result

# ====================================================================================
# 5. PYTHON NATIVE RENDER (VẼ TRỰC TIẾP LÊN ẢNH BẰNG PILLOW MÀ KHÔNG CẦN JS)
# ====================================================================================
def render_mep_map_python(bg_img: Image.Image, coords_df: pd.DataFrame) -> Image.Image:
    """Vẽ lưới 6x6m và đặt các icon lên ảnh nền bằng thư viện Pillow (Python)."""
    if bg_img is None: return None
    
    img = bg_img.copy().convert("RGBA")
    draw = ImageDraw.Draw(img)
    w, h = img.size

    # Vẽ lưới mờ 6x6m
    for i in range(1, int(GRID_METERS)):
        x = int((i / GRID_METERS) * w)
        y = int((i / GRID_METERS) * h)
        draw.line([(x, 0), (x, h)], fill=(0, 80, 136, 100), width=1)
        draw.line([(0, y), (w, y)], fill=(0, 80, 136, 100), width=1)

    # Thử load Font hệ thống, nếu không có dùng font mặc định
    try:
        font = ImageFont.truetype("arial.ttf", max(14, int(w*0.02)))
    except:
        font = ImageFont.load_default()

    # Duyệt qua bảng tọa độ để vẽ đè Icon
    for _, row in coords_df.iterrows():
        try:
            x_m = float(row["Trục X (m)"])
            y_m = float(row["Trục Y (m)"])
            
            # Chỉ vẽ nếu tọa độ hợp lệ
            if pd.isna(x_m) or pd.isna(y_m): continue
            
            px = int((x_m / GRID_METERS) * w)
            py = int((y_m / GRID_METERS) * h)
            
            color_rgb = hex_to_rgb(str(row["Màu Hex"]))
            icon_char = str(row["Ký Hiệu"])
            
            # Vẽ nền tròn cho Icon
            r = int(w * 0.015)
            draw.ellipse([px-r, py-r, px+r, py+r], fill=color_rgb, outline=(255,255,255), width=2)
            
            # Căn giữa chữ
            bbox = draw.textbbox((0, 0), icon_char, font=font)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
            draw.text((px - tw/2, py - th/2 - 2), icon_char, fill=(0,0,0), font=font)
        except Exception:
            continue

    return img.convert("RGB")

# ====================================================================================
# 6. HÀM TẠO PPTX
# ====================================================================================
def export_final_pptx(result: PipelineResult, final_drawn_img: Image.Image, legend_df: pd.DataFrame) -> bytes:
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

    # Slide 1, 2, 3
    for img in [result.booth_location_img, result.perspective_img, result.dimensions_img]:
        s = prs.slides.add_slide(blank)
        if img: add_img(s, img)

    # Slide 4: Ảnh đã được vẽ bằng Python
    s4 = prs.slides.add_slide(blank)
    if final_drawn_img:
        add_img(s4, final_drawn_img)

        # Tạo bảng Legend
        # Lọc danh sách thiết bị có số lượng > 0
        qty_counts = legend_df["Thiết Bị"].value_counts().to_dict()
        
        rows = []
        for eq in result.equipment:
            qty = qty_counts.get(eq.label, 0)
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
                table_shape.table.cell(r, 2).fill.fore_color.rgb = RGBColor(*hex_to_rgb(col))

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()

# ====================================================================================
# 7. GIAO DIỆN CHÍNH (NATIVE STREAMLIT WORKFLOW)
# ====================================================================================
def main():
    st.set_page_config(page_title="MEP Layout Auto V7.0", page_icon="⚡", layout="wide")
    
    if "step" not in st.session_state:
        st.session_state.step = 1
        st.session_state.result = None
        st.session_state.equipment_table = pd.DataFrame(DEFAULT_EQUIPMENT_CONFIG)
        st.session_state.coords_df = pd.DataFrame()

    st.title("⚡ MEP Layout Auto-Generator V7.0 (Python Native 100%)")

    # ---------------- BƯỚC 1 ----------------
    if st.session_state.step == 1:
        st.info("Bước 1: Cấu hình thiết bị và Upload file (Bản này KHÔNG sử dụng Javascript, 100% không lỗi Cloud).")
        c1, c2 = st.columns([1, 1])
        with c1:
            edited_df = st.data_editor(st.session_state.equipment_table, num_rows="dynamic", use_container_width=True)
            st.session_state.equipment_table = edited_df
        with c2:
            st.sidebar.text_input("Google Gemini API Key", key="api_key", type="password")
            st.sidebar.selectbox("Model", ["gemini-1.5-pro", "gemini-2.5-flash"], key="model_choice")
            
            q_file = st.file_uploader("File Báo Giá (PDF/Excel)")
            r_file = st.file_uploader("File 3D Review (PDF)")
            
            if st.button("🚀 Xử Lý Khởi Tạo Sơ Đồ"):
                if not st.session_state.api_key: st.error("⚠️ Vui lòng nhập API Key!")
                elif not q_file or not r_file: st.error("⚠️ Vui lòng upload đầy đủ 2 file!")
                else:
                    try:
                        ai = AIProvider(st.session_state.api_key, st.session_state.model_choice)
                        with st.spinner("AI đang bóc tách thiết bị và xử lý ảnh..."):
                            res = run_ai_processing(q_file, r_file, edited_df.to_dict("records"), ai)
                            st.session_state.result = res
                            
                            # Khởi tạo bảng danh sách từng cá thể thiết bị
                            initial_coords = []
                            for eq in res.equipment:
                                for i in range(eq.quantity):
                                    initial_coords.append({
                                        "Thiết Bị": eq.label,
                                        "Ký Hiệu": eq.icon_value,
                                        "Màu Hex": eq.color_hex,
                                        "Trục X (m)": round(random.uniform(1.0, 5.0), 1),
                                        "Trục Y (m)": round(random.uniform(1.0, 5.0), 1)
                                    })
                            st.session_state.coords_df = pd.DataFrame(initial_coords)
                            st.session_state.step = 2
                            st.rerun()
                    except Exception as e:
                        st.error(f"❌ Có lỗi xảy ra: {str(e)}")

    # ---------------- BƯỚC 2 ----------------
    elif st.session_state.step == 2:
        st.success("Bước 2: Thay đổi tọa độ X, Y (0m - 6m) hoặc thêm/xóa thiết bị trong bảng bên phải. Bản đồ sẽ tự cập nhật ngay lập tức!")
        
        if st.button("⬅️ Quay lại Bước 1"):
            st.session_state.step = 1
            st.rerun()

        res = st.session_state.result
        if not res.dimensions_img:
            st.warning("⚠️ Không tìm thấy ảnh Top View.")
            st.stop()

        col_img, col_data = st.columns([1.5, 1])
        
        with col_data:
            st.markdown("### 🛠 Bảng Tọa Độ (X, Y)")
            st.caption("Mẹo: Nhấn vào cột 'Trục X (m)' hoặc 'Trục Y (m)' để gõ số mới. Bạn cũng có thể thêm hàng trống bên dưới cùng để tăng số lượng.")
            
            # Data Editor thay thế hoàn toàn việc kéo thả JS
            new_coords_df = st.data_editor(
                st.session_state.coords_df,
                num_rows="dynamic",
                use_container_width=True,
                height=600,
                column_config={
                    "Trục X (m)": st.column_config.NumberColumn(min_value=0.0, max_value=6.0, step=0.1),
                    "Trục Y (m)": st.column_config.NumberColumn(min_value=0.0, max_value=6.0, step=0.1),
                }
            )
            # Cập nhật State
            st.session_state.coords_df = new_coords_df

        with col_img:
            st.markdown("### 🗺️ Bản Đồ Xem Trước")
            # Python vẽ lại hình ảnh dựa trên dữ liệu từ Data Editor
            final_drawn_map = render_mep_map_python(res.dimensions_img, new_coords_df)
            st.image(final_drawn_map, use_container_width=True)

        st.divider()
        st.header("📥 BƯỚC CUỐI: Chốt Bản Vẽ & Xuất File")
        st.info("Bảng Legend trên Slide 4 sẽ tự động tổng hợp đếm số lượng dựa trên những thiết bị hiện có trong Bảng Tọa Độ ở trên.")
        
        pptx_bytes = export_final_pptx(res, final_drawn_map, new_coords_df)
        st.download_button(
            label="🎉 TẢI XUỐNG FILE PPTX HOÀN CHỈNH",
            data=pptx_bytes,
            file_name="MEP_Layout_Final_V7.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary"
        )

if __name__ == "__main__":
    main()
 

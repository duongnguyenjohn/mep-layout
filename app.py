"""
====================================================================================
 HỆ THỐNG TỰ ĐỘNG BÓC TÁCH & VẼ SƠ ĐỒ ĐIỆN GIAN HÀNG (Electrical Layout Auto-Generator)
 PHIÊN BẢN V3 — Flexible Legend Editor + Dynamic Vision Analysis
====================================================================================

Streamlit end-to-end: Bảng Legend đa phương thức (chữ / emoji / ảnh) + Upload 2 file
--> Submit --> 4 Round xử lý ngầm --> Download "Electrical Layout.pptx" (đúng 4 slide).

Round 1 : Đọc & khớp số lượng thiết bị từ file Quotation (PDF/Excel) với Bảng cấu hình
          (bảng KHÔNG còn danh mục cố định — người dùng có thể thêm/xoá dòng tuỳ dự án).
Round 2 : Dynamic Vision Analysis — quét TOÀN BỘ file 3D review (số trang không cố định),
          dùng Vision LLM (Claude hoặc GPT-4o) tìm Booth Location / Perspective View /
          Booth Dimensions (Top View), rồi đối chiếu chéo (Cross-view) để suy ra toạ độ
          kỹ thuật (X, Y) trên lưới sàn 0-6m.
Round 3 : Vẽ đè ký hiệu LINH HOẠT (Pillow) — tự nhận diện biểu tượng người dùng nhập là
          chữ/ký tự/emoji (draw.text) hay đường dẫn ảnh/URL (Image.open + paste), kèm
          thuật toán chống chồng chất (Jittering Offset 5-12 pixel).
Round 4 : Đóng gói đúng 4 slide: Booth Location / Perspective View / Booth Dimensions
          (sơ bộ) / MEP Layout (chi tiết, dày ký hiệu + bảng Legend góc phải).

--------------------------------------------------------------------------------------
CẤU HÌNH macOS & POPPLER (FIX LỖI CRASH):
pdf2image cần Poppler. App tự động dò `poppler_path` theo dòng chip Mac và truyền vào
hàm convert_from_bytes() để tránh hoàn toàn lỗi PopplerNotInstalledError:
  - Apple Silicon (arm64) : /opt/homebrew/bin
  - Intel (x86_64)        : /usr/local/bin
Cài Poppler trước khi chạy: `brew install poppler`
--------------------------------------------------------------------------------------

Cấu hình API Key: nhập ở sidebar (chỉ dùng trong phiên) hoặc set biến môi trường
ANTHROPIC_API_KEY / OPENAI_API_KEY trước khi chạy `streamlit run app.py`.

Cài đặt thư viện Python:
    pip install streamlit pypdf pdf2image python-pptx Pillow pandas openpyxl anthropic openai requests
"""
"""
====================================================================================
 HỆ THỐNG TỰ ĐỘNG BÓC TÁCH & VẼ SƠ ĐỒ ĐIỆN GIAN HÀNG (Electrical Layout Auto-Generator)
 PHIÊN BẢN V5.1 FINAL — Interactive Workflow + Drag & Drop Canvas + AI Data Extraction
====================================================================================
"""
"""
====================================================================================
 HỆ THỐNG TỰ ĐỘNG BÓC TÁCH & VẼ SƠ ĐỒ ĐIỆN GIAN HÀNG (Electrical Layout Auto-Generator)
 PHIÊN BẢN V5.2 FINAL — Interactive Workflow + Gemini Pro Vision + Error Handling
====================================================================================
"""

import io
import json
import math
import os
import platform
import random
import re
import unicodedata
from dataclasses import dataclass, field
from typing import Optional
import base64
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image
from pypdf import PdfReader
import fitz
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor

# Tích hợp Google Gemini
try:
    import google.generativeai as genai
except ImportError:
    genai = None

# ====================================================================================
# 0. KHỞI TẠO CUSTOM COMPONENT GIAO TIẾP 2 CHIỀU (STREAMLIT <-> JS)
# ====================================================================================
COMPONENT_DIR = os.path.join(os.path.dirname(__file__), "map_component")
os.makedirs(COMPONENT_DIR, exist_ok=True)

COMPONENT_HTML = """
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://html2canvas.hertzen.com/dist/html2canvas.min.js"></script>
    <script src="https://unpkg.com/streamlit-component-lib@1.3.0/dist/streamlit.js"></script>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 10px; background: transparent; }
        .map-container { position: relative; width: 800px; height: 800px; background-size: 100% 100%; border: 2px solid #005088; margin: 0 auto; overflow: hidden; background-color: #f0f0f0; }
        .grid-overlay { position: absolute; top: 0; left: 0; right: 0; bottom: 0; background-image: linear-gradient(rgba(0, 80, 136, 0.4) 1px, transparent 1px), linear-gradient(90deg, rgba(0, 80, 136, 0.4) 1px, transparent 1px); background-size: 133.33px 133.33px; pointer-events: none; z-index: 1; }
        .icon-node { position: absolute; z-index: 10; cursor: grab; transform: translate(-50%, -50%); display: flex; align-items: center; justify-content: center; width: 30px; height: 30px; border-radius: 4px; color: white; font-size: 14px; font-weight: bold; box-shadow: 0 2px 5px rgba(0,0,0,0.5); user-select: none; }
        .icon-node:active { cursor: grabbing; z-index: 999; }
        .toolbar { width: 800px; margin: 10px auto; padding: 10px; background: #fff; border: 1px solid #ccc; border-radius: 8px; display: flex; gap: 10px; align-items: center; justify-content: space-between; }
        .btn { padding: 8px 16px; background: #005088; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; }
        .btn:hover { background: #003057; }
        .btn-success { background: #10b981; }
        .btn-success:hover { background: #059669; }
        .delete-hint { font-size: 12px; color: #666; }
    </style>
</head>
<body>
    <div class="toolbar">
        <div><span class="delete-hint">💡 Kéo thả để di chuyển. Click chuột phải vào Icon để XÓA.</span></div>
        <div><button class="btn btn-success" id="btn-save" onclick="saveAndSend()">📸 Chốt Bản Đồ & Lưu</button></div>
    </div>
    <div class="map-container" id="map-area">
        <div class="grid-overlay"></div>
    </div>

    <script>
        let isDragging = false;
        let currentDrag = null;
        let itemsData = [];

        function initComponent() { Streamlit.setFrameHeight(900); }

        function onDataFromPython(event) {
            if (event.data.type !== "streamlit:render") return;
            const args = event.data.args;
            const mapArea = document.getElementById("map-area");
            if (args.bg_base64) mapArea.style.backgroundImage = `url(${args.bg_base64})`;
            if (itemsData.length === 0 && args.items) {
                itemsData = args.items;
                renderItems();
            }
        }

        function renderItems() {
            const mapArea = document.getElementById("map-area");
            mapArea.querySelectorAll('.icon-node').forEach(e => e.remove());
            itemsData.forEach((item, index) => {
                const node = document.createElement("div");
                node.className = "icon-node";
                node.innerHTML = item.icon_val;
                node.style.backgroundColor = item.color;
                node.style.left = `${(item.x / 6) * 100}%`;
                node.style.top = `${(item.y / 6) * 100}%`;
                node.dataset.index = index;

                node.addEventListener("mousedown", (e) => { if (e.button === 0) { isDragging = true; currentDrag = node; } });
                node.addEventListener("contextmenu", (e) => { e.preventDefault(); if(confirm("Xóa thiết bị này?")) node.remove(); });
                mapArea.appendChild(node);
            });
        }

        document.addEventListener("mousemove", (e) => {
            if (isDragging && currentDrag) {
                const rect = document.getElementById("map-area").getBoundingClientRect();
                let x = Math.max(0, Math.min(e.clientX - rect.left, rect.width));
                let y = Math.max(0, Math.min(e.clientY - rect.top, rect.height));
                currentDrag.style.left = `${(x / rect.width) * 100}%`;
                currentDrag.style.top = `${(y / rect.height) * 100}%`;
            }
        });

        document.addEventListener("mouseup", () => { isDragging = false; currentDrag = null; });

        function saveAndSend() {
            const mapArea = document.getElementById("map-area");
            const btn = document.getElementById("btn-save");
            btn.innerText = "⏳ Đang chụp màn hình...";
            
            let finalCounts = {};
            document.querySelectorAll('.icon-node').forEach(node => {
                let key = itemsData[node.dataset.index].key;
                finalCounts[key] = (finalCounts[key] || 0) + 1;
            });

            html2canvas(mapArea, { useCORS: true, scale: 2 }).then(canvas => {
                Streamlit.setComponentValue({ status: "approved", final_image_b64: canvas.toDataURL("image/png"), final_counts: finalCounts });
                btn.innerText = "✅ Đã gửi! Vui lòng bấm Xuất File.";
                btn.classList.remove('btn-success');
            });
        }

        Streamlit.events.addEventListener(Streamlit.RENDER_EVENT, onDataFromPython);
        Streamlit.setComponentReady();
        initComponent();
    </script>
</body>
</html>
"""
with open(os.path.join(COMPONENT_DIR, "index.html"), "w", encoding="utf-8") as f:
    f.write(COMPONENT_HTML)

interactive_map_component = components.declare_component("interactive_map", path=COMPONENT_DIR)


# ====================================================================================
# 1. CẤU HÌNH MẶC ĐỊNH
# ====================================================================================
AI_MODEL_DEFAULT = "gemini-1.5-pro" # Mặc định chuyển sang dùng Gemini Pro
GRID_METERS = 6.0
STOPWORDS_VI = {"và", "cho", "của", "tại"}
SLIDE_TYPES = ["Booth Location", "Perspective View", "Booth Dimensions"]

DEFAULT_EQUIPMENT_CONFIG = [
    {"Tên thiết bị": "Đèn Floodlight 50W", "Biểu tượng vẽ": "▲", "Mã màu Hex": "#FFCD00", "Vị trí ưu tiên": "Hệ trần biên", "Công suất": "50W", "Số lượng": 30},
    {"Tên thiết bị": "Ổ cắm 5A/220V", "Biểu tượng vẽ": "●", "Mã màu Hex": "#D62728", "Vị trí ưu tiên": "Bàn tư vấn", "Công suất": "220V", "Số lượng": 3},
]


# ====================================================================================
# 2. DATA STRUCTURES & HELPER LÕI
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
    log: list = field(default_factory=list)
    errors: list = field(default_factory=list)

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
            sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, engine="openpyxl" if ext == ".xlsx" else None)
            return "\n".join(df.to_string(index=False) for df in sheets.values())
        elif ext == ".csv":
            return pd.read_csv(io.BytesIO(file_bytes)).to_string(index=False)
        else: return ""
    except: return ""

def extract_salient_tokens(label: str) -> list:
    tokens = re.split(r"[\s/()\-,]+", (label or "").lower())
    return [t for t in tokens if len(t) >= 2 and t not in STOPWORDS_VI]

def regex_extract_quantities_generic(text: str, labels: list) -> dict:
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
# 3. AI PROVIDER (TÍCH HỢP GOOGLE GEMINI PRO)
# ====================================================================================
class AIProvider:
    def __init__(self, api_key: str, model: str):
        if genai is None:
            raise RuntimeError("Chưa cài thư viện. Vui lòng chạy: pip install google-generativeai")
        
        # Cấu hình API Key cho Google Gemini
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel(model)

    def classify_slide_image(self, image: Image.Image) -> str:
        system_prompt = (
            "Phân loại hình ảnh trang PDF này vào ĐÚNG MỘT nhãn sau (chỉ trả về tên nhãn):\n"
            "- Booth Location (Mặt bằng khu vực tổng thể toàn hội chợ).\n"
            "- Perspective View (Phối cảnh 3D góc chéo, có chiều sâu 3D, thấy tường, trần, bàn ghế nổi lên).\n"
            "- Booth Dimensions (BẮT BUỘC là mặt bằng 2D nhìn thẳng góc 90 độ từ trên xuống - Top view. Thể hiện không gian phẳng 2D, thường có đường kích thước hoặc lưới. TUYỆT ĐỐI KHÔNG CHỌN nhãn này nếu hình ảnh có góc chéo 3D).\n"
            "- Other."
        )
        
        # Gemini nhận diện trực tiếp ảnh PIL rất tiện lợi
        response = self.model.generate_content([system_prompt, image])
        return response.text.strip()


# ====================================================================================
# 4. RUN AI PROCESSING (QUY TRÌNH BÓC TÁCH & PHÂN LOẠI ẢNH)
# ====================================================================================
def run_ai_processing(quotation_file, review_file, table_rows: list, ai: AIProvider) -> PipelineResult:
    result = PipelineResult()
    
    # 1. Round 1: Đọc & Bóc tách file Báo giá (Quotation)
    quotation_text = read_quotation_as_text(quotation_file.read(), quotation_file.name)
    labels = [str(r.get("Tên thiết bị", "")).strip() for r in table_rows if str(r.get("Tên thiết bị", "")).strip()]
    regex_found = regex_extract_quantities_generic(quotation_text, labels) if quotation_text else {}
    
    for idx, row in enumerate(table_rows):
        label = str(row["Tên thiết bị"]).strip()
        default_qty = int(row["Số lượng"])
        synced_qty = regex_found.get(label, default_qty) 
        
        result.equipment.append(EquipmentRow(
            key=slugify(label, idx), label=label, icon_value=row["Biểu tượng vẽ"],
            color_hex=row["Mã màu Hex"], zone_raw=row["Vị trí ưu tiên"], 
            power=row["Công suất"], quantity=int(synced_qty)
        ))
    
    # 2. Xử lý PDF lấy ảnh (Round 2)
    doc = fitz.open(stream=review_file.read(), filetype="pdf")
    classified = {t: None for t in SLIDE_TYPES}
    for page in doc:
        img = Image.open(io.BytesIO(page.get_pixmap(dpi=150).tobytes("png")))
        label = ai.classify_slide_image(img)
        for t in SLIDE_TYPES:
            if t in label and classified[t] is None:
                classified[t] = img
                break
    
    result.booth_location_img = classified["Booth Location"]
    result.perspective_img = classified["Perspective View"]
    result.dimensions_img = classified["Booth Dimensions"]
    
    # 3. Rải đều toạ độ để người dùng kéo thả trên HTML Canvas
    coords = {}
    for it in result.equipment:
        coords[it.key] = [[round(random.uniform(1.0, 5.0), 2), round(random.uniform(1.0, 5.0), 2)] for _ in range(it.quantity)]
    result.coordinates = coords
    
    return result


# ====================================================================================
# 5. HÀM TẠO PPTX (SỬ DỤNG HÌNH ẢNH BASE64 TỪ COMPONENT JS)
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

    # Slide 1, 2, 3 giữ nguyên bản gốc sạch
    for img in [result.booth_location_img, result.perspective_img, result.dimensions_img]:
        s = prs.slides.add_slide(blank)
        if img: add_img(s, img)

    # Slide 4: CHỤP TỪ MÀN HÌNH TƯƠNG TÁC HTML
    s4 = prs.slides.add_slide(blank)
    if final_b64_image:
        header, encoded = final_b64_image.split(",", 1)
        final_img = Image.open(io.BytesIO(base64.b64decode(encoded)))
        add_img(s4, final_img)

        # Tạo bảng Legend mới dựa trên tổng lượng thực tế (final_counts)
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
# 6. GIAO DIỆN CHÍNH (WORKFLOW 2 BƯỚC)
# ====================================================================================
def main():
    st.set_page_config(page_title="MEP Layout Auto V5.2", page_icon="⚡", layout="wide")
    
    if "step" not in st.session_state:
        st.session_state.step = 1
        st.session_state.result = None
        st.session_state.equipment_table = pd.DataFrame(DEFAULT_EQUIPMENT_CONFIG)

    st.title("⚡ MEP Layout Auto-Generator V5.2 — Tương Tác Kéo Thả (Gemini Pro)")
    
    # ------------------ GIAI ĐOẠN 1: NHẬP DỮ LIỆU & AI XỬ LÝ ------------------
    if st.session_state.step == 1:
        st.info("Bước 1: Cấu hình bảng Legend và Upload file thiết kế (Quotation + 3D Review).")
        c1, c2 = st.columns([1, 1])
        with c1:
            edited_df = st.data_editor(st.session_state.equipment_table, num_rows="dynamic", use_container_width=True)
            st.session_state.equipment_table = edited_df
        with c2:
            st.sidebar.text_input("Google Gemini API Key (bắt buộc)", key="api_key", type="password")
            st.sidebar.selectbox("Model", ["gemini-1.5-pro", "gemini-1.5-flash"], key="model_choice")
            
            q_file = st.file_uploader("File Báo Giá (PDF/Excel)")
            r_file = st.file_uploader("File 3D Review (PDF)")
            
            if st.button("🚀 Xử Lý AI Khởi Tạo Bản Đồ", type="primary", use_container_width=True):
                if not st.session_state.api_key: 
                    st.error("⚠️ Vui lòng nhập API Key của Google Gemini ở Sidebar!")
                elif not q_file or not r_file: 
                    st.error("⚠️ Vui lòng upload đầy đủ 2 file!")
                else:
                    try:
                        # Khởi tạo Gemini AI
                        ai = AIProvider(st.session_state.api_key, st.session_state.model_choice)
                        
                        with st.spinner("AI đang nhận diện Top View và bóc tách thiết bị từ Báo giá..."):
                            st.session_state.result = run_ai_processing(q_file, r_file, edited_df.to_dict("records"), ai)
                            st.session_state.step = 2
                            st.rerun()
                            
                    except Exception as e:
                        # BẪY LỖI TOÀN DIỆN TRÁNH CRASH APP
                        st.error(f"❌ Có lỗi xảy ra trong quá trình xử lý AI: {str(e)}")
                        st.info("💡 Lời khuyên: Hãy kiểm tra lại API Key xem đã chính xác chưa, hoặc file PDF có bị hỏng không.")

    # ------------------ GIAI ĐOẠN 2: BẢN ĐỒ TƯƠNG TÁC HTML & XUẤT FILE ------------------
    elif st.session_state.step == 2:
        st.success("Bước 2: Hệ thống đã kết nối dữ liệu! Hãy kéo thả Icon cho đúng vị trí, click phải để xóa nếu thừa, rồi bấm 'Chốt Bản Đồ & Lưu'.")
        
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

        # Chuẩn bị dữ liệu cho JS Component
        js_items = []
        for eq in result.equipment:
            pts = result.coordinates.get(eq.key, [])
            for p in pts:
                js_items.append({"key": eq.key, "icon_val": eq.icon_value, "color": eq.color_hex, "x": p[0], "y": p[1]})

        # Gọi Component
        component_value = interactive_map_component(
            bg_base64=bg_b64, 
            items=js_items, 
            key="map_interact"
        )

        # Nhận dữ liệu Component trả về -> XUẤT PPTX
        if component_value and component_value.get("status") == "approved":
            st.divider()
            st.header("📥 BƯỚC CUỐI: Duyệt & Xuất File PPTX")
            st.info("Màn hình kéo thả đã được ghi nhận thành công! Bảng Legend sẽ tự động điều chỉnh theo số lượng bạn đã chốt.")
            
            final_b64 = component_value.get("final_image_b64")
            final_counts = component_value.get("final_counts", {})

            pptx_bytes = export_final_pptx(result, final_b64, final_counts)
            
            st.download_button(
                label="🎉 TẢI XUỐNG FILE MEP LAYOUT (.PPTX)",
                data=pptx_bytes,
                file_name="MEP_Layout_Final.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
                use_container_width=True
            )

if __name__ == "__main__":
    main()
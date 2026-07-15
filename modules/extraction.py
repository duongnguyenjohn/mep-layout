import pdfplumber
from anthropic import Anthropic
import json
import os

def extract_and_align_data(quote_file, review_3d_file):
    """
    Đọc text từ file Báo giá & 3D, sau đó đẩy qua Claude 3.5 Sonnet 
    để tạo JSON ánh xạ tọa độ (Cross-View Triangulation).
    """
    quote_text = ""
    # Trích xuất dữ liệu chữ từ file Báo giá PDF
    with pdfplumber.open(quote_file) as pdf:
        for page in pdf.pages:
            quote_text += page.extract_text() or ""
            
    review_text = ""
    # Trích xuất dữ liệu chữ/thông số kỹ thuật từ file bản vẽ 3D PDF
    with pdfplumber.open(review_3d_file) as pdf:
        for page in pdf.pages:
            review_text += page.extract_text() or ""

    # Lấy API Key từ cấu hình hệ thống
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise ValueError("Không tìm thấy ANTHROPIC_API_KEY trong cấu hình hệ thống!")

    client = Anthropic(api_key=api_key)
    
    # Prompt Kỹ thuật Ép Claude xuất định dạng JSON cấu trúc cao
    prompt = f"""
    Bạn là một kỹ sư hệ thống MEP lão luyện. Nhiệm vụ của bạn là phân tích đồng thời tài liệu Báo giá và Thiết kế 3D dưới đây để hạ cánh (Triangulation) thiết bị xuống lưới tọa độ mặt bằng 2D (Lưới chuẩn 6x6 mét).

    DỮ LIỆU BÁO GIÁ (QUOTATION):
    {quote_text}

    DỮ LIỆU THIẾT KẾ 3D (3D REVIEW):
    {review_text}

    YÊU CẦU:
    1. Đọc và đếm chính xác số lượng thiết bị điện từ Báo giá.
    2. Khớp các thiết bị đó với các khu vực chức năng được mô tả trong bản vẽ 3D (vách kỹ thuật, lối đi, quầy tiếp tân...).
    3. Trả về kết quả CHỈ là một chuỗi JSON hợp lệ (Không chứa mã markdown ```json, không giải thích dài dòng).

    CẤU TRÚC JSON BẮT BUỘC TRẢ VỀ:
    [
      {{
        "item_name": "Tên thiết bị điện chính xác",
        "quantity": 1,
        "x": 4.5,
        "y": 3.2,
        "assigned_layer": "LAYER_MEP_DEVICES",
        "icon": "default_icon.png"
      }}
    ]
    Lưu ý: Nếu một thiết bị có số lượng > 1, hãy tạo ra các phần tử riêng biệt trong mảng với các tọa độ X, Y dịch chuyển nhẹ để phân bổ đủ số lượng thực tế.
    """

    response = client.messages.create(
        model="claude-3-5-sonnet-20241022",
        max_tokens=4000,
        temperature=0,
        messages=[{"role": "user", "content": prompt}]
    )
    
    # Ép kiểu dữ liệu phản hồi về dạng List/Dict Python
    try:
        raw_json = response.content[0].text.strip()
        return json.loads(raw_json)
    except Exception as e:
        # Cơ chế dự phòng (Fallback) nếu AI trả về chuỗi lỗi
        return [
            {"item_name": "Nguồn 20A Máy CNC (Lỗi Phân Tích AI)", "quantity": 1, "x": 3.0, "y": 3.0, "assigned_layer": "LAYER_MEP_DEVICES", "icon": "power.png"}
        ]

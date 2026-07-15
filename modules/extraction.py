import pdfplumber
from google import genai
import json
import os

def extract_and_align_data(quote_file, review_3d_file):
    """
    Sử dụng Gemini qua SDK mới (google-genai) để đọc text từ Báo giá & 3D, 
    xuất JSON ánh xạ tọa độ.
    """
    # 1. Trích xuất text từ Báo giá
    quote_text = ""
    with pdfplumber.open(quote_file) as pdf:
        for page in pdf.pages:
            quote_text += page.extract_text() or ""
            
    # 2. Trích xuất text từ Thiết kế 3D
    review_text = ""
    with pdfplumber.open(review_3d_file) as pdf:
        for page in pdf.pages:
            review_text += page.extract_text() or ""

    # 3. Kết nối API Gemini bằng Client mới
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("Không tìm thấy GEMINI_API_KEY trong cấu hình hệ thống!")

    # Khởi tạo Client theo chuẩn thư viện google-genai
    client = genai.Client(api_key=api_key)
    
    prompt = f"""
    Bạn là một kỹ sư hệ thống MEP. Hãy phân tích tài liệu Báo giá và Thiết kế 3D dưới đây để hạ cánh thiết bị xuống lưới tọa độ 2D (Lưới 6x6 mét).

    BÁO GIÁ:
    {quote_text}

    THIẾT KẾ 3D:
    {review_text}

    YÊU CẦU:
    Trả về CHỈ một chuỗi JSON hợp lệ cấu trúc như sau, KHÔNG CÓ BẤT KỲ VĂN BẢN NÀO KHÁC BÊN NGOÀI:
    [
      {{
        "item_name": "Tên thiết bị",
        "quantity": 1,
        "x": 4.5,
        "y": 3.2,
        "assigned_layer": "LAYER_MEP_DEVICES",
        "icon": "default.png"
      }}
    ]
    Lưu ý: Nếu một thiết bị có số lượng > 1, hãy tạo ra các phần tử riêng biệt trong mảng với tọa độ X, Y khác nhau.
    """

    # 4. Gửi yêu cầu và nhận kết quả bằng mô hình mới nhất
    response = client.models.generate_content(
        model='gemini-2.5-flash',
        contents=prompt,
    )
    
    raw_text = response.text.strip()
    
    # Xử lý làm sạch chuỗi (loại bỏ markdown block nếu có)
    if raw_text.startswith("```json"):
        raw_text = raw_text[7:]
    elif raw_text.startswith("```"):
        raw_text = raw_text[3:]
        
    if raw_text.endswith("```"):
        raw_text = raw_text[:-3]

    try:
        return json.loads(raw_text.strip())
    except Exception as e:
        # Dữ liệu dự phòng nếu xử lý text có lỗi
        return [
            {"item_name": f"Lỗi đọc JSON: {str(e)}", "quantity": 1, "x": 3.0, "y": 3.0, "assigned_layer": "LAYER_MEP_DEVICES", "icon": "error.png"}
        ]

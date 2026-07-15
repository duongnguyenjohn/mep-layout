import pdfplumber
import pandas as pd
from google import genai
import json
import os

def extract_and_align_data(quote_file, review_3d_file):
    # 1. Trích xuất text từ Báo giá (Hỗ trợ cả PDF và Excel)
    quote_text = ""
    file_extension = quote_file.name.split('.')[-1].lower()
    
    if file_extension in ['xlsx', 'xls']:
        df = pd.read_excel(quote_file)
        quote_text = df.to_string() # Chuyển bảng Excel thành text cho AI đọc
    else:
        with pdfplumber.open(quote_file) as pdf:
            for page in pdf.pages:
                quote_text += page.extract_text() or ""
            
    # 2. Trích xuất text từ Thiết kế 3D (PDF)
    review_text = ""
    with pdfplumber.open(review_3d_file) as pdf:
        for page in pdf.pages:
            review_text += page.extract_text() or ""

    api_key = os.environ.get("GEMINI_API_KEY")
    client = genai.Client(api_key=api_key)
    
    # Prompt yêu cầu AI chỉ trả tọa độ X, Y dương
    prompt = f"""
    Bạn là kỹ sư MEP. Phân tích Báo giá và Thiết kế 3D dưới đây để hạ cánh thiết bị xuống mặt bằng.
    BÁO GIÁ: {quote_text}
    THIẾT KẾ 3D: {review_text}

    YÊU CẦU QUAN TRỌNG:
    1. Tọa độ X và Y BẮT BUỘC PHẢI LÀ SỐ DƯƠNG (>0) (Góc phần tư thứ 1).
    2. Trả về CHỈ một chuỗi JSON hợp lệ, KHÔNG CÓ BẤT KỲ VĂN BẢN NÀO KHÁC BÊN NGOÀI:
    [
      {{
        "item_name": "Tên thiết bị",
        "quantity": 1,
        "x": 2.5,
        "y": 4.2,
        "icon": "Tên icon gợi ý"
      }}
    ]
    """

    response = client.models.generate_content(model='gemini-3.5-flash', contents=prompt)
    raw_text = response.text.strip()
    
    if raw_text.startswith("```json"): raw_text = raw_text[7:]
    elif raw_text.startswith("```"): raw_text = raw_text[3:]
    if raw_text.endswith("```"): raw_text = raw_text[:-3]

    try:
        return json.loads(raw_text.strip())
    except Exception as e:
        return [{"item_name": f"Lỗi đọc JSON: {str(e)}", "quantity": 1, "x": 1.0, "y": 1.0}]

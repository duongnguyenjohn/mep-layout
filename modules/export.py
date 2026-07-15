import io
from pptx import Presentation
from pptx.util import Inches, Pt

def generate_pptx_file(mapping_data):
    prs = Presentation()
    prs.slide_width = Inches(13.333) 
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6] 
    
    s1 = prs.slides.add_slide(blank_layout)
    title_box = s1.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(0.8))
    title_box.text_frame.text = "ELECTRICAL LAYOUT - BẢN VẼ BỐ TRÍ THIẾT BỊ"
    
    # Tạo bảng Legend (Ghi chú) ở góc slide
    rows = len(mapping_data) + 1
    table_shape = s1.shapes.add_table(rows, 3, Inches(1), Inches(1.5), Inches(8), Inches(0.4 * rows))
    table = table_shape.table
    
    table.cell(0, 0).text = "Tên Thiết Bị"
    table.cell(0, 1).text = "Số lượng"
    table.cell(0, 2).text = "Tọa độ chốt (X, Y)"
    
    for idx, item in enumerate(mapping_data):
        table.cell(idx + 1, 0).text = str(item.get('item_name', ''))
        table.cell(idx + 1, 1).text = str(item.get('quantity', 1))
        
        # Lấy tọa độ X, Y (Xử lý an toàn nếu AI trả về lỗi)
        try:
            x_val = round(float(item.get('x', 0)), 1)
            y_val = round(float(item.get('y', 0)), 1)
            table.cell(idx + 1, 2).text = f"X: {x_val}, Y: {y_val}"
        except:
            table.cell(idx + 1, 2).text = f"X: 0, Y: 0"
        
    stream = io.BytesIO()
    prs.save(stream)
    return stream.getvalue()

# File: modules/positioning.py
import math

def apply_jittering(device_list, safe_distance=0.2):
    """
    Hàm chống đè chồng icon (Jittering Algorithm) dựa trên khoảng cách Euclid.
    safe_distance: khoảng cách an toàn tối thiểu (tính theo hệ lưới, ví dụ 0.2 mét)
    """
    if not device_list:
        return []
        
    # So sánh từng cặp tọa độ
    for i in range(len(device_list)):
        for j in range(i + 1, len(device_list)):
            dev1 = device_list[i]
            dev2 = device_list[j]
            
            # Tính khoảng cách Euclid
            dx = float(dev1.get('x', 0)) - float(dev2.get('x', 0))
            dy = float(dev1.get('y', 0)) - float(dev2.get('y', 0))
            distance = math.sqrt(dx**2 + dy**2)
            
            # Nếu khoảng cách nhỏ hơn kích thước an toàn, tịnh tiến thiết bị thứ 2
            if distance < safe_distance:
                dev2['x'] = float(dev2.get('x', 0)) + safe_distance
                dev2['y'] = float(dev2.get('y', 0)) + (safe_distance / 2)
                
    return device_list

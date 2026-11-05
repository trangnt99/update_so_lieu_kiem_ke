import streamlit as st
from PIL import Image
import google.generativeai as genai
import io
import json
import pandas as pd

st.subheader("📸 Đọc phiếu kiểm kê tối ưu (Cắt 2 cột & Đọc đa ảnh)")

# Cho phép tải lên từ 1 đến nhiều ảnh phiếu kiểm kê
uploaded_files = st.file_uploader(
    "Tải lên các ảnh phiếu kiểm kê giấy:", 
    type=["jpg", "jpeg", "png"], 
    accept_multiple_files=True
)

# Tùy chọn tọa độ cắt (tính theo phần trăm 0.0 đến 1.0 từ trái qua phải)
# Ví dụ: chỉ lấy vùng giữa hoặc bên phải chứa cột Mã vật tư và SLKK
with st.expander("⚙️ Tùy chỉnh vùng cắt 2 cột (Mã vật tư & SLKK)"):
    col_left = st.slider("Cắt từ mép trái (%)", 0.0, 0.8, 0.3, 0.05)
    col_right = st.slider("Cắt đến mép phải (%)", 0.2, 1.0, 1.0, 0.05)
    col_top = st.slider("Cắt từ trên xuống (%)", 0.0, 0.5, 0.0, 0.05)
    col_bottom = st.slider("Cắt từ dưới lên (%)", 0.5, 1.0, 1.0, 0.05)

def crop_image_columns(image_bytes, left_p, top_p, right_p, bottom_p):
    """Hàm cắt vùng chứa 2 cột Mã vật tư và SLKK"""
    img = Image.open(io.BytesIO(image_bytes))
    width, height = img.size
    
    left = int(width * left_p)
    top = int(height * top_p)
    right = int(width * right_p)
    bottom = int(height * bottom_p)
    
    cropped = img.crop((left, top, right, bottom))
    buf = io.BytesIO()
    cropped.save(buf, format="JPEG")
    return buf.getvalue()

if uploaded_files:
    st.write(f"Đã chọn **{len(uploaded_files)}** ảnh phiếu kiểm kê.")
    
    # Hiển thị ảnh sau khi cắt thử nghiệm để kiểm tra
    for idx, file in enumerate(uploaded_files):
        cropped_bytes = crop_image_columns(file.getvalue(), col_left, col_top, col_right, col_bottom)
        st.image(cropped_bytes, caption=f"Vùng 2 cột đã cắt từ ảnh {idx+1}", width=400)

    if st.button("🚀 Bắt đầu đọc và ghép dữ liệu", type="primary"):
        with st.spinner("AI đang xử lý các ảnh, giữ nguyên dòng trống và trích xuất dữ liệu..."):
            try:
                api_key = st.secrets["GEMINI_API_KEY"]
                genai.configure(api_key=api_key)
                
                # Sử dụng gemini-1.5-flash hỗ trợ đa ảnh xuất sắc
                model = genai.GenerativeModel('gemini-1.5-flash')
                
                # Chuẩn bị danh sách nội dung gửi cho AI (gồm các ảnh đã cắt)
                contents = [
                    "Bạn là trợ lý kiểm kê kho chuyên nghiệp. Hãy đọc các ảnh phiếu kiểm kê được cung cấp dưới đây.\n"
                    "YÊU CẦU QUAN TRỌNG:\n"
                    "1. Chỉ tập trung vào 2 cột: Mã vật tư và Số lượng kiểm kê (SLKK).\n"
                    "2. Phải đọc tuần tự từng dòng từ trên xuống dưới, TUYỆT ĐỐI KHÔNG BỎ QUA các dòng trống hoặc dòng không ghi số lượng. Nếu dòng trống, hãy để giá trị slkk là null hoặc chuỗi rỗng để giữ nguyên thứ tự dòng.\n"
                    "3. Trả về kết quả dưới dạng một danh sách JSON thuần túy (JSON array) với định dạng chuẩn:\n"
                    '[{"material_no": "Mã vật tư tương ứng", "slkk": "Số lượng đếm được hoặc trống"}]'
                ]
                
                for file in uploaded_files:
                    c_bytes = crop_image_columns(file.getvalue(), col_left, col_top, col_right, col_bottom)
                    contents.append({"mime_type": "image/jpeg", "data": c_bytes})
                
                response = model.generate_content(contents)
                
                # Làm sạch chuỗi JSON trả về từ AI
                raw_text = response.text.strip()
                if "```json" in raw_text:
                    raw_text = raw_text.split("```json")[1].split("```")[0].strip()
                elif "```" in raw_text:
                    raw_text = raw_text.split("```")[1].split("```")[0].strip()
                
                data_list = json.loads(raw_text)
                df_result = pd.DataFrame(data_list)
                
                st.success("✅ Trích xuất và ghép dữ liệu thành công!")
                st.dataframe(df_result, use_container_width=True)
                
                # Lưu vào session_state để chuẩn bị map ngược lại vào hệ thống kho chính
                st.session_state['scanned_result'] = df_result
                
            except Exception as e:
                st.error(f"Lỗi trong quá trình xử lý: {e}")

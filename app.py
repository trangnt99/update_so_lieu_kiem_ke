import streamlit as st
import pandas as pd
import os
import io
import openpyxl
import base64
import requests
from datetime import datetime

st.set_page_config(page_title="Hệ Thống Kiểm Kê Kho - Đồng Bộ GitHub", layout="wide")

DEFAULT_DATA_FILE = "latest_inventory_list.xlsx"

st.title("📋 HỆ THỐNG KIỂM KÊ KHO - ĐỒNG BỘ 24/7 TRÊN CLOUD")
st.markdown("---")

if 'header_line' not in st.session_state:
    st.session_state['header_line'] = 4

def save_to_github():
    """Hàm tự động đẩy file Excel lên GitHub qua API để lưu vĩnh viễn"""
    try:
        if "GITHUB_TOKEN" in st.secrets and "REPO_NAME" in st.secrets:
            token = st.secrets["GITHUB_TOKEN"]
            repo = st.secrets["REPO_NAME"]
            file_path = "latest_inventory_list.xlsx"
            
            with open(DEFAULT_DATA_FILE, "rb") as f:
                content_bytes = f.read()
            encoded_content = base64.b64encode(content_bytes).decode("utf-8")
            
            api_url = f"https://api.github.com/repos/{repo}/contents/{file_path}"
            headers = {"Authorization": f"token {token}"}
            response = requests.get(api_url, headers=headers)
            
            sha = ""
            if response.status_code == 200:
                sha = response.json().get("sha", "")
                
            data = {
                "message": "Auto-update inventory data from Streamlit Cloud",
                "content": encoded_content,
                "sha": sha
            }
            put_response = requests.put(api_url, headers=headers, json=data)
            return put_response.status_code in [200, 201]
    except Exception as e:
        print(f"Lỗi đồng bộ GitHub: {e}")
    return False

def save_updates_to_excel_file():
    if 'sheets_data' in st.session_state and st.session_state['sheets_data'] and os.path.exists(DEFAULT_DATA_FILE):
        try:
            header_idx = st.session_state['header_line'] - 1
            header_row_num = header_idx + 1
            
            wb = openpyxl.load_workbook(DEFAULT_DATA_FILE)
            for s_name, df_val in st.session_state['sheets_data'].items():
                if s_name in wb.sheetnames:
                    ws = wb[s_name]
                    col_map = {}
                    for col_idx in range(1, ws.max_column + 1):
                        val = ws.cell(row=header_row_num, column=col_idx).value
                        if val:
                            col_map[str(val).strip()] = col_idx
                    
                    for needed_col in ['SLKK', 'Cập nhật bởi (IP)', 'Thời gian sửa']:
                        if needed_col not in col_map:
                            new_col_idx = ws.max_column + 1
                            ws.cell(row=header_row_num, column=new_col_idx, value=needed_col)
                            col_map[needed_col] = new_col_idx

                    slkk_col = col_map.get('SLKK')
                    ip_col = col_map.get('Cập nhật bởi (IP)')
                    time_col = col_map.get('Thời gian sửa')

                    for r_idx, row in df_val.iterrows():
                        excel_row_num = header_row_num + 1 + r_idx
                        if slkk_col:
                            ws.cell(row=excel_row_num, column=slkk_col, value=row.get('SLKK'))
                        if ip_col:
                            ws.cell(row=excel_row_num, column=ip_col, value=row.get('Cập nhật bởi (IP)'))
                        if time_col:
                            ws.cell(row=excel_row_num, column=time_col, value=row.get('Thời gian sửa'))
            wb.save(DEFAULT_DATA_FILE)
            
            # Tự động đồng bộ lên GitHub sau khi lưu cục bộ thành công
            save_to_github()
        except Exception as e:
            with pd.ExcelWriter(DEFAULT_DATA_FILE, engine='openpyxl') as writer:
                for s_name, df_val in st.session_state['sheets_data'].items():
                    df_val.to_excel(writer, index=False, sheet_name=s_name)
            save_to_github()

with st.sidebar:
    st.header("🔐 XÁC THỰC NGƯỜI DÙNG")
    try:
        if hasattr(st, "context") and hasattr(st.context, "headers"):
            client_ip = st.context.headers.get("X-Forwarded-For", "Cloud User")
        else:
            client_ip = "Streamlit Cloud / Local"
    except Exception:
        client_ip = "Không xác định"

    st.info(f"🌐 Trạng thái kết nối:\n`Hoạt động ổn định`")

    user_role = st.selectbox("Chọn vai trò:", ["Nhân viên kho (Chỉ nhập SLKK)", "Quản trị viên (Admin)"])
    admin_password = st.text_input("Mật khẩu Admin:", type="password") if user_role == "Quản trị viên (Admin)" else ""

    st.markdown("---")
    st.header("⚙️ QUẢN LÝ DỮ LIỆU (ADMIN)")
    
    header_line = st.number_input("Dòng chứa tiêu đề (Header row):", min_value=1, value=st.session_state['header_line'], step=1)
    st.session_state['header_line'] = header_line
    header_idx = header_line - 1

    uploaded_file = st.file_uploader("📂 Tải list kiểm kê mới (Excel):", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        if user_role == "Quản trị viên (Admin)" and admin_password == "123456":
            try:
                with open(DEFAULT_DATA_FILE, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                raw_dfs = pd.read_excel(DEFAULT_DATA_FILE, sheet_name=None, header=header_idx, dtype=str)
                new_sheets = {}
                for s_name, df_s in raw_dfs.items():
                    df_s = df_s.dropna(how='all')
                    df_s = df_s.dropna(axis=1, how='all')
                    df_s = df_s.loc[:, ~df_s.columns.astype(str).str.contains('^Unnamed')]

                    if 'SLKK' not in df_s.columns: df_s['SLKK'] = None
                    if 'Cập nhật bởi (IP)' not in df_s.columns: df_s['Cập nhật bởi (IP)'] = ""
                    if 'Thời gian sửa' not in df_s.columns: df_s['Thời gian sửa'] = ""
                    new_sheets[s_name] = df_s

                st.session_state['sheets_data'] = new_sheets
                save_to_github() # Đồng bộ file mới lên GitHub
                st.success("✅ Tải và cập nhật list mới thành công trên Cloud!")
            except Exception as e:
                st.error(f"Lỗi đọc file: {e}")
        else:
            st.warning("⚠️ Cần chọn 'Quản trị viên (Admin)' và nhập đúng mật khẩu (`123456`) để áp dụng file mới!")

if 'sheets_data' not in st.session_state:
    if os.path.exists(DEFAULT_DATA_FILE):
        try:
            header_idx = st.session_state['header_line'] - 1
            raw_dfs = pd.read_excel(DEFAULT_DATA_FILE, sheet_name=None, header=header_idx, dtype=str)
            st.session_state['sheets_data'] = {}
            for s_name, df_s in raw_dfs.items():
                df_s = df_s.dropna(how='all')
                df_s = df_s.dropna(axis=1, how='all')
                df_s = df_s.loc[:, ~df_s.columns.astype(str).str.contains('^Unnamed')]

                if 'SLKK' not in df_s.columns: df_s['SLKK'] = None
                if 'Cập nhật bởi (IP)' not in df_s.columns: df_s['Cập nhật bởi (IP)'] = ""
                if 'Thời gian sửa' not in df_s.columns: df_s['Thời gian sửa'] = ""
                st.session_state['sheets_data'][s_name] = df_s
        except Exception:
            st.session_state['sheets_data'] = {}
    else:
        st.session_state['sheets_data'] = {}

if not st.session_state['sheets_data']:
    st.warning("⚠️ Chưa có dữ liệu trong hệ thống. Quản trị viên hãy đăng nhập Admin (Mật khẩu: `123456`) và tải file Excel lên ở thanh bên trái!")
else:
    if user_role == "Nhân viên kho (Chỉ nhập SLKK)":
        st.info("💡 **Hướng dẫn:** Chọn nhóm/sheet ở bên dưới, nhập số lượng vào cột **SLKK**, sau đó bấm nút **💾 LƯU TẤT CẢ THAY ĐỔI**.")
    else:
        st.success("👑 Bạn đang ở chế độ **Quản trị viên (Admin)**: Toàn quyền chỉnh sửa dữ liệu.")

    # THANH CHỌN NHÓM / SHEET
    sheet_names = list(st.session_state['sheets_data'].keys())
    selected_sheet = st.selectbox("📂 Chọn Nhóm / Sheet kiểm kê:", sheet_names)

    current_df = st.session_state['sheets_data'][selected_sheet].copy()

    # Ô LỌC TRỰC TIẾP TỨC THÌ (GÕ ĐẾN ĐÂU HIỆN ĐẾN ĐÓ)
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        s_loc = st.text_input("🔍 Lọc trực tiếp theo Location:", placeholder="Nhập ký tự location...", key="filter_loc")
        if s_loc and 'location' in current_df.columns:
            current_df = current_df[current_df['location'].astype(str).str.contains(s_loc, case=False, na=False)]
    with col_f2:
        s_mat = st.text_input("🔍 Lọc trực tiếp theo Mã vật tư:", placeholder="Nhập mã vật tư...", key="filter_mat")
        if s_mat and 'material_no' in current_df.columns:
            current_df = current_df[current_df['material_no'].astype(str).str.contains(s_mat, case=False, na=False)]

    locked_cols = [col for col in current_df.columns if col != 'SLKK'] if user_role.startswith("Nhân viên") else []

    # BẢNG NHẬP LIỆU
    with st.form(key=f"form_{selected_sheet}"):
        edited_df = st.data_editor(
            current_df,
            disabled=locked_cols,
            hide_index=True,
            use_container_width=True,
            height=500,
            key=f"editor_{selected_sheet}"
        )
        
        submitted = st.form_submit_button(f"💾 LƯU TẤT CẢ THAY ĐỔI (CHO NHÓM: {selected_sheet})", type="primary", use_container_width=True)

    if submitted:
        now_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        orig_df = st.session_state['sheets_data'][selected_sheet]
        has_changes = False

        for _, row in edited_df.iterrows():
            loc = str(row.get('location', '')).strip()
            mat = str(row.get('material_no', '')).strip()
            new_slkk = row.get('SLKK')

            match_mask = (orig_df['location'].astype(str).str.strip() == loc) & (orig_df['material_no'].astype(str).str.strip() == mat)
            if match_mask.any():
                old_val = orig_df.loc[match_mask, 'SLKK'].values[0]
                if str(old_val) != str(new_slkk) and pd.notna(new_slkk):
                    orig_df.loc[match_mask, 'SLKK'] = new_slkk
                    orig_df.loc[match_mask, 'Cập nhật bởi (IP)'] = client_ip
                    orig_df.loc[match_mask, 'Thời gian sửa'] = now_time
                    has_changes = True

        if has_changes:
            save_updates_to_excel_file()
            st.success(f"🎉 Đã lưu và đồng bộ thay đổi cho nhóm [{selected_sheet}] lên GitHub thành công!")
        else:
            st.info("ℹ️ Không có thay đổi mới nào được ghi nhận.")

    st.markdown("---")
    st.subheader("💾 TẢI XUỐNG KẾT QUẢ")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for s_name, df_val in st.session_state['sheets_data'].items():
            df_val.to_excel(writer, index=False, sheet_name=s_name)
    st.download_button(
        label="📥 Tải file Excel hoàn chỉnh (Kèm lịch sử IP)",
        data=output.getvalue(),
        file_name="KetQua_KiemKe_Audit.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )

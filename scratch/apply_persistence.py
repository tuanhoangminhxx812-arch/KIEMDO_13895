import re

with open("app_tk13895.py", "r", encoding="utf-8") as f:
    content = f.read()

# 1. Add import json
content = content.replace("import os\nimport re", "import os\nimport re\nimport json")

# 2. Add Persistence Constants and Helpers before Streamlit UI
target_ui_header = """# ==============================================================================
# 5. GIAO DIỆN CHÍNH STREAMLIT
# ==============================================================================
# Header Banner
st.markdown(\"\"\"
<div class=\"main-header\">
    <h1>HỆ THỐNG ĐỐI CHIẾU & KIỂM DÒ THUẾ GTGT ĐẦU VÀO TK 13895</h1>
    <p>Công ty Điện lực Vũng Tàu — Phân loại tự động 5 tầng: HĐ tồn, Đã giải trừ, Giải trừ số dư 2025, Bù trừ kỹ thuật & Điện mua nội bộ</p>
</div>
\"\"\", unsafe_allow_html=True)

# Sidebar hướng dẫn & tải file
with st.sidebar:
    st.markdown(\"\"\"
    <div style=\"background-color: #1E3A8A; color: white; padding: 12px 14px; border-radius: 8px; text-align: center; margin-bottom: 15px;\">
        <div style=\"font-size: 20px; font-weight: 800; letter-spacing: 1px;\">⚡ EVNHCMC</div>
        <div style=\"font-size: 12px; font-weight: 500; color: #93C5FD; margin-top: 2px;\">CÔNG TY ĐIỆN LỰC VŨNG TÀU</div>
    </div>
    \"\"\", unsafe_allow_html=True)
    st.markdown(\"### 📥 Nạp Sổ EVN_GL_094\")
    uploaded_file = st.file_uploader(
        \"Kéo thả hoặc chọn file Excel sổ chi tiết TK 13895 (.xlsx)\",
        type=[\"xlsx\"],
        help=\"Xuất file từ phân hệ GL ERP của EVN (EVN_GL_094)\"
    )
    st.markdown(\"---\")
    st.markdown(\"#### ⚙️ Quy trình kiểm dò 5 tầng:\")
    st.caption(\"**Tầng 1:** Triệt tiêu nội bộ (+/- Nợ AP, GL 1366, AR-GL, số lẻ bù trừ)\")
    st.caption(\"**Tầng 2:** Khớp chính xác 1-1 Số HĐ + Số tiền\")
    st.caption(\"**Tầng 3:** Khớp gộp n dòng Nợ theo công trình AP (như BIDV)\")
    st.caption(\"**Tầng 4:** Khớp thông minh các ca gõ nhầm số HĐ khi giải trừ\")
    st.caption(\"**Tầng 5:** Tách HĐ tồn 2026 vs. Giải trừ số dư đầu kỳ (HĐ 2025)\")

if uploaded_file is None:
    st.info(\"👋 **Chào anh! Vui lòng tải file sổ chi tiết tài khoản 13895 (EVN_GL_094...xlsx) vào khung bên trái để bắt đầu phân tích.**\")
    
    # Nút nạp thử file mẫu nếu có trong thư mục
    try:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        sample_path = os.path.join(current_dir, \"GL_094_TK13895 (010126-311026).xlsx\")
        if not os.path.exists(sample_path):
            sample_path = \"GL_094_TK13895 (010126-311026).xlsx\"
        if os.path.exists(sample_path):
            with open(sample_path, \"rb\") as f:
                sample_bytes = f.read()
            if st.button(\"📂 Nạp trực tiếp file mẫu: GL_094_TK13895 (010126-311026).xlsx\"):
                uploaded_file = io.BytesIO(sample_bytes)
                uploaded_file.name = \"GL_094_TK13895 (010126-311026).xlsx\"
                st.rerun()
    except Exception:
        pass

if uploaded_file is not None:
    with st.spinner(\"Đang xử lý dữ liệu và đối chiếu 5 tầng thông minh...\"):
        open_bal_in, df_in = parse_gl094_excel(uploaded_file)"""

replacement_ui = """# ==============================================================================
# 5. QUẢN LÝ DỮ LIỆU BỀN VỮNG & GIAO DIỆN STREAMLIT
# ==============================================================================
APP_DIR = os.path.dirname(os.path.abspath(__file__))
SAVED_FILE_PATH = os.path.join(APP_DIR, "saved_gl094.xlsx")
DEFAULT_FILE_PATH = os.path.join(APP_DIR, "GL_094_TK13895 (010126-311026).xlsx")
META_FILE_PATH = os.path.join(APP_DIR, "saved_gl094_meta.json")

def load_saved_meta():
    if os.path.exists(META_FILE_PATH):
        try:
            with open(META_FILE_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return None

def save_active_file(file_bytes, filename):
    try:
        with open(SAVED_FILE_PATH, "wb") as f:
            f.write(file_bytes)
        meta = {
            "filename": filename,
            "updated_at": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
            "size_kb": round(len(file_bytes) / 1024, 1)
        }
        with open(META_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        return meta
    except Exception:
        return {"filename": filename, "updated_at": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"), "size_kb": 0}

# Sidebar: Tải và cập nhật file mới
with st.sidebar:
    st.markdown(\"\"\"
    <div style=\"background-color: #1E3A8A; color: white; padding: 12px 14px; border-radius: 8px; text-align: center; margin-bottom: 15px;\">
        <div style=\"font-size: 20px; font-weight: 800; letter-spacing: 1px;\">⚡ EVNHCMC</div>
        <div style=\"font-size: 12px; font-weight: 500; color: #93C5FD; margin-top: 2px;\">CÔNG TY ĐIỆN LỰC VŨNG TÀU</div>
    </div>
    \"\"\", unsafe_allow_html=True)
    
    st.markdown(\"### 📤 Cập Nhật Số Liệu Mới\")
    uploaded_file = st.file_uploader(
        \"Kéo thả hoặc chọn file Excel GL_094 mới (.xlsx):\",
        type=[\"xlsx\"],
        help=\"Khi anh tải file mới lên, hệ thống sẽ tự động cập nhật số liệu và lưu lại làm dữ liệu mặc định cho những lần mở sau.\"
    )

# Xác định file dữ liệu hoạt động (Ưu tiên: Upload mới > File đã lưu > File mặc định)
active_bytes = None
active_filename = None
active_updated_at = None

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    if st.session_state.get('last_uploaded_name') != uploaded_file.name or st.session_state.get('last_uploaded_size') != len(file_bytes):
        meta = save_active_file(file_bytes, uploaded_file.name)
        st.session_state['last_uploaded_name'] = uploaded_file.name
        st.session_state['last_uploaded_size'] = len(file_bytes)
        st.sidebar.success(f\"🎉 Đã lưu và cập nhật dữ liệu từ: **{uploaded_file.name}**!\")
    active_bytes = file_bytes
    active_filename = uploaded_file.name
    meta = load_saved_meta()
    active_updated_at = meta.get(\"updated_at\") if meta else datetime.datetime.now().strftime(\"%d/%m/%Y %H:%M\")
else:
    if os.path.exists(SAVED_FILE_PATH):
        try:
            with open(SAVED_FILE_PATH, \"rb\") as f:
                active_bytes = f.read()
            meta = load_saved_meta()
            active_filename = meta.get(\"filename\", \"saved_gl094.xlsx\") if meta else \"saved_gl094.xlsx\"
            active_updated_at = meta.get(\"updated_at\", \"Đã lưu trước đó\") if meta else \"Đã lưu trước đó\"
        except Exception:
            active_bytes = None
            
    if active_bytes is None:
        if os.path.exists(DEFAULT_FILE_PATH):
            try:
                with open(DEFAULT_FILE_PATH, \"rb\") as f:
                    active_bytes = f.read()
                active_filename = \"GL_094_TK13895 (010126-311026).xlsx\"
                active_updated_at = \"Bản gốc mặc định (01/01 - 31/10/2026)\"
                save_active_file(active_bytes, active_filename)
            except Exception:
                active_bytes = None

# Sidebar: Thẻ thông tin file dữ liệu đang hoạt động & Nút khôi phục
with st.sidebar:
    if active_filename:
        st.markdown(f\"\"\"
        <div style=\"background-color: #F0FDF4; border: 1px solid #86EFAC; border-left: 4px solid #16A34A; border-radius: 8px; padding: 12px 14px; margin-top: 10px; margin-bottom: 12px;\">
            <div style=\"font-size: 11px; font-weight: bold; color: #166534; text-transform: uppercase;\">📂 Dữ liệu đang hoạt động</div>
            <div style=\"font-size: 13px; font-weight: 700; color: #1E3A8A; margin-top: 4px; word-break: break-all;\">{active_filename}</div>
            <div style=\"font-size: 11px; color: #475569; margin-top: 3px;\">🕒 Cập nhật: <b>{active_updated_at}</b></div>
            <div style=\"font-size: 11px; color: #166534; margin-top: 4px;\">✅ <i>Tự động nạp sẵn mỗi khi mở app</i></div>
        </div>
        \"\"\", unsafe_allow_html=True)
        
        if os.path.exists(DEFAULT_FILE_PATH) and active_filename != \"GL_094_TK13895 (010126-311026).xlsx\":
            if st.button(\"🔄 Khôi phục về sổ gốc ban đầu\"):
                try:
                    with open(DEFAULT_FILE_PATH, \"rb\") as f:
                        def_bytes = f.read()
                    save_active_file(def_bytes, \"GL_094_TK13895 (010126-311026).xlsx\")
                    st.session_state.pop('last_uploaded_name', None)
                    st.session_state.pop('last_uploaded_size', None)
                    st.rerun()
                except Exception:
                    pass

    st.markdown(\"---\")
    st.markdown(\"#### ⚙️ Quy trình kiểm dò 5 tầng:\")
    st.caption(\"**Tầng 1:** Triệt tiêu nội bộ (+/- Nợ AP, GL 1366, AR-GL, số lẻ bù trừ)\")
    st.caption(\"**Tầng 2:** Khớp chính xác 1-1 Số HĐ + Số tiền\")
    st.caption(\"**Tầng 3:** Khớp gộp n dòng Nợ theo công trình AP (như BIDV)\")
    st.caption(\"**Tầng 4:** Khớp thông minh các ca gõ nhầm số HĐ khi giải trừ\")
    st.caption(\"**Tầng 5:** Tách HĐ tồn 2026 vs. Giải trừ số dư đầu kỳ (HĐ 2025)\")

# Header Banner
st.markdown(f\"\"\"
<div class=\"main-header\">
    <h1>HỆ THỐNG ĐỐI CHIẾU & KIỂM DÒ THUẾ GTGT ĐẦU VÀO TK 13895</h1>
    <p>Công ty Điện lực Vũng Tàu — Phân loại tự động 5 tầng: HĐ tồn, Đã giải trừ, Giải trừ số dư 2025, Bù trừ kỹ thuật & Điện mua nội bộ</p>
</div>
<div style=\"background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; padding: 8px 16px; margin-top: -12px; margin-bottom: 20px; font-size: 13px; color: #475569; display: flex; justify-content: space-between; align-items: center;\">
    <span>📂 <b>Sổ chi tiết đang đối chiếu:</b> <code style=\"color: #1E3A8A; font-weight: bold;\">{active_filename or 'Chưa có file'}</code></span>
    <span>🕒 <b>Thời gian cập nhật:</b> <i>{active_updated_at or 'Chưa xác định'}</i></span>
</div>
\"\"\", unsafe_allow_html=True)

if active_bytes is None:
    st.info(\"👋 **Chào anh! Vui lòng tải file sổ chi tiết tài khoản 13895 (EVN_GL_094...xlsx) vào khung bên trái để bắt đầu phân tích.**\")

if active_bytes is not None:
    active_stream = io.BytesIO(active_bytes)
    active_stream.name = active_filename
    with st.spinner(\"Đang tải dữ liệu và đối chiếu 5 tầng thông minh...\"):
        open_bal_in, df_in = parse_gl094_excel(active_stream)"""

assert target_ui_header in content, "target_ui_header not found in content!"
content = content.replace(target_ui_header, replacement_ui)

# 3. Update export_excel_report call to pass io.BytesIO(active_bytes)
old_export_call = "excel_bytes = export_excel_report(open_bal, df_ton, df_da_gt, df_2025, df_dnb, df_t1_netoff, df_mistyped, uploaded_file=uploaded_file, df_raw=df_raw)"
new_export_call = "excel_bytes = export_excel_report(open_bal, df_ton, df_da_gt, df_2025, df_dnb, df_t1_netoff, df_mistyped, uploaded_file=io.BytesIO(active_bytes), df_raw=df_raw)"
assert old_export_call in content, "old_export_call not found!"
content = content.replace(old_export_call, new_export_call)

# Validate syntax
import ast
ast.parse(content)
print("AST validation passed! Writing changes to app_tk13895.py...")

with open("app_tk13895.py", "w", encoding="utf-8") as f:
    f.write(content)

print("Updated app_tk13895.py successfully!")

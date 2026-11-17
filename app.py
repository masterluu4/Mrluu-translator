# app.py
# Ứng dụng Web dịch tài liệu kỹ thuật chuyên nghiệp (THACO)
# Có UI/UX nâng cao, Console Logging và API Verification

import streamlit as st
import google.generativeai as genai
import time
import os
import tempfile
from copy import copy

# ----- Thư viện xử lý File -----
from docx import Document
from docx.text.run import Run
from docx.oxml.ns import qn
from docx.table import Table
import openpyxl
from pdf2docx import Converter

# ============================================================
# CẤU HÌNH GIAO DIỆN & CSS (PRO UI)
# ============================================================
st.set_page_config(page_title="THACO Technical Translator", page_icon="🌐", layout="wide")

st.markdown("""
    <style>
    .main-title { font-size: 38px !important; font-weight: 700; color: #1E3A8A; margin-bottom: 0px; }
    .sub-title { font-size: 16px; color: #64748B; margin-bottom: 20px; }
    .log-container { background-color: #0F172A; color: #10B981; font-family: 'Courier New', monospace; padding: 15px; border-radius: 8px; height: 300px; overflow-y: auto; font-size: 13px; line-height: 1.5; }
    .stProgress > div > div > div > div { background-color: #3B82F6; }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# SYSTEM PROMPT (NGỮ CẢNH THACO)
# ============================================================
SYSTEM_PROMPT = """You are an expert technical translator specializing in automotive manufacturing, vehicle engineering, and after-sales quality management. Your task is to translate Vietnamese text into precise, grammatically correct, and professional English.
Context: The documents are from the After-Sales Quality Management Department of THACO AUTO.
Requirements:
- Use standard automotive industry and THACO-specific terminology (e.g., IATF 16949, 8D report, PFMEA, 5-Why, Ishikawa 4M1E, RACI, NCR/CAR, Warranty Claims, CKD assembly).
- Tone: Maintain a highly formal, precise, and PROACTIVE tone. Suitable for reporting to executives and international partners.
- Output: Return ONLY the translated English text. No notes or explanations."""

# ============================================================
# LOGGING SYSTEM (THỜI GIAN THỰC)
# ============================================================
if "logs" not in st.session_state:
    st.session_state.logs = []

def add_log(message: str, log_placeholder):
    """Thêm log mới và cuộn hiển thị trên giao diện"""
    time_str = time.strftime("%H:%M:%S")
    full_msg = f"[{time_str}] {message}"
    st.session_state.logs.append(full_msg)
    # Giữ lại 30 dòng log gần nhất để không nặng UI
    display_logs = "\n".join(st.session_state.logs[-30:])
    log_placeholder.markdown(f"<div class='log-container'>{display_logs}</div>", unsafe_allow_html=True)

# ============================================================
# KIỂM TRA & KHỞI TẠO API
# ============================================================
def verify_api_key(api_key: str):
    """Gửi một request nhỏ để test API Key"""
    try:
        genai.configure(api_key=api_key)
        test_model = genai.GenerativeModel("gemini-flash-latest")
        test_model.generate_content("Hi")
        return True, "Kết nối API Google Gemini thành công!"
    except Exception as e:
        return False, f"Lỗi API Key: {str(e)}"

def init_gemini():
    return genai.GenerativeModel(
        model_name="gemini-flash-latest",
        system_instruction=SYSTEM_PROMPT,
    )

def translate_text(model, text: str, log_placeholder) -> str:
    """Hàm dịch có xử lý Rate Limit và Log chi tiết"""
    if not text or not text.strip():
        return text

    core = text.strip()
    lead = text[: len(text) - len(text.lstrip())]
    trail = text[len(text.rstrip()) :]

    # Trích xuất 40 ký tự đầu để ghi log cho gọn
    short_text = core[:40] + "..." if len(core) > 40 else core
    add_log(f"Đang dịch: '{short_text}'", log_placeholder)

    for attempt in range(4):
        try:
            response = model.generate_content(core)
            translated = response.text.strip() if response.text else core
            time.sleep(0.5) # Chống spam API
            return lead + translated + trail
        except Exception as e:
            error_str = str(e)
            if "429" in error_str or "ResourceExhausted" in error_str:
                wait_time = (2 ** attempt) * 2  # 2s, 4s, 8s
                add_log(f"⚠️ [Rate Limit] API quá tải. Đang đợi {wait_time} giây để thử lại...", log_placeholder)
                time.sleep(wait_time)
            else:
                add_log(f"❌ [Lỗi] {error_str}. Đang thử lại...", log_placeholder)
                time.sleep(2)
                
    add_log(f"🚫 [Bỏ qua] Không thể dịch đoạn này sau nhiều lần thử.", log_placeholder)
    return lead + core + trail

# ============================================================
# CÁC HÀM XỬ LÝ FILE (WORD, EXCEL, PDF)
# ============================================================
def process_excel(input_path, output_path, model, log_placeholder, progress_bar):
    wb = openpyxl.load_workbook(input_path)
    add_log("Đang quét cấu trúc file Excel...", log_placeholder)
    
    # Bước 1: Đếm tổng số ô cần dịch để làm Progress Bar
    cells_to_translate = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and not cell.value.startswith("=") and cell.value.strip():
                    cells_to_translate.append(cell)
                    
    total_cells = len(cells_to_translate)
    add_log(f"Phát hiện tổng cộng {total_cells} ô dữ liệu cần dịch.", log_placeholder)

    if total_cells == 0:
        wb.save(output_path)
        progress_bar.progress(1.0)
        return

    # Bước 2: Dịch và cập nhật Progress Bar
    for i, cell in enumerate(cells_to_translate):
        cell.value = translate_text(model, cell.value, log_placeholder)
        progress_bar.progress((i + 1) / total_cells)

    add_log("Đang lưu lại định dạng Excel...", log_placeholder)
    wb.save(output_path)

def process_word(input_path, output_path, model, log_placeholder, progress_bar):
    add_log("Đang quét cấu trúc file Word...", log_placeholder)
    doc = Document(input_path)
    
    runs_to_translate = []
    # Quét Body
    for p in doc.paragraphs:
        for r in p.runs:
            if r.text.strip(): runs_to_translate.append(r)
    # Quét Table
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        if r.text.strip(): runs_to_translate.append(r)
                        
    total_runs = len(runs_to_translate)
    add_log(f"Phát hiện tổng cộng {total_runs} phân đoạn văn bản cần dịch.", log_placeholder)

    if total_runs == 0:
        doc.save(output_path)
        progress_bar.progress(1.0)
        return

    for i, run in enumerate(runs_to_translate):
        run.text = translate_text(model, run.text, log_placeholder)
        progress_bar.progress((i + 1) / total_runs)

    add_log("Đang lưu lại định dạng Word...", log_placeholder)
    doc.save(output_path)

def process_pdf(input_path, output_path, model, log_placeholder, progress_bar):
    add_log("Đang trích xuất layout PDF sang Word ẩn...", log_placeholder)
    temp_docx = input_path.replace(".pdf", "_temp.docx")
    
    # Bước 1: Chuyển PDF -> Word (Chiếm 20% tiến trình)
    cv = Converter(input_path)
    cv.convert(temp_docx)
    cv.close()
    progress_bar.progress(0.2)
    add_log("Chuyển đổi layout thành công. Bắt đầu dịch văn bản...", log_placeholder)

    # Bước 2: Dịch file Word tạm (Chiếm 80% tiến trình)
    # Tạo một hàm callback bọc (wrapper) để scale progress bar từ 20% -> 100%
    class ScaledProgressBar:
        def __init__(self, pb): self.pb = pb
        def progress(self, val): self.pb.progress(0.2 + (val * 0.8))
        
    scaled_pb = ScaledProgressBar(progress_bar)
    process_word(temp_docx, output_path, model, log_placeholder, scaled_pb)

    if os.path.exists(temp_docx):
        os.remove(temp_docx)

# ============================================================
# XÂY DỰNG GIAO DIỆN CHÍNH
# ============================================================
st.markdown('<p class="main-title">🌐 Hệ thống Dịch thuật Kỹ thuật Số</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Bảo toàn cấu trúc & Định dạng • Tối ưu thuật ngữ Quản lý Chất lượng THACO</p>', unsafe_allow_html=True)

col_left, col_right = st.columns([1.2, 2])

with col_left:
    st.markdown("### ⚙️ Cấu hình Hệ thống")
    with st.container(border=True):
        api_key = st.text_input("Mã API Google Gemini:", type="password", help="Nhập API Key để kích hoạt hệ thống")
        
        if st.button("🔌 Kiểm tra Kết nối API", use_container_width=True):
            if not api_key:
                st.warning("Vui lòng nhập API Key trước.")
            else:
                with st.spinner("Đang kết nối tới server Google..."):
                    is_valid, msg = verify_api_key(api_key)
                if is_valid:
                    st.success(msg)
                else:
                    st.error(msg)
                    
    st.markdown("### 📂 Tải lên Tài liệu")
    uploaded_file = st.file_uploader("Kéo thả file Word, Excel hoặc PDF vào đây", type=["docx", "xlsx", "pdf"])
    
    if uploaded_file:
        file_ext = os.path.splitext(uploaded_file.name)[1].lower()
        if file_ext == ".pdf":
            st.info("💡 File PDF sẽ được tự động chuyển đổi sang Word sau khi dịch để bạn dễ chỉnh sửa mà không vỡ layout.")

with col_right:
    st.markdown("### 🖥️ Màn hình Giám sát (Live Console)")
    log_placeholder = st.empty()
    # Hiển thị terminal trống ban đầu
    log_placeholder.markdown("<div class='log-container'>Hệ thống sẵn sàng...<br>Vui lòng cấu hình API và tải file lên để bắt đầu.</div>", unsafe_allow_html=True)
    
    progress_bar = st.empty()
    action_placeholder = st.empty()

    if uploaded_file and api_key:
        if action_placeholder.button("🚀 BẮT ĐẦU DỊCH TỰ ĐỘNG", type="primary", use_container_width=True):
            st.session_state.logs = [] # Reset logs
            add_log("Bắt đầu khởi tạo quy trình...", log_placeholder)
            
            # 1. Test Key ngầm trước khi chạy nặng
            is_valid, msg = verify_api_key(api_key)
            if not is_valid:
                add_log(f"❌ HỦY BỎ: {msg}", log_placeholder)
                st.error("API Key không hợp lệ. Quá trình dịch đã bị hủy.")
                st.stop()
                
            # 2. Chuẩn bị file
            with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp_in:
                tmp_in.write(uploaded_file.getbuffer())
                input_path = tmp_in.name

            output_ext = ".docx" if file_ext == ".pdf" else file_ext
            output_name = uploaded_file.name.replace(file_ext, f"_THACO_EN{output_ext}")
            output_path = input_path.replace(file_ext, f"_out{output_ext}")

            # 3. Tiến hành dịch
            try:
                model = init_gemini()
                pb = progress_bar.progress(0.0)
                
                if file_ext == ".docx":
                    process_word(input_path, output_path, model, log_placeholder, pb)
                elif file_ext == ".xlsx":
                    process_excel(input_path, output_path, model, log_placeholder, pb)
                elif file_ext == ".pdf":
                    process_pdf(input_path, output_path, model, log_placeholder, pb)
                    
                add_log("✅ XONG! Đã hoàn tất toàn bộ tiến trình dịch thuật.", log_placeholder)
                
                # 4. Hiển thị nút tải file
                action_placeholder.empty() # Xóa nút Bắt đầu
                with open(output_path, "rb") as f:
                    st.download_button(
                        label=f"⬇️ TẢI XUỐNG TÀI LIỆU HOÀN THIỆN ({output_name})",
                        data=f,
                        file_name=output_name,
                        mime="application/octet-stream",
                        type="primary",
                        use_container_width=True
                    )
            except Exception as e:
                add_log(f"❌ LỖI HỆ THỐNG NGHIÊM TRỌNG: {str(e)}", log_placeholder)
            finally:
                if os.path.exists(input_path):
                    os.remove(input_path)
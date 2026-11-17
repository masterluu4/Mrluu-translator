# app.py
# Ứng dụng Web dịch tài liệu kỹ thuật (Tối ưu cho gói Google Gemini Free Tier)
# Chống lỗi 429 bằng cơ chế Rate Limiter (4.2 giây/request)

import streamlit as st
import google.generativeai as genai
import time
import os
import tempfile

# ----- Thư viện xử lý File -----
from docx import Document
from docx.table import Table
import openpyxl
from pdf2docx import Converter

# ============================================================
# CẤU HÌNH GIAO DIỆN & CSS
# ============================================================
st.set_page_config(page_title="THACO Technical Translator", page_icon="🌐", layout="wide")

st.markdown("""
    <style>
    .main-title { font-size: 38px !important; font-weight: 700; color: #1E3A8A; margin-bottom: 0px; }
    .sub-title { font-size: 16px; color: #64748B; margin-bottom: 20px; }
    .log-container { background-color: #0F172A; color: #10B981; font-family: 'Courier New', monospace; padding: 15px; border-radius: 8px; height: 350px; overflow-y: auto; font-size: 13px; line-height: 1.5; }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# SYSTEM PROMPT (NGỮ CẢNH THACO)
# ============================================================
SYSTEM_PROMPT = """You are an expert technical translator specializing in automotive manufacturing, vehicle engineering, and after-sales quality management. Your task is to translate Vietnamese text into precise, grammatically correct, and professional English.
Context: Documents from After-Sales Quality Management Department of THACO AUTO.
Requirements:
- Use standard automotive/THACO terminology (e.g., IATF 16949, 8D report, PFMEA, 5-Why, Ishikawa, RACI, NCR/CAR, CKD).
- Tone: Formal, precise, and PROACTIVE.
- Output: Return ONLY the translated English text. No notes or explanations."""

# ============================================================
# LOGGING SYSTEM
# ============================================================
if "logs" not in st.session_state:
    st.session_state.logs = []

def add_log(message: str, log_placeholder):
    time_str = time.strftime("%H:%M:%S")
    full_msg = f"[{time_str}] {message}"
    st.session_state.logs.append(full_msg)
    display_logs = "\n".join(st.session_state.logs[-35:])
    log_placeholder.markdown(f"<div class='log-container'>{display_logs}</div>", unsafe_allow_html=True)

# ============================================================
# KHỞI TẠO API (CỐ ĐỊNH GEMINI-1.5-FLASH CHO FREE TIER)
# ============================================================
def verify_api_key(api_key: str):
    try:
        genai.configure(api_key=api_key)
        test_model = genai.GenerativeModel("gemini-1.5-flash")
        test_model.generate_content("Hi")
        return True, "Kết nối thành công! Đang sử dụng model: gemini-1.5-flash (Free Tier)"
    except Exception as e:
        return False, f"Lỗi API: {str(e)}"

def init_gemini():
    return genai.GenerativeModel(
        model_name="gemini-1.5-flash",
        system_instruction=SYSTEM_PROMPT,
    )

def translate_text(model, text: str, log_placeholder) -> str:
    if not text or not text.strip():
        return text

    core = text.strip()
    lead = text[: len(text) - len(text.lstrip())]
    trail = text[len(text.rstrip()) :]

    short_text = core[:40] + "..." if len(core) > 40 else core
    add_log(f"Dịch: '{short_text}'", log_placeholder)

    # Vòng lặp thử lại nếu lỗi
    for attempt in range(3):
        try:
            response = model.generate_content(core)
            translated = response.text.strip() if response.text else core
            
            # GIỚI HẠN TỐC ĐỘ FREE TIER (15 Request / Phút -> ~4.1s / Request)
            add_log("⏳ Đang ngủ 4.2s để tránh vượt limit Free Tier...", log_placeholder)
            time.sleep(4.2) 
            
            return lead + translated + trail
            
        except Exception as e:
            error_str = str(e)
            if "429" in error_str or "ResourceExhausted" in error_str:
                wait_time = 15 # Nếu lỡ dính 429 thì nghỉ hẳn 15 giây
                add_log(f"⚠️ [Quá tải] Nghỉ {wait_time}s rồi thử lại...", log_placeholder)
                time.sleep(wait_time)
            else:
                add_log(f"❌ [Lỗi] {error_str}. Thử lại...", log_placeholder)
                time.sleep(3)
                
    add_log(f"🚫 [Bỏ qua] Không thể dịch đoạn này.", log_placeholder)
    return lead + core + trail

# ============================================================
# CÁC HÀM XỬ LÝ FILE
# ============================================================
def process_excel(input_path, output_path, model, log_placeholder, progress_bar):
    wb = openpyxl.load_workbook(input_path)
    cells_to_translate = []
    
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and not str(cell.value).startswith("=") and str(cell.value).strip():
                    cells_to_translate.append(cell)
                    
    total = len(cells_to_translate)
    add_log(f"Phát hiện {total} ô dữ liệu. Thời gian dự kiến: ~{int(total * 4.2 / 60)} phút.", log_placeholder)

    if total == 0:
        wb.save(output_path)
        progress_bar.progress(1.0)
        return

    for i, cell in enumerate(cells_to_translate):
        cell.value = translate_text(model, cell.value, log_placeholder)
        progress_bar.progress((i + 1) / total)

    wb.save(output_path)

def process_word(input_path, output_path, model, log_placeholder, progress_bar):
    doc = Document(input_path)
    runs_to_translate = []
    
    for p in doc.paragraphs:
        for r in p.runs:
            if r.text.strip(): runs_to_translate.append(r)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        if r.text.strip(): runs_to_translate.append(r)
                        
    total = len(runs_to_translate)
    add_log(f"Phát hiện {total} đoạn văn. Thời gian dự kiến: ~{int(total * 4.2 / 60)} phút.", log_placeholder)

    if total == 0:
        doc.save(output_path)
        progress_bar.progress(1.0)
        return

    for i, run in enumerate(runs_to_translate):
        run.text = translate_text(model, run.text, log_placeholder)
        progress_bar.progress((i + 1) / total)

    doc.save(output_path)

def process_pdf(input_path, output_path, model, log_placeholder, progress_bar):
    add_log("Đang trích xuất layout PDF sang Word ẩn...", log_placeholder)
    temp_docx = input_path.replace(".pdf", "_temp.docx")
    
    cv = Converter(input_path)
    cv.convert(temp_docx)
    cv.close()
    progress_bar.progress(0.1)

    class ScaledProgressBar:
        def __init__(self, pb): self.pb = pb
        def progress(self, val): self.pb.progress(0.1 + (val * 0.9))
        
    scaled_pb = ScaledProgressBar(progress_bar)
    process_word(temp_docx, output_path, model, log_placeholder, scaled_pb)

    if os.path.exists(temp_docx):
        os.remove(temp_docx)

# ============================================================
# GIAO DIỆN CHÍNH
# ============================================================
st.markdown('<p class="main-title">🌐 Hệ thống Dịch thuật Kỹ thuật Số</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Phiên bản Tối ưu Free Tier • Bảo toàn Cấu trúc • Ngữ cảnh THACO</p>', unsafe_allow_html=True)

col_left, col_right = st.columns([1.2, 2])

with col_left:
    st.markdown("### ⚙️ Cấu hình Hệ thống")
    with st.container(border=True):
        api_key = st.text_input("Mã API Google Gemini:", type="password")
        
        if st.button("🔌 Kiểm tra Kết nối API", use_container_width=True):
            if not api_key:
                st.warning("Vui lòng nhập API Key.")
            else:
                with st.spinner("Đang kết nối..."):
                    is_valid, msg = verify_api_key(api_key)
                if is_valid: st.success(msg)
                else: st.error(msg)
                    
    st.markdown("### 📂 Tải lên Tài liệu")
    uploaded_file = st.file_uploader("Kéo thả file Word, Excel, PDF (Ưu tiên file nhỏ)", type=["docx", "xlsx", "pdf"])

with col_right:
    st.markdown("### 🖥️ Màn hình Giám sát (Live Console)")
    log_placeholder = st.empty()
    log_placeholder.markdown("<div class='log-container'>Hệ thống sẵn sàng (Chế độ Rate Limit 4.2s)...<br>Lưu ý: Tốc độ dịch ở gói Free là 15 câu/phút.</div>", unsafe_allow_html=True)
    
    progress_bar = st.empty()
    action_placeholder = st.empty()

    if uploaded_file and api_key:
        if action_placeholder.button("🚀 BẮT ĐẦU DỊCH", type="primary", use_container_width=True):
            st.session_state.logs = []
            add_log("Bắt đầu khởi tạo quy trình...", log_placeholder)
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1].lower()) as tmp_in:
                tmp_in.write(uploaded_file.getbuffer())
                input_path = tmp_in.name

            file_ext = os.path.splitext(uploaded_file.name)[1].lower()
            output_ext = ".docx" if file_ext == ".pdf" else file_ext
            output_name = uploaded_file.name.replace(file_ext, f"_THACO_EN{output_ext}")
            output_path = input_path.replace(file_ext, f"_out{output_ext}")

            try:
                genai.configure(api_key=api_key)
                model = init_gemini()
                pb = progress_bar.progress(0.0)
                
                if file_ext == ".docx": process_word(input_path, output_path, model, log_placeholder, pb)
                elif file_ext == ".xlsx": process_excel(input_path, output_path, model, log_placeholder, pb)
                elif file_ext == ".pdf": process_pdf(input_path, output_path, model, log_placeholder, pb)
                    
                add_log("✅ HOÀN TẤT!", log_placeholder)
                action_placeholder.empty()
                
                with open(output_path, "rb") as f:
                    st.download_button(
                        label=f"⬇️ TẢI XUỐNG TÀI LIỆU ({output_name})",
                        data=f, file_name=output_name, mime="application/octet-stream",
                        type="primary", use_container_width=True
                    )
            except Exception as e:
                add_log(f"❌ LỖI HỆ THỐNG: {str(e)}", log_placeholder)
            finally:
                if os.path.exists(input_path): os.remove(input_path)
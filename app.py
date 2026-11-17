# app.py
# Ứng dụng Dịch tài liệu Kỹ thuật THACO
# Sử dụng Groq API qua Cloudflare Proxy bảo mật (Tích hợp Auto-Discovery Model)

import streamlit as st
import time
import os
import tempfile
from groq import Groq

# ----- Thư viện xử lý File -----
from docx import Document
import openpyxl
from pdf2docx import Converter

# ============================================================
# 1. CẤU HÌNH GIAO DIỆN & CSS
# ============================================================
st.set_page_config(page_title="THACO Technical Translator", page_icon="⚡", layout="wide")

st.markdown("""
    <style>
    .main-title { font-size: 38px !important; font-weight: 700; color: #1E3A8A; margin-bottom: 0px; }
    .sub-title { font-size: 16px; color: #64748B; margin-bottom: 20px; }
    .log-container { background-color: #0F172A; color: #10B981; font-family: 'Courier New', monospace; padding: 15px; border-radius: 8px; height: 350px; overflow-y: auto; font-size: 13px; line-height: 1.5; }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# 2. SYSTEM PROMPT (NGỮ CẢNH THACO)
# ============================================================
SYSTEM_PROMPT = """You are an expert technical translator specializing in automotive manufacturing, vehicle engineering, and after-sales quality management. Your task is to translate Vietnamese text into precise, grammatically correct, and professional English.
Context: Documents from After-Sales Quality Management Department of THACO AUTO.
Requirements:
- Use standard automotive/THACO terminology (e.g., IATF 16949, 8D report, PFMEA, 5-Why, Ishikawa, RACI, NCR/CAR, CKD).
- Tone: Formal, precise, and PROACTIVE.
- Output: Return ONLY the translated English text. No conversational fillers, no quotes, no explanations."""

# ============================================================
# 3. KHỞI TẠO GROQ API (QUA CLOUDFLARE PROXY)
# ============================================================
# BƯỚC QUAN TRỌNG: Thay đường dẫn Cloudflare Worker của bạn vào đây
CLOUDFLARE_WORKER_URL = "https://mrluu-groq-proxy.mrscong4.workers.dev"

def get_groq_client():
    """Khởi tạo Groq client trỏ về proxy Cloudflare thay vì server gốc"""
    return Groq(
        api_key="hidden-by-cloudflare", 
        base_url=CLOUDFLARE_WORKER_URL
    )

# ============================================================
# 4. HỆ THỐNG LOG & HÀM DỊCH THUẬT
# ============================================================
if "logs" not in st.session_state:
    st.session_state.logs = []

def add_log(message: str, log_placeholder):
    time_str = time.strftime("%H:%M:%S")
    full_msg = f"[{time_str}] {message}"
    st.session_state.logs.append(full_msg)
    display_logs = "\n".join(st.session_state.logs[-35:])
    log_placeholder.markdown(f"<div class='log-container'>{display_logs}</div>", unsafe_allow_html=True)

def translate_text(client: Groq, text: str, log_placeholder) -> str:
    if not text or not text.strip():
        return text

    core = text.strip()
    lead = text[: len(text) - len(text.lstrip())]
    trail = text[len(text.rstrip()) :]

    short_text = core[:40] + "..." if len(core) > 40 else core
    add_log(f"Đang dịch: '{short_text}'", log_placeholder)

    for attempt in range(3):
        try:
            response = client.chat.completions.create(
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": core}
                ],
                # Lấy model đã được hệ thống tự động dò tìm
                model=st.session_state.get("selected_model", "llama-3.1-8b-instant"),
                temperature=0.2, 
                max_tokens=1024
            )
            
            translated = response.choices[0].message.content.strip()
            time.sleep(1.5) 
            return lead + translated + trail
            
        except Exception as e:
            error_str = str(e)
            if "429" in error_str or "rate_limit" in error_str.lower():
                wait_time = 5
                add_log(f"⚠️ [Rate Limit] Groq yêu cầu chờ. Nghỉ {wait_time}s...", log_placeholder)
                time.sleep(wait_time)
            else:
                add_log(f"❌ [Lỗi] {error_str}. Đang thử lại...", log_placeholder)
                time.sleep(2)
                
    add_log(f"🚫 [Bỏ qua] Không thể dịch đoạn này sau nhiều lần thử.", log_placeholder)
    return lead + core + trail

# ============================================================
# 5. CÁC HÀM XỬ LÝ FILE
# ============================================================
def process_excel(input_path, output_path, client, log_placeholder, progress_bar):
    wb = openpyxl.load_workbook(input_path)
    cells_to_translate = []
    
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and not str(cell.value).startswith("=") and str(cell.value).strip():
                    cells_to_translate.append(cell)
                    
    total = len(cells_to_translate)
    add_log(f"Phát hiện {total} ô dữ liệu. Thời gian dự kiến: ~{int(total * 1.5 / 60) + 1} phút.", log_placeholder)

    if total == 0:
        wb.save(output_path)
        progress_bar.progress(1.0)
        return

    for i, cell in enumerate(cells_to_translate):
        cell.value = translate_text(client, cell.value, log_placeholder)
        progress_bar.progress((i + 1) / total)

    wb.save(output_path)

def process_word(input_path, output_path, client, log_placeholder, progress_bar):
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
    add_log(f"Phát hiện {total} đoạn văn bản. Thời gian dự kiến: ~{int(total * 1.5 / 60) + 1} phút.", log_placeholder)

    if total == 0:
        doc.save(output_path)
        progress_bar.progress(1.0)
        return

    for i, run in enumerate(runs_to_translate):
        run.text = translate_text(client, run.text, log_placeholder)
        progress_bar.progress((i + 1) / total)

    doc.save(output_path)

def process_pdf(input_path, output_path, client, log_placeholder, progress_bar):
    add_log("Đang chuyển đổi PDF sang Word ẩn để bảo toàn Layout...", log_placeholder)
    temp_docx = input_path.replace(".pdf", "_temp.docx")
    
    cv = Converter(input_path)
    cv.convert(temp_docx)
    cv.close()
    progress_bar.progress(0.1)

    class ScaledProgressBar:
        def __init__(self, pb): self.pb = pb
        def progress(self, val): self.pb.progress(0.1 + (val * 0.9))
        
    scaled_pb = ScaledProgressBar(progress_bar)
    process_word(temp_docx, output_path, client, log_placeholder, scaled_pb)

    if os.path.exists(temp_docx):
        os.remove(temp_docx)

# ============================================================
# 6. GIAO DIỆN CHÍNH
# ============================================================
st.markdown('<p class="main-title">⚡ Hệ thống Dịch thuật Kỹ thuật THACO</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Auto-Discovery Model (Chống lỗi 404/400) • Siêu tốc & Bảo toàn Cấu trúc</p>', unsafe_allow_html=True)

col_left, col_right = st.columns([1.2, 2])

with col_left:
    st.markdown("### 📂 Tải lên Tài liệu")
    with st.container(border=True):
        st.success("✅ Hệ thống đã kết nối máy chủ dịch thuật nội bộ an toàn qua Cloudflare. Sẵn sàng xử lý.")
    uploaded_file = st.file_uploader("Kéo thả file Word, Excel, PDF vào đây", type=["docx", "xlsx", "pdf"])

with col_right:
    st.markdown("### 🖥️ Màn hình Giám sát (Live Console)")
    log_placeholder = st.empty()
    log_placeholder.markdown("<div class='log-container'>Hệ thống sẵn sàng...<br>API Key đã được mã hóa ở cấp độ mạng qua Cloudflare.</div>", unsafe_allow_html=True)
    
    progress_bar = st.empty()
    action_placeholder = st.empty()

    if uploaded_file:
        if action_placeholder.button("🚀 BẮT ĐẦU DỊCH", type="primary", use_container_width=True):
            st.session_state.logs = []
            add_log("Khởi động kết nối bảo mật qua Cloudflare...", log_placeholder)
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1].lower()) as tmp_in:
                tmp_in.write(uploaded_file.getbuffer())
                input_path = tmp_in.name

            file_ext = os.path.splitext(uploaded_file.name)[1].lower()
            output_ext = ".docx" if file_ext == ".pdf" else file_ext
            output_name = uploaded_file.name.replace(file_ext, f"_EN_Llama{output_ext}")
            output_path = input_path.replace(file_ext, f"_out{output_ext}")

            try:
                client = get_groq_client()
                
                # --- TÍNH NĂNG TỰ ĐỘNG DÒ TÌM MODEL ---
                add_log("Đang dò tìm danh sách Model khả dụng (Chống lỗi 404/400)...", log_placeholder)
                try:
                    available_models = [m.id for m in client.models.list().data]
                    # Ưu tiên các model mạnh nhất, nếu không có sẽ tự lùi về model thấp hơn
                    priorities = ["llama-3.3-70b-versatile", "llama-3.1-8b-instant", "mixtral-8x7b-32768"]
                    
                    selected_model = next((m for m in priorities if m in available_models), available_models[0])
                    st.session_state.selected_model = selected_model
                    add_log(f"✅ Đã kết nối tự động tới model: {selected_model}", log_placeholder)
                except Exception as e:
                    st.session_state.selected_model = "llama-3.1-8b-instant"
                    add_log("⚠️ Chạy model dự phòng: llama-3.1-8b-instant", log_placeholder)
                # --------------------------------------

                pb = progress_bar.progress(0.0)
                
                if file_ext == ".docx": process_word(input_path, output_path, client, log_placeholder, pb)
                elif file_ext == ".xlsx": process_excel(input_path, output_path, client, log_placeholder, pb)
                elif file_ext == ".pdf": process_pdf(input_path, output_path, client, log_placeholder, pb)
                    
                add_log("✅ HOÀN TẤT!", log_placeholder)
                action_placeholder.empty()
                
                with open(output_path, "rb") as f:
                    st.download_button(
                        label=f"⬇️ TẢI XUỐNG TÀI LIỆU HOÀN THIỆN ({output_name})",
                        data=f, file_name=output_name, mime="application/octet-stream",
                        type="primary", use_container_width=True
                    )
            except Exception as e:
                add_log(f"❌ LỖI HỆ THỐNG: {str(e)}", log_placeholder)
            finally:
                if os.path.exists(input_path): os.remove(input_path)
# app.py
# Hệ thống Dịch thuật Kỹ thuật bằng AI (Groq LLM qua Cloudflare Proxy)
# Hỗ trợ: Excel (.xlsx), Word (.docx), PowerPoint (.pptx), PDF (.pdf)
# Cơ chế JSON Batching: Dịch 25-30 ô/lần gọi AI - Chuẩn thuật ngữ Ô tô & THACO

import streamlit as st
import requests
import json
import time
import os
import tempfile
import re

# ----- Thư viện xử lý File -----
from docx import Document
import openpyxl
from pdf2docx import Converter
from pptx import Presentation

# ============================================================
# 1. CẤU HÌNH ĐƯỜNG DẪN CLOUDFLARE WORKER CỦA ANH
# ============================================================
DEFAULT_PROXY_URL = "https://mrluu-groq-proxy.mrscong4.workers.dev"

st.set_page_config(page_title="THACO AI Technical Translator", page_icon="⚡", layout="wide")

st.markdown("""
    <style>
    .main-title { font-size: 34px !important; font-weight: 700; color: #1E3A8A; margin-bottom: 0px; }
    .sub-title { font-size: 15px; color: #64748B; margin-bottom: 20px; }
    .log-container { background-color: #0F172A; color: #10B981; font-family: 'Courier New', monospace; padding: 15px; border-radius: 8px; height: 380px; overflow-y: auto; font-size: 13px; line-height: 1.6; }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# 2. SYSTEM PROMPT CHUYÊN SÂU CHO AI (NGỮ CẢNH THACO & BJM)
# ============================================================
SYSTEM_PROMPT = """You are an expert automotive and after-sales quality management translator at THACO AUTO.
Translate the values in the provided JSON object from Vietnamese into professional, accurate English.

CRITICAL RULES:
1. Context: Automotive spare parts, bus/truck components, warranty claims, quality inspection, executive presentations, and international shipping (e.g., partner BJM).
2. Domain terminology examples (for context understanding, adapt naturally to any technical phrase):
   - "STT" -> "No."
   - "Số BCKT" -> "Technical Report No."
   - "Hình thức vận chuyển" -> "Shipping Method"
   - "BJM xách tay" -> "Hand-carried by BJM"
   - "Xách tay/ hoặc gửi theo lô hàng xuất khẩu..." -> "Hand-carried / or shipped with export batch on..."
   - "Chưa có báo cáo" -> "No report yet"
   - "Mã linh kiện" / "Mã vật tư" -> "Part Number"
   - "Tên linh kiện" / "Tên vật tư" -> "Part Name"
   - "Tên tiếng anh" -> "English Name"
   - "Số lượng" -> "Quantity"
   - "ĐVT" -> "Unit" ("Cái" -> "Pcs", "Bộ" -> "Set", "Ống" -> "Pipe/Hose")
   - "Dài-rộng-cao (mm)" -> "L-W-H (mm)"
   - "Cân nặng (kg)" -> "Weight (kg)"
   - "Tổng" -> "Total"
   - Technical parts: "Sacvo ly hợp" -> "Clutch Servo", "Càng đẩy bi T" -> "Clutch Release Fork", "Bạc đạn bi T" -> "Clutch Release Bearing", "Bộ truyền động puly quạt gió" -> "Fan Pulley Drive Assembly", "Ống dẫn nước làm mát" -> "Coolant Hose".
3. Keep part codes, model numbers (e.g., 1005P320, CNA6128Y-1312010, JS 24V, LH, RH), and dates intact.
4. If a phrase is ALREADY in correct English (e.g., "LH Headlamp Module", "Clutch Servo"), keep it unchanged.
5. OUTPUT FORMAT: Return ONLY a valid JSON object with the EXACT same keys ("0", "1", "2", ...) mapped to their translated English strings. Do NOT add markdown blocks, explanations, or extra text."""

# ============================================================
# 3. HÀM KẾT NỐI CLOUDFLARE & TỰ ĐỘNG LẤY MODEL SỐNG
# ============================================================
if "logs" not in st.session_state:
    st.session_state.logs = []

def add_log(message: str, log_placeholder):
    time_str = time.strftime("%H:%M:%S")
    st.session_state.logs.append(f"[{time_str}] {message}")
    display_logs = "<br>".join(st.session_state.logs[-45:])
    log_placeholder.markdown(f"<div class='log-container'>{display_logs}</div>", unsafe_allow_html=True)

def fetch_live_chat_models(proxy_url: str, direct_key: str = ""):
    """Lấy danh sách các model Chat thực sự đang hoạt động trên Groq"""
    url = f"{proxy_url.rstrip('/')}/openai/v1/models" if not direct_key else "https://api.groq.com/openai/v1/models"
    headers = {"Content-Type": "application/json"}
    if direct_key:
        headers["Authorization"] = f"Bearer {direct_key.strip()}"
        
    resp = requests.get(url, headers=headers, timeout=15)
    if resp.status_code != 200:
        raise Exception(f"HTTP {resp.status_code}: {resp.text}")
        
    data = resp.json().get("data", [])
    all_ids = [m["id"] for m in data if m.get("active", True)]
    
    # Loại bỏ các model âm thanh (whisper/orpheus), bảo vệ (guard), nhúng (embed) hoặc vision
    exclude_keywords = ["whisper", "guard", "orpheus", "tts", "vision", "embed", " tool", "playai"]
    chat_models = [
        m for m in all_ids 
        if not any(ex in m.lower() for ex in exclude_keywords)
    ]
    
    # Sắp xếp ưu tiên các model lớn và thông minh lên đầu
    def sort_key(m_name):
        name = m_name.lower()
        if "qwen" in name or "70b" in name: return 0
        if "deepseek" in name or "llama-4" in name: return 1
        if "gemma" in name or "mixtral" in name: return 2
        if "8b" in name: return 3
        return 4
        
    chat_models.sort(key=sort_key)
    return chat_models

def call_ai_batch(batch_dict: dict, model_id: str, proxy_url: str, direct_key: str, log_placeholder) -> dict:
    """Gửi cả cụm 25-30 câu lên AI dịch trong 1 lần duy nhất bằng định dạng JSON"""
    url = f"{proxy_url.rstrip('/')}/openai/v1/chat/completions" if not direct_key else "https://api.groq.com/openai/v1/chat/completions"
    headers = {"Content-Type": "application/json"}
    if direct_key:
        headers["Authorization"] = f"Bearer {direct_key.strip()}"

    payload = {
        "model": model_id,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": json.dumps(batch_dict, ensure_ascii=False)}
        ],
        "temperature": 0.1,
        "response_format": {"type": "json_object"}
    }

    for attempt in range(3):
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=45)
            # Nếu model không hỗ trợ ép kiểu json_object thì bỏ tham số đó đi và gọi lại
            if resp.status_code == 400 and "response_format" in resp.text:
                payload.pop("response_format", None)
                resp = requests.post(url, headers=headers, json=payload, timeout=45)

            if resp.status_code == 429:
                add_log("⚠️ Chạm giới hạn tốc độ (429), đang chờ 6 giây...", log_placeholder)
                time.sleep(6)
                continue

            if resp.status_code != 200:
                raise Exception(f"HTTP {resp.status_code} - {resp.text[:200]}")

            content = resp.json()["choices"][0]["message"]["content"].strip()
            # Xử lý trường hợp model có thẻ <think>...</think> (như Qwen/DeepSeek) trước khi lấy JSON
            content = re.sub(r'<think>.*?</think>', '', content, flags=re.DOTALL).strip()
            
            json_match = re.search(r'\{.*\}', content, re.DOTALL)
            if json_match:
                content = json_match.group(0)
            return json.loads(content)

        except Exception as e:
            add_log(f"⚠️ Thử lại lần {attempt + 1} do lỗi: {str(e)}", log_placeholder)
            time.sleep(2)

    add_log("❌ Không thể dịch cụm này sau 3 lần thử, giữ nguyên văn bản gốc.", log_placeholder)
    return batch_dict

# ============================================================
# 4. LOGIC LỌC & DỊCH FILE GIỮ NGUYÊN ĐỊNH DẠNG
# ============================================================
def should_translate(text: str) -> bool:
    """Chỉ bỏ qua số thuần túy, công thức Excel hoặc mã phụ tùng thuần túy (VD: 1005P320, CNA6128Y-1312010)"""
    t = text.strip()
    if not t or t.startswith("="):
        return False
    # Nếu có bất kỳ chữ cái tiếng Việt có dấu nào -> CHẮC CHẮN PHẢI DỊCH
    if re.search(r'[àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴĐ]', t):
        return True
    # Những từ viết tắt không dấu tiếng Việt cần dịch
    if t.upper() in ["STT", "SL", "TONG"]:
        return True
    # Nếu chỉ gồm chữ in hoa, số, dấu gạch ngang/chấm/chéo (Mã phụ tùng, kích thước 390-275-225, ngày tháng) -> Bỏ qua
    if re.match(r'^[A-Z0-9\-_./&\s()]+$', t):
        return False
    return True

def translate_all_strings(texts: list, model_id: str, proxy_url: str, direct_key: str, log_placeholder, progress_bar) -> list:
    results = list(texts)
    
    # Gom các chuỗi trùng nhau (Deduplication) để tiết kiệm 50% lượng gọi AI và đồng nhất từ vựng
    unique_map = {} # clean_text -> list of original indices
    for idx, raw_text in enumerate(texts):
        if should_translate(raw_text):
            clean = raw_text.strip()
            unique_map.setdefault(clean, []).append(idx)

    unique_strings = list(unique_map.keys())
    total_unique = len(unique_strings)
    add_log(f"🔍 Phát hiện {len(texts)} thành phần văn bản -> Lọc ra {total_unique} câu/cụm từ duy nhất cần AI dịch.", log_placeholder)

    if total_unique == 0:
        progress_bar.progress(1.0)
        return results

    # Chia thành từng lô 25 câu để gửi AI
    BATCH_SIZE = 25
    translated_unique = {}
    
    for start_idx in range(0, total_unique, BATCH_SIZE):
        batch_slice = unique_strings[start_idx : start_idx + BATCH_SIZE]
        batch_payload = {str(i): s for i, s in enumerate(batch_slice)}
        
        sample_preview = ", ".join([f"'{s[:20]}'" for s in batch_slice[:3]])
        add_log(f"🤖 Đang gửi AI ({model_id}) dịch lô {start_idx + 1} -> {start_idx + len(batch_slice)} ({sample_preview}...)", log_placeholder)
        
        ai_response = call_ai_batch(batch_payload, model_id, proxy_url, direct_key, log_placeholder)
        
        for i, original_clean in enumerate(batch_slice):
            trans_val = ai_response.get(str(i), original_clean)
            if not isinstance(trans_val, str) or not trans_val.strip():
                trans_val = original_clean
            translated_unique[original_clean] = trans_val.strip()
            
        progress_bar.progress(min((start_idx + len(batch_slice)) / total_unique, 1.0))
        time.sleep(1.0) # Nghỉ nhẹ 1s giữa các lô

    # Điền trả lại vào đúng vị trí các ô ban đầu (giữ nguyên khoảng trắng đầu/cuối nếu có)
    for clean_str, indices in unique_map.items():
        trans_str = translated_unique.get(clean_str, clean_str)
        for idx in indices:
            orig = texts[idx]
            lead = orig[:len(orig) - len(orig.lstrip())]
            trail = orig[len(orig.rstrip()):]
            results[idx] = lead + trans_str + trail

    return results

# ============================================================
# 5. CÁC HÀM XỬ LÝ FILE (EXCEL, WORD, POWERPOINT, PDF)
# ============================================================
def process_excel(input_path, output_path, model_id, proxy_url, direct_key, log_placeholder, progress_bar):
    add_log("📂 Đang mở file Excel và quét các Sheet...", log_placeholder)
    wb = openpyxl.load_workbook(input_path)
    cells, texts = [], []
    
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip():
                    cells.append(cell)
                    texts.append(cell.value)
                    
    translated = translate_all_strings(texts, model_id, proxy_url, direct_key, log_placeholder, progress_bar)
    for cell, new_val in zip(cells, translated):
        cell.value = new_val
        
    add_log("💾 Đang lưu file Excel hoàn thiện (giữ nguyên định dạng gốc)...", log_placeholder)
    wb.save(output_path)

def process_word(input_path, output_path, model_id, proxy_url, direct_key, log_placeholder, progress_bar):
    add_log("📂 Đang mở file Word...", log_placeholder)
    doc = Document(input_path)
    runs, texts = [], []
    
    for p in doc.paragraphs:
        for r in p.runs:
            if r.text.strip():
                runs.append(r); texts.append(r.text)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        if r.text.strip():
                            runs.append(r); texts.append(r.text)
                            
    translated = translate_all_strings(texts, model_id, proxy_url, direct_key, log_placeholder, progress_bar)
    for r, new_val in zip(runs, translated):
        r.text = new_val
        
    add_log("💾 Đang lưu file Word...", log_placeholder)
    doc.save(output_path)

def extract_pptx_shape_runs(shape, runs_list, texts_list):
    """Hàm đệ quy quét toàn bộ Text Frame, Bảng (Table) và Group Shapes trong PowerPoint"""
    # 1. Nếu shape là một Group (nhóm nhiều hình lại với nhau) -> Quét đệ quy bên trong
    if shape.shape_type == 6: # MSO_SHAPE_TYPE.GROUP
        for sub_shape in shape.shapes:
            extract_pptx_shape_runs(sub_shape, runs_list, texts_list)
            
    # 2. Nếu shape có chứa khung văn bản (Text Box, AutoShape, Title...)
    if shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                if r.text and r.text.strip():
                    runs_list.append(r)
                    texts_list.append(r.text)
                    
    # 3. Nếu shape là Bảng biểu (Table) trên Slide
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for p in cell.text_frame.paragraphs:
                    for r in p.runs:
                        if r.text and r.text.strip():
                            runs_list.append(r)
                            texts_list.append(r.text)

def process_pptx(input_path, output_path, model_id, proxy_url, direct_key, log_placeholder, progress_bar):
    add_log("📂 Đang mở file PowerPoint và quét tất cả các Slide...", log_placeholder)
    prs = Presentation(input_path)
    runs, texts = [], []
    
    for slide_idx, slide in enumerate(prs.slides):
        # Quét toàn bộ đối tượng trên Slide (Shape, Table, Group...)
        for shape in slide.shapes:
            extract_pptx_shape_runs(shape, runs, texts)
            
        # Quét luôn phần Ghi chú của diễn giả (Speaker Notes) nếu có
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            for p in slide.notes_slide.notes_text_frame.paragraphs:
                for r in p.runs:
                    if r.text and r.text.strip():
                        runs.append(r)
                        texts.append(r.text)
                        
    add_log(f"📊 Đã quét {len(prs.slides)} Slide (gồm Textbox, Table, Group & Notes).", log_placeholder)
    translated = translate_all_strings(texts, model_id, proxy_url, direct_key, log_placeholder, progress_bar)
    
    for r, new_val in zip(runs, translated):
        r.text = new_val
        
    add_log("💾 Đang lưu file PowerPoint hoàn thiện (giữ nguyên 100% bố cục, màu sắc, font chữ)...", log_placeholder)
    prs.save(output_path)

def process_pdf(input_path, output_path, model_id, proxy_url, direct_key, log_placeholder, progress_bar):
    add_log("🔄 Đang chuyển PDF sang cấu trúc Word...", log_placeholder)
    temp_docx = input_path.replace(".pdf", "_temp.docx")
    cv = Converter(input_path)
    cv.convert(temp_docx)
    cv.close()
    progress_bar.progress(0.15)

    class ScaledPB:
        def __init__(self, pb): self.pb = pb
        def progress(self, val): self.pb.progress(0.15 + val * 0.85)
        
    process_word(temp_docx, output_path, model_id, proxy_url, direct_key, log_placeholder, ScaledPB(progress_bar))
    if os.path.exists(temp_docx):
        os.remove(temp_docx)

# ============================================================
# 6. GIAO DIỆN CHÍNH
# ============================================================
st.markdown('<p class="main-title">⚡ Hệ thống Dịch thuật AI Chuyên ngành Ô tô - THACO</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Hỗ trợ Excel, Word, PowerPoint, PDF • Dịch chuẩn ngữ cảnh Kỹ thuật & Phụ tùng • Giữ nguyên cấu trúc File</p>', unsafe_allow_html=True)

col_left, col_right = st.columns([1.25, 2])

with col_left:
    st.markdown("### ⚙️ 1. Kết nối Máy chủ AI")
    with st.container(border=True):
        proxy_url = st.text_input("🔗 Đường dẫn Cloudflare Worker:", value=DEFAULT_PROXY_URL)
        direct_key = st.text_input("🔑 Hoặc nhập trực tiếp Groq API Key (Tùy chọn nếu không qua Cloudflare):", type="password")
        
        if st.button("🔄 Kiểm tra kết nối & Tải danh sách Model AI", use_container_width=True):
            try:
                with st.spinner("Đang lấy danh sách model đang hoạt động..."):
                    models = fetch_live_chat_models(proxy_url, direct_key)
                    st.session_state.available_models = models
                    st.success(f"✅ Kết nối thành công! Tìm thấy {len(models)} model AI khả dụng.")
            except Exception as e:
                st.error(f"❌ Lỗi kết nối: {str(e)}")

        # Đưa qwen/qwen3.8-27b lên làm model mặc định đầu tiên
        model_list = st.session_state.get("available_models", [
            "qwen/qwen3.8-27b",
            "meta-llama/llama-4-scout-17b-16e-instruct",
            "meta-llama/llama-4-maverick-17b-128e-instruct",
            "qwen-qwq-32b",
            "deepseek-r1-distill-llama-70b",
            "gemma2-9b-it"
        ])
        selected_model = st.selectbox("🤖 Chọn Model AI dịch thuật:", options=model_list)

    st.markdown("### 📂 2. Tải lên Tài liệu")
    uploaded_file = st.file_uploader(
        "Kéo thả file Excel, Word, PowerPoint, PDF", 
        type=["xlsx", "docx", "pptx", "pdf"]
    )

with col_right:
    st.markdown("### 🖥️ Màn hình Giám sát (Live Console)")
    log_placeholder = st.empty()
    log_placeholder.markdown(
        "<div class='log-container'>Hệ thống AI sẵn sàng!<br>"
        "• Đã hỗ trợ: <b>Excel (.xlsx), Word (.docx), PowerPoint (.pptx), PDF (.pdf)</b>.<br>"
        "👉 Hướng dẫn: Kiểm tra kết nối Model AI -> Tải file lên -> Bấm <b>Bắt đầu dịch bằng AI</b>.</div>",
        unsafe_allow_html=True
    )
    
    progress_bar = st.empty()
    action_placeholder = st.empty()

    if uploaded_file:
        if action_placeholder.button("🚀 BẮT ĐẦU DỊCH BẰNG AI", type="primary", use_container_width=True):
            st.session_state.logs = []
            add_log(f"🚀 Khởi chạy tiến trình dịch AI với model: {selected_model}", log_placeholder)
            
            file_ext = os.path.splitext(uploaded_file.name)[1].lower()
            with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp_in:
                tmp_in.write(uploaded_file.getbuffer())
                input_path = tmp_in.name

            output_ext = ".docx" if file_ext == ".pdf" else file_ext
            output_name = uploaded_file.name.replace(file_ext, f"_AI_EN{output_ext}")
            output_path = input_path.replace(file_ext, f"_out{output_ext}")

            try:
                pb = progress_bar.progress(0.0)
                start_time = time.time()
                
                if file_ext == ".xlsx":
                    process_excel(input_path, output_path, selected_model, proxy_url, direct_key, log_placeholder, pb)
                elif file_ext == ".docx":
                    process_word(input_path, output_path, selected_model, proxy_url, direct_key, log_placeholder, pb)
                elif file_ext == ".pptx":
                    process_pptx(input_path, output_path, selected_model, proxy_url, direct_key, log_placeholder, pb)
                elif file_ext == ".pdf":
                    process_pdf(input_path, output_path, selected_model, proxy_url, direct_key, log_placeholder, pb)
                    
                elapsed = round(time.time() - start_time, 1)
                add_log(f"🎉 HOÀN TẤT 100% BẰNG AI TRONG {elapsed} GIÂY!", log_placeholder)
                action_placeholder.empty()
                
                with open(output_path, "rb") as f:
                    st.download_button(
                        label=f"⬇️ TẢI XUỐNG FILE ĐÃ DỊCH ({output_name})",
                        data=f, file_name=output_name, mime="application/octet-stream",
                        type="primary", use_container_width=True
                    )
            except Exception as e:
                add_log(f"❌ LỖI HỆ THỐNG: {str(e)}", log_placeholder)
            finally:
                if os.path.exists(input_path):
                    os.remove(input_path)
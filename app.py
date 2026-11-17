# app.py
# THACO AI Technical Translator PRO - Full Features & Robust Logic

import streamlit as st
import requests
import json
import time
import os
import tempfile
import re
import unicodedata

# ----- Thư viện xử lý File -----
from docx import Document
import openpyxl
from pdf2docx import Converter
from pptx import Presentation

DEFAULT_PROXY_URL = "https://mrluu-groq-proxy.mrscong4.workers.dev"
TM_FILE = "translation_memory.json"

st.set_page_config(page_title="Mr lưu AI Technical Translator PRO", page_icon="⚡", layout="wide")

st.markdown("""
    <style>
    .main-title { font-size: 32px !important; font-weight: 700; color: #1E3A8A; margin-bottom: 0px; }
    .sub-title { font-size: 14px; color: #64748B; margin-bottom: 15px; }
    .log-container { background-color: #0F172A; color: #10B981; font-family: 'Courier New', monospace; padding: 12px; border-radius: 8px; height: 320px; overflow-y: auto; font-size: 12px; line-height: 1.5; }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# 1. BỘ NHỚ ĐỆM DỊCH THUẬT (TRANSLATION MEMORY - TM)
# ============================================================
def load_tm() -> dict:
    if os.path.exists(TM_FILE):
        try:
            with open(TM_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_tm(tm: dict):
    try:
        with open(TM_FILE, "w", encoding="utf-8") as f:
            json.dump(tm, f, ensure_ascii=False, indent=2)
    except:
        pass

if "tm_data" not in st.session_state:
    st.session_state.tm_data = load_tm()

# ============================================================
# 2. SYSTEM PROMPT CHUYÊN SÂU NGÀNH Ô TÔ & QUẢN LÝ CHẤT LƯỢNG
# ============================================================
SYSTEM_PROMPT = """You are a senior automotive engineering and after-sales quality management translator at THACO AUTO.
Translate the text strings in the provided JSON object accurately and professionally.

DOMAIN CONTEXT & RULES:
1. Terminology: Use standard automotive manufacturing, 8D report, and after-sales terms (e.g., Root Cause, Corrective Action, Preventive Action, Part Number, Charging Port, Resistor, Voltage, Supplier, Warranty Claim).
2. Preservation: Keep all part numbers (e.g., USB-CDQ23035B, 1005P320), VIN numbers, vehicle models (TB120B, TB120S), quantities, percentages, dates, and person names completely intact.
3. Flow: Ensure the translation reads like a professional technical report rather than a word-for-word translation.
4. OUTPUT FORMAT: Return ONLY a valid JSON object with the exact same keys mapping to their translated strings. No markdown formatting blocks, no explanations."""

# ============================================================
# 3. KẾT NỐI AI & GỌI BATCH AN TOÀN (CHỐNG LỖI 429)
# ============================================================
if "logs" not in st.session_state:
    st.session_state.logs = []

def add_log(message: str, log_placeholder):
    time_str = time.strftime("%H:%M:%S")
    st.session_state.logs.append(f"[{time_str}] {message}")
    display_logs = "<br>".join(st.session_state.logs[-50:])
    log_placeholder.markdown(f"<div class='log-container'>{display_logs}</div>", unsafe_allow_html=True)

def fetch_live_chat_models(proxy_url: str, direct_key: str = ""):
    url = f"{proxy_url.rstrip('/')}/openai/v1/models" if not direct_key else "https://api.groq.com/openai/v1/models"
    headers = {"Content-Type": "application/json"}
    if direct_key:
        headers["Authorization"] = f"Bearer {direct_key.strip()}"
        
    resp = requests.get(url, headers=headers, timeout=15)
    if resp.status_code != 200:
        raise Exception(f"HTTP {resp.status_code}: {resp.text}")
        
    data = resp.json().get("data", [])
    all_ids = [m["id"] for m in data if m.get("active", True)]
    exclude = ["whisper", "guard", "orpheus", "tts", "vision", "embed", "tool", "playai"]
    chat_models = [m for m in all_ids if not any(ex in m.lower() for ex in exclude)]
    
    def sort_key(name):
        n = name.lower()
        if "qwen" in n or "70b" in n: return 0
        if "deepseek" in n or "llama-4" in n: return 1
        if "gemma" in n or "mixtral" in n: return 2
        return 3
    chat_models.sort(key=sort_key)
    return chat_models

def call_ai_batch(batch_dict: dict, model_id: str, proxy_url: str, direct_key: str, direction: str, log_placeholder) -> dict:
    url = f"{proxy_url.rstrip('/')}/openai/v1/chat/completions" if not direct_key else "https://api.groq.com/openai/v1/chat/completions"
    headers = {"Content-Type": "application/json"}
    if direct_key:
        headers["Authorization"] = f"Bearer {direct_key.strip()}"

    dir_prompt = "Translate from Vietnamese to English." if direction == "Việt → Anh" else "Translate from English to Vietnamese."
    full_system = SYSTEM_PROMPT + f"\nDirection: {dir_prompt}"

    payload = {
        "model": model_id,
        "messages": [
            {"role": "system", "content": full_system},
            {"role": "user", "content": json.dumps(batch_dict, ensure_ascii=False)}
        ],
        "temperature": 0.1,
        "response_format": {"type": "json_object"}
    }

    for attempt in range(4):
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=60)
            if resp.status_code == 400 and "response_format" in resp.text:
                payload.pop("response_format", None)
                resp = requests.post(url, headers=headers, json=payload, timeout=60)

            if resp.status_code == 429:
                wait_time = 8 * (attempt + 1)
                add_log(f"⚠️ Giới hạn tốc độ (429), đang chờ {wait_time} giây...", log_placeholder)
                time.sleep(wait_time)
                continue

            if resp.status_code != 200:
                raise Exception(f"HTTP {resp.status_code} - {resp.text[:150]}")

            content = resp.json()["choices"][0]["message"]["content"].strip()
            content = re.sub(r'<think>.*?</think>', '', content, flags=re.DOTALL).strip()
            
            match = re.search(r'\{.*\}', content, re.DOTALL)
            if match:
                content = match.group(0)
            return json.loads(content)

        except Exception as e:
            add_log(f"⚠️ Thử lại lần {attempt + 1}: {str(e)}", log_placeholder)
            time.sleep(3)

    return batch_dict

# ============================================================
# 4. BỘ LỌC ĐOẠN VĂN & XỬ LÝ DỊCH
# ============================================================
def should_translate(text: str, direction: str) -> bool:
    t = text.strip()
    if not t or t.startswith("="):
        return False
    if not any(ch.isalpha() for ch in t):
        return False
    if direction == "Việt → Anh":
        if re.search(r'[àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđĐ]', t):
            return True
        if t.upper() in ["STT", "SL", "TONG", "NGUYEN NHAN", "GIAI PHAP", "MO TA"]:
            return True
        if re.match(r'^[A-Z0-9\-_./&#\s():%]+$', t):
            return False
        return True
    else:
        if re.match(r'^[0-9\-_./&#\s():%]+$', t):
            return False
        return True

def process_translation_batch(texts: list, model_id: str, proxy_url: str, direct_key: str, direction: str, log_placeholder, progress_bar) -> list:
    results = list(texts)
    tm = st.session_state.tm_data
    
    unique_map = {}
    for idx, raw in enumerate(texts):
        if should_translate(raw, direction):
            clean = raw.strip()
            unique_map.setdefault(clean, []).append(idx)

    uniques = list(unique_map.keys())
    to_translate = []
    translated_cache = {}

    for u in uniques:
        if u in tm:
            translated_cache[u] = tm[u]
        else:
            to_translate.append(u)

    add_log(f"🔍 Tổng số đoạn: {len(texts)} | Lấy từ Bộ nhớ đệm (TM): {len(uniques) - len(to_translate)} | Cần AI dịch: {len(to_translate)}", log_placeholder)

    if to_translate:
        BATCH_SIZE = 12  # Lô nhỏ gọn để đảm bảo an toàn tuyệt đối qua Cloudflare
        for start in range(0, len(to_translate), BATCH_SIZE):
            slice_items = to_translate[start : start + BATCH_SIZE]
            payload = {str(i): s for i, s in enumerate(slice_items)}
            
            add_log(f"🤖 Đang gửi AI dịch lô {start + 1} -> {start + len(slice_items)}...", log_placeholder)
            res = call_ai_batch(payload, model_id, proxy_url, direct_key, direction, log_placeholder)
            
            for i, orig in enumerate(slice_items):
                val = res.get(str(i), orig)
                if isinstance(val, str) and val.strip():
                    translated_cache[orig] = val.strip()
                    tm[orig] = val.strip()
                else:
                    translated_cache[orig] = orig
            
            save_tm(tm)
            progress_bar.progress(min((start + len(slice_items)) / len(to_translate), 1.0))
            time.sleep(1.5)

    for clean_str, indices in unique_map.items():
        trans = translated_cache.get(clean_str, clean_str)
        for idx in indices:
            orig = texts[idx]
            lead = orig[:len(orig) - len(orig.lstrip())]
            trail = orig[len(orig.rstrip()):]
            results[idx] = lead + trans + trail

    return results

# ============================================================
# 5. XỬ LÝ ĐỊNH DẠNG FILE
# ============================================================
class ParagraphSegment:
    def __init__(self, runs):
        self.runs = runs
        self.text = "".join(r.text for r in runs if r.text)

    def apply_result(self, new_text: str, bilingual: bool):
        if not self.runs:
            return
        final_text = f"{self.text}\n({new_text})" if bilingual else new_text
        self.runs[0].text = final_text
        for r in self.runs[1:]:
            r.text = ""

def collect_segments(p, segs, texts):
    runs = [r for r in p.runs if r.text is not None]
    if not runs:
        return
    full = "".join(r.text for r in runs)
    if not full.strip():
        return
    seg = ParagraphSegment(runs)
    segs.append(seg)
    texts.append(seg.text)

def process_excel(input_path, output_path, model_id, proxy_url, direct_key, direction, bilingual, log_placeholder, progress_bar):
    add_log("📂 Đang quét Sheets & Tabs Excel...", log_placeholder)
    wb = openpyxl.load_workbook(input_path)
    
    # Dịch tên các tab Sheet dưới chân file Excel
    sheet_titles = wb.sheetnames
    trans_titles = process_translation_batch(sheet_titles, model_id, proxy_url, direct_key, direction, log_placeholder, progress_bar)
    for old_t, new_t in zip(sheet_titles, trans_titles):
        wb[old_t].title = new_t

    cells, texts = [], []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip():
                    cells.append(cell)
                    texts.append(cell.value)
                    cell.alignment = openpyxl.styles.Alignment(wrap_text=True)

    translated = process_translation_batch(texts, model_id, proxy_url, direct_key, direction, log_placeholder, progress_bar)
    for cell, orig, new_val in zip(cells, texts, translated):
        cell.value = f"{orig}\n({new_val})" if bilingual else new_val

    wb.save(output_path)

def process_word(input_path, output_path, model_id, proxy_url, direct_key, direction, bilingual, log_placeholder, progress_bar):
    add_log("📂 Đang mở file Word...", log_placeholder)
    doc = Document(input_path)
    segs, texts = [], []
    for p in doc.paragraphs:
        collect_segments(p, segs, texts)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    collect_segments(p, segs, texts)

    translated = process_translation_batch(texts, model_id, proxy_url, direct_key, direction, log_placeholder, progress_bar)
    for seg, new_val in zip(segs, translated):
        seg.apply_result(new_val, bilingual)

    doc.save(output_path)

def extract_pptx(shape, segs, texts):
    if shape.shape_type == 6:
        for sub in shape.shapes:
            extract_pptx(sub, segs, texts)
    if shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            collect_segments(p, segs, texts)
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for p in cell.text_frame.paragraphs:
                    collect_segments(p, segs, texts)

def process_pptx(input_path, output_path, model_id, proxy_url, direct_key, direction, bilingual, log_placeholder, progress_bar):
    add_log("📂 Đang mở file PowerPoint (quét Shapes, Tables, Groups, Notes)...", log_placeholder)
    prs = Presentation(input_path)
    segs, texts = [], []
    for slide in prs.slides:
        for shape in slide.shapes:
            extract_pptx(shape, segs, texts)
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            for p in slide.notes_slide.notes_text_frame.paragraphs:
                collect_segments(p, segs, texts)

    translated = process_translation_batch(texts, model_id, proxy_url, direct_key, direction, log_placeholder, progress_bar)
    for seg, new_val in zip(segs, translated):
        seg.apply_result(new_val, bilingual)

    prs.save(output_path)

def process_pdf(input_path, output_path, model_id, proxy_url, direct_key, direction, bilingual, log_placeholder, progress_bar):
    add_log("🔄 Đang chuyển PDF sang Word...", log_placeholder)
    temp_docx = input_path.replace(".pdf", "_temp.docx")
    cv = Converter(input_path)
    cv.convert(temp_docx)
    cv.close()
    progress_bar.progress(0.15)
    class ScaledPB:
        def __init__(self, pb): self.pb = pb
        def progress(self, val): self.pb.progress(0.15 + val * 0.85)
    process_word(temp_docx, output_path, model_id, proxy_url, direct_key, direction, bilingual, log_placeholder, ScaledPB(progress_bar))
    if os.path.exists(temp_docx):
        os.remove(temp_docx)

# ============================================================
# 6. GIAO DIỆN STREAMLIT PRO
# ============================================================
st.markdown('<p class="main-title">⚡ THACO AI Technical Translator PRO</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Chống lỗi 429 • Song Ngữ • Dịch Đa Chiều • Bộ nhớ đệm TM • Auto-Fit Layout</p>', unsafe_allow_html=True)

col_left, col_right = st.columns([1.3, 2])

with col_left:
    st.markdown("### ⚙️ 1. Cấu hình AI & Tùy chọn")
    with st.container(border=True):
        proxy_url = st.text_input("🔗 Đường dẫn Cloudflare Worker:", value=DEFAULT_PROXY_URL)
        direct_key = st.text_input("🔑 Hoặc Groq API Key trực tiếp:", type="password")
        
        if st.button("🔄 Tải danh sách Model AI", use_container_width=True):
            try:
                models = fetch_live_chat_models(proxy_url, direct_key)
                st.session_state.available_models = models
                st.success(f"✅ Tải thành công {len(models)} model!")
            except Exception as e:
                st.error(f"❌ Lỗi: {str(e)}")

        model_list = st.session_state.get("available_models", ["qwen/qwen3.8-27b", "deepseek-r1-distill-llama-70b", "gemma2-9b-it"])
        selected_model = st.selectbox("🤖 Chọn Model AI:", options=model_list)

        direction = st.radio("🌐 Chiều dịch:", ["Việt → Anh", "Anh → Việt"], horizontal=True)
        bilingual = st.checkbox("📖 Xuất Song Ngữ (Nối tiếp dòng dưới)", value=False)

    st.markdown("### 📂 2. Tải lên Tài liệu")
    uploaded_file = st.file_uploader("Hỗ trợ Excel, Word, PowerPoint, PDF", type=["xlsx", "docx", "pptx", "pdf"])

with col_right:
    st.markdown("### 🖥️ Màn hình Giám sát & Nhật ký")
    log_placeholder = st.empty()
    log_placeholder.markdown("<div class='log-container'>Hệ thống sẵn sàng bảo mật qua Cloudflare Worker.</div>", unsafe_allow_html=True)
    
    progress_bar = st.empty()
    action_placeholder = st.empty()

    if uploaded_file:
        if action_placeholder.button("🚀 BẮT ĐẦU DỊCH & XỬ LÝ", type="primary", use_container_width=True):
            st.session_state.logs = []
            add_log(f"🚀 Khởi chạy với model {selected_model} ({direction})", log_placeholder)
            
            file_ext = os.path.splitext(uploaded_file.name)[1].lower()
            with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp_in:
                tmp_in.write(uploaded_file.getbuffer())
                input_path = tmp_in.name

            output_ext = ".docx" if file_ext == ".pdf" else file_ext
            output_name = uploaded_file.name.replace(file_ext, f"_Pro_EN{output_ext}")
            output_path = input_path.replace(file_ext, f"_out{output_ext}")

            try:
                pb = progress_bar.progress(0.0)
                start_t = time.time()
                
                if file_ext == ".xlsx":
                    process_excel(input_path, output_path, selected_model, proxy_url, direct_key, direction, bilingual, log_placeholder, pb)
                elif file_ext == ".docx":
                    process_word(input_path, output_path, selected_model, proxy_url, direct_key, direction, bilingual, log_placeholder, pb)
                elif file_ext == ".pptx":
                    process_pptx(input_path, output_path, selected_model, proxy_url, direct_key, direction, bilingual, log_placeholder, pb)
                elif file_ext == ".pdf":
                    process_pdf(input_path, output_path, selected_model, proxy_url, direct_key, direction, bilingual, log_placeholder, pb)

                elapsed = round(time.time() - start_t, 1)
                add_log(f"🎉 HOÀN TẤT TRONG {elapsed} GIÂY!", log_placeholder)
                action_placeholder.empty()

                with open(output_path, "rb") as f:
                    st.download_button(
                        label=f"⬇️ TẢI XUỐNG FILE HOÀN THIỆN ({output_name})",
                        data=f, file_name=output_name, mime="application/octet-stream",
                        type="primary", use_container_width=True
                    )
            except Exception as e:
                add_log(f"❌ LỖI: {str(e)}", log_placeholder)
            finally:
                if os.path.exists(input_path):
                    os.remove(input_path)
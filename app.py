# app.py
# Hệ thống Dịch thuật AI Chuyên ngành Ô tô - THACO (Phiên bản Enterprise Pro - Safe Boot)
# Chống lỗi sập App ("Oh no") + Tích hợp đầy đủ 6 tính năng nâng cao

import streamlit as st
import requests
import json
import time
import os
import tempfile
import re
import pandas as pd

# ----- Safe Import các thư viện xử lý File (Chống crash app nếu server chưa kịp cài) -----
from docx import Document
import openpyxl
from openpyxl.styles import Alignment

try:
    from pdf2docx import Converter
    HAS_PDF = True
except ImportError:
    HAS_PDF = False

try:
    from pptx import Presentation
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

# ============================================================
# 1. CẤU HÌNH MẶC ĐỊNH & GIAO DIỆN
# ============================================================
DEFAULT_PROXY_URL = "https://mrluu-groq-proxy.mrscong4.workers.dev"
TM_FILE_PATH = "thaco_translation_memory.json"

st.set_page_config(page_title="THACO AI Technical Translator Pro", page_icon="⚡", layout="wide")

st.markdown("""
    <style>
    .main-title { font-size: 32px !important; font-weight: 700; color: #1E3A8A; margin-bottom: 0px; }
    .sub-title { font-size: 15px; color: #64748B; margin-bottom: 15px; }
    .log-container { background-color: #0F172A; color: #10B981; font-family: 'Courier New', monospace; padding: 14px; border-radius: 8px; height: 320px; overflow-y: auto; font-size: 13px; line-height: 1.5; }
    </style>
""", unsafe_allow_html=True)

# Khởi tạo các biến Session State
if "logs" not in st.session_state:
    st.session_state.logs = []
if "translation_memory" not in st.session_state:
    if os.path.exists(TM_FILE_PATH):
        try:
            with open(TM_FILE_PATH, "r", encoding="utf-8") as f:
                st.session_state.translation_memory = json.load(f)
        except Exception:
            st.session_state.translation_memory = {}
    else:
        st.session_state.translation_memory = {}
if "review_df" not in st.session_state:
    st.session_state.review_df = None
if "file_bytes" not in st.session_state:
    st.session_state.file_bytes = None
if "file_name" not in st.session_state:
    st.session_state.file_name = None
if "translation_done" not in st.session_state:
    st.session_state.translation_done = False

def save_tm_to_disk():
    try:
        with open(TM_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(st.session_state.translation_memory, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def add_log(message: str, log_placeholder=None):
    time_str = time.strftime("%H:%M:%S")
    st.session_state.logs.append(f"[{time_str}] {message}")
    if log_placeholder is not None:
        display_logs = "<br>".join(st.session_state.logs[-45:])
        log_placeholder.markdown(f"<div class='log-container'>{display_logs}</div>", unsafe_allow_html=True)

# ============================================================
# 2. SYSTEM PROMPT ĐA CHIỀU (VIỆT <-> ANH) CHUẨN THACO
# ============================================================
def get_system_prompt(direction: str) -> str:
    if direction == "Việt → Anh (VI → EN)":
        return """You are an expert automotive and after-sales quality management translator at THACO AUTO.
Translate the values in the provided JSON object from Vietnamese into professional, accurate English.

CRITICAL RULES:
1. Context: Automotive manufacturing, passenger cars, trucks, buses, spare parts, warranty claims, 8D reports, PFMEA, IATF 16949, and international shipping (e.g., partner BJM).
2. Standard THACO terminology:
   - "STT" -> "No.", "Số BCKT" -> "Technical Report No.", "BCKT" -> "Technical Report"
   - "Hình thức vận chuyển" -> "Shipping Method", "BJM xách tay" -> "Hand-carried by BJM"
   - "Chưa có báo cáo" -> "No report yet", "Mã linh kiện" / "Mã vật tư" -> "Part Number"
   - "Tên linh kiện" / "Tên vật tư" -> "Part Name", "Tên tiếng anh" -> "English Name"
   - "Số lượng" / "SL" -> "Quantity", "ĐVT" -> "Unit" ("Cái" -> "Pcs", "Bộ" -> "Set", "Ống" -> "Pipe/Hose")
   - "Dài-rộng-cao (mm)" -> "L-W-H (mm)", "Cân nặng (kg)" -> "Weight (kg)", "Tổng" -> "Total"
   - Technical parts: "Sacvo ly hợp" -> "Clutch Servo", "Càng đẩy bi T" -> "Clutch Release Fork", "Bạc đạn bi T" -> "Clutch Release Bearing", "Bộ truyền động puly quạt gió" -> "Fan Pulley Drive Assembly", "Ống dẫn nước làm mát" -> "Coolant Hose".
3. Keep part codes, model numbers (e.g., 1005P320, CNA6128Y-1312010, JS 24V, LH, RH), and dates intact.
4. If a phrase is ALREADY in correct English, keep it unchanged.
5. OUTPUT FORMAT: Return ONLY a valid JSON object with the EXACT same keys ("0", "1", ...) mapped to their translated English strings. No markdown, no explanations."""
    else:
        return """You are an expert automotive and after-sales quality management translator at THACO AUTO.
Translate the values in the provided JSON object from English into professional, accurate Vietnamese used in Vietnam's automotive industry.

CRITICAL RULES:
1. Context: Automotive engineering, spare parts, warranty claims, 8D/5-Why root cause analysis, quality inspection at THACO AUTO.
2. Standard THACO terminology:
   - "No." -> "STT", "Technical Report No." -> "Số BCKT", "Part Number" -> "Mã linh kiện"
   - "Part Name" -> "Tên linh kiện", "Quantity" / "Qty" -> "Số lượng", "Unit" -> "ĐVT" ("Pcs" -> "Cái", "Set" -> "Bộ")
   - "Root Cause" -> "Nguyên nhân gốc rễ", "Corrective Action" -> "Hành động khắc phục", "Preventive Action" -> "Hành động phòng ngừa"
   - "Clutch Servo" -> "Sacvo ly hợp", "Clutch Release Fork" -> "Càng đẩy bi T", "Clutch Release Bearing" -> "Bạc đạn bi T", "Coolant Hose" -> "Ống dẫn nước làm mát".
3. Keep part numbers, vehicle codes (CKD, CBU, LH, RH), and numbers intact.
4. If a phrase is ALREADY in Vietnamese, keep it unchanged.
5. OUTPUT FORMAT: Return ONLY a valid JSON object with the EXACT same keys ("0", "1", ...) mapped to their translated Vietnamese strings. No markdown, no explanations."""

# ============================================================
# 3. KẾT NỐI AI & BỘ LỌC NHẬN DIỆN THÔNG MINH
# ============================================================
def fetch_live_chat_models(proxy_url: str, direct_key: str = ""):
    url = f"{proxy_url.rstrip('/')}/openai/v1/models" if not direct_key else "https://api.groq.com/openai/v1/models"
    headers = {"Content-Type": "application/json"}
    if direct_key:
        headers["Authorization"] = f"Bearer {direct_key.strip()}"
        
    resp = requests.get(url, headers=headers, timeout=15)
    if resp.status_code != 200:
        raise Exception(f"HTTP {resp.status_code}: {resp.text[:150]}")
        
    data = resp.json().get("data", [])
    all_ids = [m["id"] for m in data if m.get("active", True)]
    exclude_kw = ["whisper", "guard", "orpheus", "tts", "vision", "embed", "tool", "playai"]
    chat_models = [m for m in all_ids if not any(ex in m.lower() for ex in exclude_kw)]
    
    def sort_key(m_name):
        name = m_name.lower()
        if "qwen" in name: return 0
        if "70b" in name: return 1
        if "llama-4" in name or "deepseek" in name: return 2
        return 3
        
    chat_models.sort(key=sort_key)
    return chat_models

def is_part_code_or_number(t: str) -> bool:
    if re.match(r'^[\d\s.,\-+/&:()]+$', t):
        return True
    if " " not in t and any(c.isdigit() for c in t) and re.match(r'^[A-Za-z0-9\-_./]+$', t):
        return True
    return False

def should_translate(text: str, direction: str) -> bool:
    t = text.strip()
    if not t or t.startswith("=") or is_part_code_or_number(t):
        return False
    if not re.search(r'[A-Za-zÀ-ỹ]', t):
        return False

    if direction == "Việt → Anh (VI → EN)":
        if re.search(r'[àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴĐ]', t):
            return True
        if t.upper() in ["STT", "SL", "TONG", "DVT", "BCKT"]:
            return True
        return False
    else:
        if t.upper() in ["LH", "RH", "ID", "OK", "VIN", "CKD", "CBU"]:
            return False
        return True

def call_ai_batch(batch_dict: dict, model_id: str, direction: str, proxy_url: str, direct_key: str, log_placeholder) -> dict:
    url = f"{proxy_url.rstrip('/')}/openai/v1/chat/completions" if not direct_key else "https://api.groq.com/openai/v1/chat/completions"
    headers = {"Content-Type": "application/json"}
    if direct_key:
        headers["Authorization"] = f"Bearer {direct_key.strip()}"

    payload = {
        "model": model_id,
        "messages": [
            {"role": "system", "content": get_system_prompt(direction)},
            {"role": "user", "content": json.dumps(batch_dict, ensure_ascii=False)}
        ],
        "temperature": 0.1,
        "response_format": {"type": "json_object"}
    }

    for attempt in range(3):
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=50)
            if resp.status_code == 400 and "response_format" in resp.text:
                payload.pop("response_format", None)
                resp = requests.post(url, headers=headers, json=payload, timeout=50)

            if resp.status_code == 429:
                add_log("⚠️ Chạm giới hạn tốc độ (429), chờ 5 giây...", log_placeholder)
                time.sleep(5)
                continue

            if resp.status_code != 200:
                raise Exception(f"HTTP {resp.status_code}: {resp.text[:150]}")

            content = resp.json()["choices"][0]["message"]["content"].strip()
            content = re.sub(r'<think>.*?</think>', '', content, flags=re.DOTALL).strip()
            json_match = re.search(r'\{.*\}', content, re.DOTALL)
            if json_match:
                content = json_match.group(0)
            return json.loads(content)

        except Exception as e:
            add_log(f"⚠️ Thử lại lần {attempt + 1} ({str(e)})...", log_placeholder)
            time.sleep(2)

    add_log("❌ Lô này gặp sự cố mạng, giữ nguyên văn bản gốc.", log_placeholder)
    return batch_dict

# ============================================================
# 4. QUÉT FILE & TẠO BẢNG DỊCH THUẬT
# ============================================================
def collect_docx_texts(doc, out_list: list):
    for p in doc.paragraphs:
        for r in p.runs:
            if r.text and r.text.strip():
                out_list.append(r.text.strip())
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        if r.text and r.text.strip():
                            out_list.append(r.text.strip())
    for txbx in doc.element.body.iter():
        if txbx.tag.endswith('txbxContent'):
            for p_elem in txbx.iter():
                if p_elem.tag.endswith('t') and p_elem.text and p_elem.text.strip():
                    out_list.append(p_elem.text.strip())

def collect_pptx_shape_texts(shape, out_list: list):
    if shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                if r.text and r.text.strip():
                    out_list.append(r.text.strip())
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for p in cell.text_frame.paragraphs:
                    for r in p.runs:
                        if r.text and r.text.strip():
                            out_list.append(r.text.strip())
    if shape.shape_type == 6 and hasattr(shape, "shapes"):
        for sub_shape in shape.shapes:
            collect_pptx_shape_texts(sub_shape, out_list)

def extract_unique_strings_from_file(input_path: str, file_ext: str, direction: str, translate_sheets: bool) -> list:
    raw_candidates = []
    if file_ext == ".xlsx":
        wb = openpyxl.load_workbook(input_path)
        for ws in wb.worksheets:
            if translate_sheets and ws.title:
                raw_candidates.append(ws.title.strip())
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.strip():
                        raw_candidates.append(cell.value.strip())
    elif file_ext == ".docx":
        doc = Document(input_path)
        collect_docx_texts(doc, raw_candidates)
    elif file_ext == ".pdf":
        if not HAS_PDF:
            raise Exception("Thư viện pdf2docx chưa được cài đặt trên máy chủ.")
        temp_docx = input_path.replace(".pdf", "_scan.docx")
        cv = Converter(input_path)
        cv.convert(temp_docx)
        cv.close()
        doc = Document(temp_docx)
        collect_docx_texts(doc, raw_candidates)
        if os.path.exists(temp_docx):
            os.remove(temp_docx)
    elif file_ext == ".pptx":
        if not HAS_PPTX:
            raise Exception("Thư viện python-pptx chưa được cài đặt. Hãy kiểm tra requirements.txt.")
        prs = Presentation(input_path)
        for slide in prs.slides:
            for shape in slide.shapes:
                collect_pptx_shape_texts(shape, raw_candidates)

    seen = set()
    unique_valid = []
    for s in raw_candidates:
        if s not in seen and should_translate(s, direction):
            seen.add(s)
            unique_valid.append(s)
    return unique_valid

def build_translation_map(unique_strings: list, model_id: str, direction: str, proxy_url: str, direct_key: str, use_tm: bool, log_placeholder, progress_bar) -> dict:
    trans_map = {}
    to_ask_ai = []
    tm_dict = st.session_state.translation_memory

    for s in unique_strings:
        tm_key = f"{direction}||{s}"
        if use_tm and tm_key in tm_dict:
            trans_map[s] = tm_dict[tm_key]
        else:
            to_ask_ai.append(s)

    tm_hits = len(unique_strings) - len(to_ask_ai)
    if tm_hits > 0:
        add_log(f"⚡ Lấy tức thì {tm_hits}/{len(unique_strings)} cụm từ từ Bộ nhớ dịch thuật (TM).", log_placeholder)

    total_ai = len(to_ask_ai)
    if total_ai == 0:
        progress_bar.progress(1.0)
        return trans_map

    add_log(f"🤖 Cần gửi AI dịch {total_ai} cụm từ mới...", log_placeholder)
    BATCH_SIZE = 25
    for start_idx in range(0, total_ai, BATCH_SIZE):
        batch_slice = to_ask_ai[start_idx : start_idx + BATCH_SIZE]
        batch_payload = {str(i): s for i, s in enumerate(batch_slice)}
        
        preview = ", ".join([f"'{x[:18]}'" for x in batch_slice[:3]])
        add_log(f"🚀 Đang dịch lô {start_idx + 1} -> {start_idx + len(batch_slice)} ({preview}...)", log_placeholder)
        
        ai_resp = call_ai_batch(batch_payload, model_id, direction, proxy_url, direct_key, log_placeholder)
        
        for i, orig_str in enumerate(batch_slice):
            val = ai_resp.get(str(i), orig_str)
            clean_val = val.strip() if isinstance(val, str) and val.strip() else orig_str
            trans_map[orig_str] = clean_val
            st.session_state.translation_memory[f"{direction}||{orig_str}"] = clean_val
            
        progress_bar.progress(min((start_idx + len(batch_slice)) / total_ai, 1.0))
        time.sleep(0.8)

    save_tm_to_disk()
    return trans_map

# ============================================================
# 5. KẾT XUẤT FILE HOÀN THIỆN
# ============================================================
def format_final_text(original_raw: str, trans_map: dict, bilingual: bool) -> str:
    clean = original_raw.strip()
    if clean not in trans_map:
        return original_raw
    translated = trans_map[clean]
    lead = original_raw[:len(original_raw) - len(original_raw.lstrip())]
    trail = original_raw[len(original_raw.rstrip()):]
    
    if bilingual and clean.lower() != translated.lower():
        return f"{lead}{clean}\n({translated}){trail}"
    return f"{lead}{translated}{trail}"

def sanitize_sheet_title(title: str, existing_titles: list) -> str:
    clean = re.sub(r'[\\/*?:\[\]]', '-', title).strip()
    if not clean:
        clean = "Sheet"
    base = clean[:28]
    candidate = clean[:31]
    counter = 1
    while candidate.lower() in [t.lower() for t in existing_titles]:
        candidate = f"{base}_{counter}"
        counter += 1
    return candidate

def apply_to_excel(input_path: str, output_path: str, trans_map: dict, bilingual: bool, translate_sheets: bool, auto_fit: bool):
    wb = openpyxl.load_workbook(input_path)
    
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip() in trans_map:
                    cell.value = format_final_text(cell.value, trans_map, bilingual)
                    if auto_fit or bilingual:
                        old_align = cell.alignment
                        if old_align:
                            cell.alignment = Alignment(
                                horizontal=old_align.horizontal,
                                vertical=old_align.vertical or "center",
                                wrap_text=True
                            )
                        else:
                            cell.alignment = Alignment(wrap_text=True, vertical="center")
                            
        if bilingual or auto_fit:
            for row_idx in range(1, ws.max_row + 1):
                ws.row_dimensions[row_idx].height = None

    if translate_sheets:
        used_titles = [ws.title for ws in wb.worksheets]
        for ws in wb.worksheets:
            orig_title = ws.title.strip()
            if orig_title in trans_map:
                new_raw = trans_map[orig_title]
                other_titles = [t for t in used_titles if t != ws.title]
                safe_title = sanitize_sheet_title(new_raw, other_titles)
                used_titles.remove(ws.title)
                ws.title = safe_title
                used_titles.append(safe_title)

    wb.save(output_path)

def apply_to_word(input_path: str, output_path: str, trans_map: dict, bilingual: bool):
    doc = Document(input_path)
    for p in doc.paragraphs:
        for r in p.runs:
            if r.text and r.text.strip() in trans_map:
                r.text = format_final_text(r.text, trans_map, bilingual)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        if r.text and r.text.strip() in trans_map:
                            r.text = format_final_text(r.text, trans_map, bilingual)
    for txbx in doc.element.body.iter():
        if txbx.tag.endswith('txbxContent'):
            for p_elem in txbx.iter():
                if p_elem.tag.endswith('t') and p_elem.text and p_elem.text.strip() in trans_map:
                    p_elem.text = format_final_text(p_elem.text, trans_map, bilingual)
    doc.save(output_path)

def apply_pptx_shape(shape, trans_map: dict, bilingual: bool):
    if shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                if r.text and r.text.strip() in trans_map:
                    r.text = format_final_text(r.text, trans_map, bilingual)
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for p in cell.text_frame.paragraphs:
                    for r in p.runs:
                        if r.text and r.text.strip() in trans_map:
                            r.text = format_final_text(r.text, trans_map, bilingual)
    if shape.shape_type == 6 and hasattr(shape, "shapes"):
        for sub in shape.shapes:
            apply_pptx_shape(sub, trans_map, bilingual)

def apply_to_pptx(input_path: str, output_path: str, trans_map: dict, bilingual: bool):
    prs = Presentation(input_path)
    for slide in prs.slides:
        for shape in slide.shapes:
            apply_pptx_shape(shape, trans_map, bilingual)
    prs.save(output_path)

def generate_final_output_bytes(file_bytes: bytes, file_ext: str, trans_map: dict, bilingual: bool, translate_sheets: bool, auto_fit: bool) -> bytes:
    with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp_in:
        tmp_in.write(file_bytes)
        in_path = tmp_in.name

    out_ext = ".docx" if file_ext == ".pdf" else file_ext
    out_path = in_path.replace(file_ext, f"_final{out_ext}")

    try:
        if file_ext == ".xlsx":
            apply_to_excel(in_path, out_path, trans_map, bilingual, translate_sheets, auto_fit)
        elif file_ext == ".docx":
            apply_to_word(in_path, out_path, trans_map, bilingual)
        elif file_ext == ".pdf":
            temp_docx = in_path.replace(".pdf", "_temp.docx")
            cv = Converter(in_path)
            cv.convert(temp_docx)
            cv.close()
            apply_to_word(temp_docx, out_path, trans_map, bilingual)
            if os.path.exists(temp_docx):
                os.remove(temp_docx)
        elif file_ext == ".pptx":
            apply_to_pptx(in_path, out_path, trans_map, bilingual)

        with open(out_path, "rb") as f:
            result_bytes = f.read()
        return result_bytes
    finally:
        if os.path.exists(in_path): os.remove(in_path)
        if os.path.exists(out_path): os.remove(out_path)

# ============================================================
# 6. GIAO DIỆN CHÍNH (ENTERPRISE UI)
# ============================================================
st.markdown('<p class="main-title">⚡ Hệ thống Dịch thuật AI Chuyên ngành Ô tô - THACO (Pro)</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Hỗ trợ Excel, Word, PowerPoint, PDF • Chế độ Song Ngữ • Hiệu đính Trực tiếp • Bộ nhớ Dịch thuật TM</p>', unsafe_allow_html=True)

col_left, col_right = st.columns([1.15, 2])

with col_left:
    st.markdown("#### ⚙️ 1. Kết nối & Chế độ Dịch")
    with st.container(border=True):
        proxy_url = st.text_input("🔗 Đường dẫn Cloudflare Worker:", value=DEFAULT_PROXY_URL)
        direct_key = st.text_input("🔑 Groq API Key trực tiếp (Tùy chọn):", type="password")
        
        if st.button("🔄 Kiểm tra & Tải danh sách Model AI", use_container_width=True):
            try:
                with st.spinner("Đang quét máy chủ AI..."):
                    models = fetch_live_chat_models(proxy_url, direct_key)
                    st.session_state.available_models = models
                    st.success(f"✅ Đã tải {len(models)} model AI đang hoạt động!")
            except Exception as e:
                st.error(f"❌ Lỗi kết nối: {str(e)}")

        default_models = st.session_state.get("available_models", [
            "qwen/qwen3.8-27b",
            "meta-llama/llama-4-scout-17b-16e-instruct",
            "meta-llama/llama-4-maverick-17b-128e-instruct",
            "deepseek-r1-distill-llama-70b"
        ])
        selected_model = st.selectbox("🤖 Chọn Model AI:", options=default_models)
        direction = st.radio("🔄 Chiều dịch thuật:", ["Việt → Anh (VI → EN)", "Anh → Việt (EN → VI)"], horizontal=True)

    st.markdown("#### 💎 2. Tùy chọn Nâng cao (Pro Features)")
    with st.container(border=True):
        bilingual_mode = st.checkbox("🌐 Xuất Song Ngữ (Giữ câu gốc + Mở ngoặc bản dịch ở dòng dưới)", value=False)
        translate_sheets = st.checkbox("📑 Tự động dịch tên Sheet Excel (Giới hạn chuẩn 31 ký tự)", value=True)
        auto_fit_excel = st.checkbox("📐 Tự động bật Wrap Text & Giãn dòng Excel (Chống khuất chữ)", value=True)
        enable_review = st.checkbox("✏️ Bật Bảng Hiệu đính Nhanh trước khi tải file", value=True)
        use_tm = st.checkbox(f"🧠 Sử dụng Bộ nhớ Dịch thuật TM (Đang lưu {len(st.session_state.translation_memory)} câu)", value=True)
        
        col_tm1, col_tm2 = st.columns(2)
        with col_tm1:
            if st.session_state.translation_memory:
                tm_json_bytes = json.dumps(st.session_state.translation_memory, ensure_ascii=False, indent=2).encode("utf-8")
                st.download_button("💾 Tải từ điển TM", data=tm_json_bytes, file_name="THACO_TM.json", mime="application/json", use_container_width=True)
        with col_tm2:
            if st.button("🗑️ Xóa bộ nhớ TM", use_container_width=True):
                st.session_state.translation_memory = {}
                if os.path.exists(TM_FILE_PATH): os.remove(TM_FILE_PATH)
                st.rerun()

    st.markdown("#### 📂 3. Tải lên Tài liệu")
    allowed_exts = ["xlsx", "docx"]
    if HAS_PPTX: allowed_exts.append("pptx")
    if HAS_PDF: allowed_exts.append("pdf")
    uploaded_file = st.file_uploader("Hỗ trợ Excel (.xlsx), Word (.docx), PowerPoint (.pptx), PDF (.pdf)", type=allowed_exts)

with col_right:
    st.markdown("#### 🖥️ Màn hình Giám sát (Live Console)")
    log_placeholder = st.empty()
    display_logs = "<br>".join(st.session_state.logs[-45:]) if st.session_state.logs else (
        "✅ Hệ thống THACO AI Translator Pro sẵn sàng!<br>"
        "• Đã tích hợp 6 nâng cấp: PowerPoint (.pptx), Song ngữ, Đa chiều (Việt↔Anh), Hiệu đính, Translation Memory & Auto-Fit.<br>"
        "• Tải file lên và bấm <b>Bắt đầu Dịch bằng AI</b>."
    )
    log_placeholder.markdown(f"<div class='log-container'>{display_logs}</div>", unsafe_allow_html=True)
    
    progress_bar = st.empty()

    if uploaded_file:
        if st.button("🚀 BẮT ĐẦU QUÉT & DỊCH BẰNG AI", type="primary", use_container_width=True):
            st.session_state.logs = []
            st.session_state.translation_done = False
            st.session_state.file_bytes = uploaded_file.getvalue()
            st.session_state.file_name = uploaded_file.name
            file_ext = os.path.splitext(uploaded_file.name)[1].lower()
            
            add_log(f"🚀 Bắt đầu xử lý: {uploaded_file.name} | Chiều: {direction}", log_placeholder)
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp_in:
                tmp_in.write(st.session_state.file_bytes)
                tmp_path = tmp_in.name

            try:
                pb = progress_bar.progress(0.0)
                start_t = time.time()
                
                add_log("🔍 Đang quét cấu trúc tài liệu và lọc bỏ mã phụ tùng/số liệu...", log_placeholder)
                unique_strs = extract_unique_strings_from_file(tmp_path, file_ext, direction, translate_sheets)
                add_log(f"📊 Tìm thấy {len(unique_strs)} câu/thuật ngữ duy nhất cần dịch.", log_placeholder)
                
                trans_map = build_translation_map(
                    unique_strs, selected_model, direction, proxy_url, direct_key, use_tm, log_placeholder, pb
                )
                
                df_rows = [{"Văn bản gốc": k, "Bản dịch AI (Có thể sửa trực tiếp)": v} for k, v in trans_map.items()]
                st.session_state.review_df = pd.DataFrame(df_rows)
                st.session_state.translation_done = True
                st.session_state.last_direction = direction
                
                elapsed = round(time.time() - start_t, 1)
                add_log(f"🎉 Đã dịch xong toàn bộ nội dung trong {elapsed} giây! Kiểm tra bên dưới để tải file.", log_placeholder)
            except Exception as e:
                add_log(f"❌ LỖI HỆ THỐNG: {str(e)}", log_placeholder)
            finally:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)

    if st.session_state.translation_done and st.session_state.file_bytes is not None:
        st.markdown("---")
        final_trans_map = {}
        
        if enable_review and st.session_state.review_df is not None and not st.session_state.review_df.empty:
            st.markdown("#### ✏️ Bảng Hiệu đính Nhanh (Interactive Review)")
            st.caption("Anh có thể nhấp đúp vào cột **Bản dịch AI** bên dưới để sửa lại bất kỳ thuật ngữ nào. Mọi chỉnh sửa sẽ tự động được lưu vào Bộ nhớ TM và áp dụng ngay vào file tải xuống.")
            edited_df = st.data_editor(
                st.session_state.review_df,
                use_container_width=True,
                height=260,
                num_rows="fixed",
                disabled=["Văn bản gốc"]
            )
            for _, row in edited_df.iterrows():
                orig_k = str(row["Văn bản gốc"])
                edited_v = str(row["Bản dịch AI (Có thể sửa trực tiếp)"]).strip()
                final_trans_map[orig_k] = edited_v
                dir_key = st.session_state.get("last_direction", direction)
                st.session_state.translation_memory[f"{dir_key}||{orig_k}"] = edited_v
            save_tm_to_disk()
        elif st.session_state.review_df is not None:
            for _, row in st.session_state.review_df.iterrows():
                final_trans_map[str(row["Văn bản gốc"])] = str(row["Bản dịch AI (Có thể sửa trực tiếp)"])

        orig_name = st.session_state.file_name
        f_ext = os.path.splitext(orig_name)[1].lower()
        out_ext = ".docx" if f_ext == ".pdf" else f_ext
        suffix_mode = "_Bilingual" if bilingual_mode else ("_EN" if "Việt → Anh" in direction else "_VI")
        out_filename = orig_name.replace(f_ext, f"{suffix_mode}{out_ext}")

        with st.spinner("Đang đóng gói định dạng file hoàn thiện..."):
            output_bytes = generate_final_output_bytes(
                st.session_state.file_bytes,
                f_ext,
                final_trans_map,
                bilingual=bilingual_mode,
                translate_sheets=translate_sheets,
                auto_fit=auto_fit_excel
            )

        st.download_button(
            label=f"⬇️ TẢI XUỐNG TÀI LIỆU HOÀN THIỆN ({out_filename})",
            data=output_bytes,
            file_name=out_filename,
            mime="application/octet-stream",
            type="primary",
            use_container_width=True
        )
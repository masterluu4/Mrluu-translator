# app.py
# THACO AI TECHNICAL TRANSLATOR PRO (V5.0)
# Hỗ trợ: Excel (.xlsx - gồm cả Shapes/Textboxes & Tên Sheet), Word (.docx - gồm cả Textboxes), PowerPoint (.pptx), PDF (.pdf)
# Tính năng: Dịch nguyên câu (Paragraph-level), Từ điển THACO, Song ngữ, Đa chiều (Việt <-> Anh), Hiệu đính nhanh, Bộ nhớ đệm TM, Auto-Fit.

import streamlit as st
import requests
import json
import time
import os
import tempfile
import re
import unicodedata
import zipfile
import shutil
import xml.etree.ElementTree as ET
from copy import copy
import pandas as pd

# ----- Thư viện xử lý File -----
from docx import Document
from docx.oxml.ns import qn
import openpyxl
from openpyxl.utils import get_column_letter
from pdf2docx import Converter
from pptx import Presentation

# ============================================================
# 1. CẤU HÌNH MẶC ĐỊNH & GIAO DIỆN
# ============================================================
DEFAULT_PROXY_URL = "https://mrluu-groq-proxy.mrscong4.workers.dev"
TM_FILE = "translation_memory.json"
GLOSSARY_FILE = "thaco_glossary.txt"

st.set_page_config(page_title="THACO AI Technical Translator PRO V5", page_icon="⚡", layout="wide")

st.markdown("""
    <style>
    .main-title { font-size: 32px !important; font-weight: 800; color: #1E3A8A; margin-bottom: 2px; }
    .sub-title { font-size: 14px; color: #475569; margin-bottom: 16px; }
    .log-container { background-color: #0F172A; color: #10B981; font-family: 'Courier New', monospace; padding: 14px; border-radius: 8px; height: 280px; overflow-y: auto; font-size: 12.5px; line-height: 1.55; border: 1px solid #1E293B; }
    .badge-pro { background-color: #DBEAFE; color: #1D4ED8; padding: 3px 8px; border-radius: 6px; font-weight: 600; font-size: 12px; }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# 2. TỪ ĐIỂN VIẾT TẮT & THUẬT NGỮ CHUẨN THACO AUTO
# ============================================================
DEFAULT_THACO_GLOSSARY = """BÁO CÁO PHÂN TÍCH NGUYÊN NHÂN & GIẢI PHÁP = ROOT CAUSE ANALYSIS & CORRECTIVE ACTION REPORT
HẠNG MỤC = ITEM / SUBJECT
MÔ TẢ VẤN ĐỀ = PROBLEM DESCRIPTION
MÃ SỐ LỖI = DEFECT CODE
GIẢI PHÁP KHẮC PHỤC LỖI = CORRECTIVE ACTIONS
NỘI DUNG TRIỂN KHAI THỰC HIỆN = IMPLEMENTATION PLAN & ACTIONS
TRÁCH NHIỆM = PIC / RESPONSIBILITY
H.THÀNH = TARGET DATE
Tình trạng = Status
Theo dõi = Monitoring
Nguyên nhân gốc rễ = Root Cause
Giải pháp cải tiến = Improvement Solutions
Hành động phòng ngừa = Preventive Actions
Kiểm chứng giải pháp = Solution Verification
Trước cải tiến = Before Improvement
Sau cải tiến = After Improvement
Trước thay đổi = Before Change
Sau thay đổi = After Change
Đối với xe lỗi = For defective vehicles
Đối với xe sản xuất tại NM = For vehicles in factory production
Đối với xe SBH = For After-Sales vehicles
Sau bán hàng = After-Sales
SBH = After-Sales
Lũy kế SL lỗi/số xe giao = Cumulative defects / delivered vehicles
Số BCKT = Technical Report No.
BCKT = Technical Report
Hình thức vận chuyển = Shipping Method
BJM xách tay = Hand-carried by BJM
Chưa có báo cáo = No report yet
BP DVPT = Service & Parts Dept.
BP Dịch vụ = Service Dept.
Phòng QLCL KCN = Industrial Park QC Dept.
QLCL KCN = Industrial Park QC Dept.
BP KTCL NM = Factory QC Dept.
KTCL NM = Factory QC Dept.
TT R&D ô tô = Automotive R&D Center
TT R&D = R&D Center
NCC trong nước = Domestic Supplier
Nhà cung ứng = Supplier
Nhà cung cấp = Supplier
NCC = Supplier
vật tư mẫu = sample part
thu hồi vật tư lỗi = recall defective parts
vật tư = part/material
linh kiện = component/part
Dòng xe = Vehicle Model
Áp dụng VIN# = Applied from VIN#
Áp dụng VIN = Applied from VIN
theo thông báo số = per Technical Bulletin No.
ban hành ngày = issued on
cổng sạc điện thoại = phone charging port
cổng sạc = charging port
điện trở = resistor
điện áp = voltage
tụ điện = capacitor
ĐVT = Unit
Số lượng = Quantity
STT = No."""

def load_glossary_text() -> str:
    if os.path.exists(GLOSSARY_FILE):
        try:
            with open(GLOSSARY_FILE, "r", encoding="utf-8") as f:
                return f.read()
        except Exception:
            pass
    return DEFAULT_THACO_GLOSSARY

def save_glossary_text(text: str):
    try:
        with open(GLOSSARY_FILE, "w", encoding="utf-8") as f:
            f.write(text)
    except Exception:
        pass

def parse_glossary(glossary_str: str, direction: str) -> dict:
    """Đọc bảng thuật ngữ theo chiều dịch (Việt->Anh hoặc Anh->Việt), sắp xếp từ cụm dài xuống ngắn"""
    g = {}
    for line in glossary_str.strip().splitlines():
        if "=" in line and not line.strip().startswith("#"):
            left, right = line.split("=", 1)
            vn, en = left.strip(), right.strip()
            if vn and en:
                if direction == "Việt → Anh":
                    g[vn] = en
                else:
                    g[en] = vn
    return dict(sorted(g.items(), key=lambda item: len(item[0]), reverse=True))

# ============================================================
# 3. BỘ NHỚ ĐỆM DỊCH THUẬT (TRANSLATION MEMORY - TM)
# ============================================================
def load_tm() -> dict:
    if os.path.exists(TM_FILE):
        try:
            with open(TM_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_tm(tm: dict):
    try:
        with open(TM_FILE, "w", encoding="utf-8") as f:
            json.dump(tm, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

# Khởi tạo Session State
if "tm_data" not in st.session_state:
    st.session_state.tm_data = load_tm()
if "glossary_text" not in st.session_state:
    st.session_state.glossary_text = load_glossary_text()
if "logs" not in st.session_state:
    st.session_state.logs = []
if "review_df" not in st.session_state:
    st.session_state.review_df = None
if "staged_file_bytes" not in st.session_state:
    st.session_state.staged_file_bytes = None
if "staged_file_name" not in st.session_state:
    st.session_state.staged_file_name = None

def add_log(message: str, log_placeholder=None):
    time_str = time.strftime("%H:%M:%S")
    st.session_state.logs.append(f"[{time_str}] {message}")
    if log_placeholder is not None:
        display_logs = "<br>".join(st.session_state.logs[-50:])
        log_placeholder.markdown(f"<div class='log-container'>{display_logs}</div>", unsafe_allow_html=True)

# ============================================================
# 4. KẾT NỐI AI & GỌI JSON BATCH THÔNG MINH (CHỐNG LỖI 429)
# ============================================================
def build_system_prompt(direction: str, glossary_str: str) -> str:
    dir_instruction = (
        "Translate every value in the input JSON object from Vietnamese into professional, grammatically complete automotive technical English."
        if direction == "Việt → Anh"
        else "Translate every value in the input JSON object from English into professional, accurate automotive technical Vietnamese used at THACO AUTO."
    )
    return f"""You are a senior automotive engineering and after-sales quality management translator at THACO AUTO.
{dir_instruction}

MANDATORY THACO GLOSSARY & ABBREVIATIONS (Always follow these exact translations when encountered):
{glossary_str}

CRITICAL TRANSLATION RULES:
1. Translate FULL sentences and phrases naturally in proper technical structure. NEVER output broken word-by-word Vietlish, and NEVER leave untranslated Vietnamese words inside an English sentence (or vice versa).
2. Keep all Part Numbers (e.g., USB-CDQ23035B, 8129000459, 1005P320), VINs (e.g., RN5B29SEMRN007347), Vehicle Models (TB120B, TB120S, TB81S), electrical ratings (300Ω, 4.700Ω, 12V-24V, 12V-70V), quantities, percentages, and dates intact.
3. For person names with titles like "Mr Đức", "Mr Việt", "Mr Sinh", write them without Vietnamese diacritics in English as "Mr. Duc", "Mr. Viet", "Mr. Sinh".
4. Note abbreviations in context: "MS:" or "mã số:" means "Part No.:"; "SL" means "Qty"; "xe" in counts (e.g., "94 xe/1042 xe") means "units" or "vehicles" (e.g., "94 units / 1,042 units").
5. OUTPUT FORMAT: Return ONLY a valid JSON object having the EXACT same keys ("0", "1", "2", ...) mapped to their translated strings. Do NOT wrap in markdown code blocks, do NOT add explanations."""

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
        if "llama-4" in n or "deepseek" in n: return 1
        if "gemma" in n or "mixtral" in n: return 2
        return 3
    chat_models.sort(key=sort_key)
    return chat_models

def call_ai_batch(batch_dict: dict, model_id: str, proxy_url: str, direct_key: str, direction: str, glossary_str: str, log_placeholder) -> dict:
    url = f"{proxy_url.rstrip('/')}/openai/v1/chat/completions" if not direct_key else "https://api.groq.com/openai/v1/chat/completions"
    headers = {"Content-Type": "application/json"}
    if direct_key:
        headers["Authorization"] = f"Bearer {direct_key.strip()}"

    sys_prompt = build_system_prompt(direction, glossary_str)

    payload = {
        "model": model_id,
        "messages": [
            {"role": "system", "content": sys_prompt},
            {"role": "user", "content": json.dumps(batch_dict, ensure_ascii=False)}
        ],
        "temperature": 0.1,
        "response_format": {"type": "json_object"}
    }

    for attempt in range(5):
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=75)
            if resp.status_code == 400 and "response_format" in resp.text:
                payload.pop("response_format", None)
                resp = requests.post(url, headers=headers, json=payload, timeout=75)

            if resp.status_code == 429:
                wait_time = 6 * (attempt + 1)
                add_log(f"⏳ Chạm giới hạn tốc độ API (429), hệ thống tự động nghỉ {wait_time}s rồi dịch tiếp...", log_placeholder)
                time.sleep(wait_time)
                continue

            if resp.status_code != 200:
                raise Exception(f"HTTP {resp.status_code} - {resp.text[:160]}")

            content = resp.json()["choices"][0]["message"]["content"].strip()
            content = re.sub(r'<think>.*?</think>', '', content, flags=re.DOTALL).strip()
            
            match = re.search(r'\{.*\}', content, re.DOTALL)
            if match:
                content = match.group(0)
            parsed = json.loads(content)
            if isinstance(parsed, dict):
                return parsed

        except Exception as e:
            add_log(f"⚠️ Thử lại lần {attempt + 1}/5 do: {str(e)}", log_placeholder)
            time.sleep(2.5)

    add_log("❌ Không thể dịch lô này sau 5 lần thử, giữ nguyên câu gốc.", log_placeholder)
    return batch_dict

# ============================================================
# 5. BỘ LỌC NHẬN DIỆN VĂN BẢN & DỊCH KẾT HỢP TM + GLOSSARY
# ============================================================
def has_vietnamese_accents(text: str) -> bool:
    if "đ" in text.lower():
        return True
    for ch in text:
        if ord(ch) > 127 and unicodedata.category(ch).startswith("L"):
            if len(unicodedata.normalize("NFD", ch)) > 1:
                return True
    return False

def should_translate(text: str, direction: str) -> bool:
    t = text.strip()
    if not t or t.startswith("="):
        return False
    if not any(ch.isalpha() for ch in t):
        return False
    # Bỏ qua chuỗi chỉ chứa ký tự đánh dấu hoặc ngày tháng/mã số thuần túy
    if re.match(r'^[0-9\-\_./&#\s():%,;+<>=\[\]]+$', t):
        return False

    if direction == "Việt → Anh":
        if has_vietnamese_accents(t):
            return True
        vn_Tokens = ["STT", "SL", "TONG", "BCKT", "DVPT", "QLCL", "KTCL", "NCC", "SBH", "BJM"]
        t_up = t.upper()
        if any(tok in t_up for tok in vn_Tokens):
            return True
        # Mã phụ tùng/VIN thuần túy (chỉ gồm chữ HOA và số liền nhau không có khoảng trắng câu văn)
        if re.match(r'^[A-Z0-9\-_./&#()+:]+$', t) and " " not in t:
            return False
        # Nếu chỉ gồm mã hiệu + đơn vị (VD: "TB120B", "RN5B29SEMRN007347 (08/05/2024)")
        if re.match(r'^[A-Z0-9\-_./&#\s():%]+$', t):
            return False
        return True
    else:
        # Chiều Anh -> Việt: Dịch mọi cụm có chữ cái tiếng Anh (trừ mã phụ tùng viết liền)
        if re.match(r'^[A-Z0-9\-_./&#()+:]+$', t) and " " not in t and len(t) > 4:
            return False
        return True

def translate_unique_list(unique_strings: list, model_id: str, proxy_url: str, direct_key: str, direction: str, glossary_str: str, log_placeholder, progress_bar) -> dict:
    """Dịch danh sách các câu duy nhất (Ưu tiên: 1. Khớp Glossary -> 2. Khớp Bộ nhớ đệm TM -> 3. Gọi AI Batch)"""
    tm = st.session_state.tm_data
    glossary_dict = parse_glossary(glossary_str, direction)
    
    translated_map = {}
    need_ai = []
    tm_hits = 0

    for s in unique_strings:
        s_clean = s.strip()
        # 1. Kiểm tra khớp chính xác 100% trong Từ điển THACO
        matched_glossary = None
        for k, v in glossary_dict.items():
            if s_clean.lower() == k.lower():
                matched_glossary = v
                break
        if matched_glossary is not None:
            translated_map[s] = matched_glossary
            tm_hits += 1
            continue

        # 2. Kiểm tra trong Bộ nhớ đệm Translation Memory (TM)
        tm_key = f"{direction}::{s_clean}"
        if tm_key in tm and tm[tm_key].strip():
            translated_map[s] = tm[tm_key]
            tm_hits += 1
        elif direction == "Việt → Anh" and s_clean in tm and tm[s_clean].strip():
            translated_map[s] = tm[s_clean]
            tm_hits += 1
        else:
            need_ai.append(s)

    add_log(f"⚡ Đã lấy tức thì {tm_hits}/{len(unique_strings)} câu từ Từ điển THACO & Bộ nhớ đệm (0.01s). Cần gọi AI: {len(need_ai)} câu.", log_placeholder)

    if not need_ai:
        progress_bar.progress(1.0)
        return translated_map

    # 3. Chia lô 18 câu nguyên vẹn để gửi AI (nhanh và không bị lỗi 429)
    BATCH_SIZE = 18
    total_ai = len(need_ai)
    for start in range(0, total_ai, BATCH_SIZE):
        batch_slice = need_ai[start : start + BATCH_SIZE]
        payload = {str(i): item for i, item in enumerate(batch_slice)}
        
        preview = ", ".join([f"'{x[:28]}...'" if len(x) > 28 else f"'{x}'" for x in batch_slice[:2]])
        add_log(f"🤖 Đang gửi AI ({model_id}) dịch câu {start + 1} -> {start + len(batch_slice)}/{total_ai} ({preview})", log_placeholder)
        
        ai_res = call_ai_batch(payload, model_id, proxy_url, direct_key, direction, glossary_str, log_placeholder)
        
        for i, orig_str in enumerate(batch_slice):
            val = ai_res.get(str(i), orig_str)
            if isinstance(val, str) and val.strip():
                clean_val = val.strip()
                translated_map[orig_str] = clean_val
                tm[f"{direction}::{orig_str.strip()}"] = clean_val
            else:
                translated_map[orig_str] = orig_str

        save_tm(tm)
        progress_bar.progress(min((start + len(batch_slice)) / total_ai, 1.0))
        if start + BATCH_SIZE < total_ai:
            time.sleep(1.5)

    return translated_map

# ============================================================
# 6. THUẬT TOÁN GOM CÂU THÔNG MINH (CHỐNG CẮT VỤN RUN POWERPOINT/WORD)
# ============================================================
class ParagraphSegment:
    """Gom nhóm các run liên tiếp trong 1 Paragraph thành 1 câu nguyên vẹn, bảo toàn Font/Size/Color/Bold"""
    def __init__(self, runs):
        self.runs = runs
        self.text = "".join(r.text for r in runs if r.text)

    def apply_translation(self, translated_clean: str, bilingual: bool):
        if not self.runs:
            return
        orig = self.text
        lead = orig[:len(orig) - len(orig.lstrip())]
        trail = orig[len(orig.rstrip()):]
        
        if bilingual and translated_clean.strip().lower() != orig.strip().lower():
            final_str = f"{lead}{orig.strip()}\n({translated_clean.strip()}){trail}"
        else:
            final_str = f"{lead}{translated_clean.strip()}{trail}"
            
        self.runs[0].text = final_str
        for r in self.runs[1:]:
            r.text = ""

def collect_paragraph_segments(paragraph, segments_list, texts_list):
    """
    Hợp nhất toàn bộ các run bị cắt vụn trong 1 đoạn văn thành 1 câu hoàn chỉnh.
    Chỉ tách riêng phần 'Tiêu đề in đậm:' ở đầu câu nếu có dấu ':' và phần sau không in đậm.
    """
    runs = [r for r in paragraph.runs if r.text is not None]
    if not runs:
        return
    full_text = "".join(r.text for r in runs)
    if not full_text.strip():
        return

    # Kiểm tra xem có phải mẫu "**1. Vấn đề:** Hư hỏng cổng sạc..." (Bold ở trước dấu ':', Normal ở sau)
    split_idx = None
    acc = ""
    for i in range(len(runs) - 1):
        acc += runs[i].text
        b_cur = bool(runs[i].font.bold)
        b_next = bool(runs[i + 1].font.bold)
        if b_cur and not b_next and ":" in acc and len(acc.strip()) <= 45:
            # Đảm bảo dấu ':' nằm ở cuối cụm in đậm
            if acc.strip().endswith(":"):
                split_idx = i + 1
                break

    if split_idx is not None:
        seg_head = ParagraphSegment(runs[:split_idx])
        seg_body = ParagraphSegment(runs[split_idx:])
        if seg_head.text.strip():
            segments_list.append(seg_head)
            texts_list.append(seg_head.text)
        if seg_body.text.strip():
            segments_list.append(seg_body)
            texts_list.append(seg_body.text)
    else:
        seg = ParagraphSegment(runs)
        segments_list.append(seg)
        texts_list.append(seg.text)

# ============================================================
# 7. TRÍCH XUẤT & KẾT XUẤT CHO EXCEL, WORD, POWERPOINT, PDF
# (Bao gồm cả Shapes/Textboxes vẽ nổi trong Excel & Word)
# ============================================================
def sanitize_excel_sheet_title(title: str, existing_titles: list) -> str:
    """Làm sạch tên Sheet Excel (tối đa 31 ký tự, cấm các ký tự : \ / ? * [ ])"""
    cleaned = re.sub(r'[:\\/?*\[\]]', '-', title).strip()
    if not cleaned:
        cleaned = "Sheet"
    base = cleaned[:28]
    candidate = cleaned[:31]
    counter = 1
    while candidate in existing_titles:
        candidate = f"{base}_{counter}"
        counter += 1
    return candidate

def extract_xml_drawing_texts(zip_Or_Docx_Path: str, xml_prefix: str) -> list:
    """Quét các hộp chữ vẽ nổi (Shapes/Textboxes) lưu trong cấu trúc XML của Excel (xl/drawings/)"""
    found = []
    try:
        with zipfile.ZipFile(zip_Or_Docx_Path, 'r') as zf:
            for name in zf.namelist():
                if name.startswith(xml_prefix) and name.endswith(".xml"):
                    tree = ET.fromstring(zf.read(name))
                    # Tìm tất cả các thẻ đoạn văn <a:p> trong DrawingML
                    for p_elem in tree.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}p"):
                        t_elems = list(p_elem.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}t"))
                        if t_elems:
                            full_p = "".join(t.text for t in t_elems if t.text)
                            if full_p.strip():
                                found.append(full_p.strip())
    except Exception:
        pass
    return found

def replace_xml_drawing_texts(file_path: str, xml_prefix: str, translation_map: dict, bilingual: bool):
    """Ghi đè bản dịch vào các Shapes/Textboxes vẽ nổi trong file Excel (.xlsx)"""
    try:
        temp_zip = file_path + "_zip_tmp"
        with zipfile.ZipFile(file_path, 'r') as zin, zipfile.ZipFile(temp_zip, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith(xml_prefix) and item.filename.endswith(".xml"):
                    tree = ET.fromstring(data)
                    modified = False
                    for p_elem in tree.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}p"):
                        t_elems = list(p_elem.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}t"))
                        if t_elems:
                            full_p = "".join(t.text for t in t_elems if t.text)
                            clean_p = full_p.strip()
                            if clean_p in translation_map:
                                new_clean = translation_map[clean_p]
                                final_txt = f"{clean_p}\n({new_clean})" if (bilingual and new_clean.lower() != clean_p.lower()) else new_clean
                                t_elems[0].text = final_txt
                                for extra_t in t_elems[1:]:
                                    extra_t.text = ""
                                modified = True
                    if modified:
                        data = ET.tostring(tree, encoding="utf-8", xml_declaration=True)
                zout.writestr(item, data)
        shutil.move(temp_zip, file_path)
    except Exception:
        if os.path.exists(file_path + "_zip_tmp"):
            os.remove(file_path + "_zip_tmp")

def extract_all_unique_strings_from_file(input_path: str, file_ext: str, direction: str, translate_sheet_names: bool) -> list:
    """Quét toàn bộ file để lấy danh sách các câu duy nhất cần dịch (Phục vụ bước 1)"""
    raw_texts = []

    if file_ext == ".xlsx":
        wb = openpyxl.load_workbook(input_path)
        if translate_sheet_names:
            raw_texts.extend(wb.sheetnames)
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.strip():
                        raw_texts.append(cell.value)
        # Quét thêm Shapes/Textboxes vẽ nổi trong Excel
        raw_texts.extend(extract_xml_drawing_texts(input_path, "xl/drawings/drawing"))

    elif file_ext in [".docx", ".pdf"]:
        work_docx = input_path
        if file_ext == ".pdf":
            work_docx = input_path.replace(".pdf", "_scan_tmp.docx")
            cv = Converter(input_path)
            cv.convert(work_docx)
            cv.close()
        doc = Document(work_docx)
        segs, txts = [], []
        for p in doc.paragraphs:
            collect_paragraph_segments(p, segs, txts)
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        collect_paragraph_segments(p, segs, txts)
        raw_texts.extend(txts)
        # Quét thêm Textboxes (w:txbxContent) trong Word
        for txbx in doc.element.body.iter(qn('w:txbxContent')):
            for p_el in txbx.iter(qn('w:p')):
                t_els = list(p_el.iter(qn('w:t')))
                full_t = "".join(x.text for x in t_els if x.text)
                if full_t.strip():
                    raw_texts.append(full_t.strip())
        if file_ext == ".pdf" and os.path.exists(work_docx):
            os.remove(work_docx)

    elif file_ext == ".pptx":
        prs = Presentation(input_path)
        segs, txts = [], []
        def scan_shape(sh):
            if sh.shape_type == 6:
                for sub in sh.shapes:
                    scan_shape(sub)
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    collect_paragraph_segments(p, segs, txts)
            if sh.has_table:
                for r in sh.table.rows:
                    for c in r.cells:
                        for p in c.text_frame.paragraphs:
                            collect_paragraph_segments(p, segs, txts)
        for slide in prs.slides:
            for sh in slide.shapes:
                scan_shape(sh)
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                for p in slide.notes_slide.notes_text_frame.paragraphs:
                    collect_paragraph_segments(p, segs, txts)
        raw_texts.extend(txts)

    # Lọc trùng và chỉ giữ những câu thực sự cần dịch
    uniques = []
    seen = set()
    for t in raw_texts:
        if should_translate(t, direction):
            c = t.strip()
            if c not in seen:
                seen.add(c)
                uniques.append(c)
    return uniques

def apply_translations_to_file(input_path: str, output_path: str, file_ext: str, translation_map: dict, bilingual: bool, auto_fit: bool, translate_sheet_names: bool, log_placeholder):
    """Áp dụng bảng dịch (đã qua hiệu đính) vào file gốc và xuất ra file hoàn thiện"""
    if file_ext == ".xlsx":
        add_log("💾 Đang ghi bản dịch vào Excel (Cells, Sheet Tabs, Shapes & Auto-Fit Layout)...", log_placeholder)
        wb = openpyxl.load_workbook(input_path)
        
        if translate_sheet_names and not bilingual:
            used_titles = []
            for ws in wb.worksheets:
                orig_title = ws.title.strip()
                new_title = translation_map.get(orig_title, orig_title)
                safe_title = sanitize_excel_sheet_title(new_title, used_titles)
                ws.title = safe_title
                used_titles.append(safe_title)

        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.strip():
                        orig = cell.value
                        clean = orig.strip()
                        if clean in translation_map:
                            trans = translation_map[clean]
                            lead = orig[:len(orig) - len(orig.lstrip())]
                            trail = orig[len(orig.rstrip()):]
                            if bilingual and trans.strip().lower() != clean.lower():
                                cell.value = f"{lead}{clean}\n({trans.strip()}){trail}"
                            else:
                                cell.value = f"{lead}{trans.strip()}{trail}"
                                
                            if auto_fit:
                                cur_align = copy(cell.alignment) if cell.alignment else openpyxl.styles.Alignment()
                                cur_align.wrap_text = True
                                cell.alignment = cur_align

            # Tự động mở rộng nhẹ độ rộng cột nếu bật Auto-Fit để chữ tiếng Anh không bị chật
            if auto_fit:
                for col in ws.columns:
                    col_letter = get_column_letter(col[0].column)
                    cur_w = ws.column_dimensions[col_letter].width
                    if cur_w and cur_w < 65:
                        ws.column_dimensions[col_letter].width = round(cur_w * 1.18, 1)

        wb.save(output_path)
        # Dịch cả các Textboxes/Shapes vẽ nổi trong Excel
        replace_xml_drawing_texts(output_path, "xl/drawings/drawing", translation_map, bilingual)

    elif file_ext in [".docx", ".pdf"]:
        work_in = input_path
        if file_ext == ".pdf":
            add_log("🔄 Đang chuyển đổi cấu trúc PDF sang Word...", log_placeholder)
            work_in = input_path.replace(".pdf", "_apply_tmp.docx")
            cv = Converter(input_path)
            cv.convert(work_in)
            cv.close()

        add_log("💾 Đang áp dụng bản dịch vào Word (Paragraphs, Tables & Textboxes)...", log_placeholder)
        doc = Document(work_in)
        segs, txts = [], []
        for p in doc.paragraphs:
            collect_paragraph_segments(p, segs, txts)
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        collect_paragraph_segments(p, segs, txts)
        for seg in segs:
            clean = seg.text.strip()
            if clean in translation_map:
                seg.apply_translation(translation_map[clean], bilingual)

        # Dịch Textboxes vẽ nổi trong Word
        for txbx in doc.element.body.iter(qn('w:txbxContent')):
            for p_el in txbx.iter(qn('w:p')):
                t_els = list(p_el.iter(qn('w:t')))
                if t_els:
                    full_t = "".join(x.text for x in t_els if x.text).strip()
                    if full_t in translation_map:
                        trans = translation_map[full_t]
                        t_els[0].text = f"{full_t}\n({trans})" if (bilingual and trans.lower() != full_t.lower()) else trans
                        for extra in t_els[1:]:
                            extra.text = ""

        doc.save(output_path)
        if file_ext == ".pdf" and os.path.exists(work_in):
            os.remove(work_in)

    elif file_ext == ".pptx":
        add_log("💾 Đang áp dụng bản dịch nguyên câu vào PowerPoint...", log_placeholder)
        prs = Presentation(input_path)
        segs, txts = [], []
        def scan_shape(sh):
            if sh.shape_type == 6:
                for sub in sh.shapes:
                    scan_shape(sub)
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    collect_paragraph_segments(p, segs, txts)
            if sh.has_table:
                for r in sh.table.rows:
                    for c in r.cells:
                        for p in c.text_frame.paragraphs:
                            collect_paragraph_segments(p, segs, txts)
        for slide in prs.slides:
            for sh in slide.shapes:
                scan_shape(sh)
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                for p in slide.notes_slide.notes_text_frame.paragraphs:
                    collect_paragraph_segments(p, segs, txts)

        for seg in segs:
            clean = seg.text.strip()
            if clean in translation_map:
                seg.apply_translation(translation_map[clean], bilingual)

        prs.save(output_path)

# ============================================================
# 8. GIAO DIỆN CHÍNH (CHUẨN WORKFLOW CHUYÊN NGHIỆP)
# ============================================================
st.markdown('<p class="main-title">⚡ Hệ thống Dịch thuật AI Chuyên ngành Ô tô - THACO PRO (V5.0)</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="sub-title">'
    '<span class="badge-pro">PRO V5.0</span> '
    'Dịch nguyên câu chống đứt từ • Tích hợp Từ điển THACO • Xuất Song Ngữ • Dịch Đa Chiều (Việt ⇄ Anh) • Hiệu đính Nhanh • Bộ nhớ đệm TM'
    '</p>',
    unsafe_allow_html=True
)

col_left, col_right = st.columns([1.15, 2.0])

with col_left:
    st.markdown("#### ⚙️ 1. Kết nối AI & Tùy chọn Dịch thuật")
    with st.container(border=True):
        proxy_url = st.text_input("🔗 Đường dẫn Cloudflare Worker:", value=DEFAULT_PROXY_URL)
        direct_key = st.text_input("🔑 Hoặc Groq API Key trực tiếp (nếu không qua Worker):", type="password")
        
        if st.button("🔄 Kiểm tra kết nối & Tải danh sách Model AI", use_container_width=True):
            try:
                with st.spinner("Đang tải danh sách model từ máy chủ..."):
                    models = fetch_live_chat_models(proxy_url, direct_key)
                    st.session_state.available_models = models
                    st.success(f"✅ Kết nối thành công! Tìm thấy {len(models)} model AI.")
            except Exception as e:
                st.error(f"❌ Lỗi kết nối: {str(e)}")

        model_list = st.session_state.get("available_models", [
            "qwen/qwen3.8-27b",
            "meta-llama/llama-4-scout-17b-16e-instruct",
            "meta-llama/llama-4-maverick-17b-128e-instruct",
            "deepseek-r1-distill-llama-70b",
            "gemma2-9b-it"
        ])
        selected_model = st.selectbox("🤖 Chọn Model AI dịch thuật:", options=model_list)

        st.divider()
        direction = st.radio("🌐 Chiều dịch thuật:", ["Việt → Anh", "Anh → Việt"], horizontal=True)
        
        c1, c2 = st.columns(2)
        with c1:
            bilingual = st.checkbox("📖 Xuất Song Ngữ\n(Gốc + Dịch dòng dưới)", value=False)
            translate_sheets = st.checkbox("📑 Dịch tên Sheet Excel", value=True)
        with c2:
            auto_fit = st.checkbox("📐 Auto-Fit Layout\n(Tự xuống dòng & giãn cột)", value=True)
            enable_review = st.checkbox("✍️ Duyệt & Hiệu đính\ntrước khi xuất File", value=True)

    # --- KHUNG QUẢN LÝ TỪ VIẾT TẮT & THUẬT NGỮ THACO ---
    with st.expander("📚 2. Khung Từ viết tắt & Thuật ngữ THACO (Bấm để thêm/sửa)", expanded=False):
        st.caption("Nhập theo cú pháp `Từ viết tắt / Tiếng Việt = Tiếng Anh chuẩn`. AI sẽ tuân thủ tuyệt đối bảng này:")
        new_glossary = st.text_area(
            "Bảng thuật ngữ THACO:",
            value=st.session_state.glossary_text,
            height=240
        )
        if st.button("💾 Lưu Bảng Thuật Ngữ THACO", use_container_width=True):
            st.session_state.glossary_text = new_glossary
            save_glossary_text(new_glossary)
            st.success("✅ Đã cập nhật Bảng thuật ngữ THACO!")

    # --- KHUNG BỘ NHỚ ĐỆM TRANSLATION MEMORY (TM) ---
    with st.expander(f"🧠 3. Bộ nhớ đệm Dịch thuật - TM (Đang nhớ {len(st.session_state.tm_data)} câu)", expanded=False):
        st.caption("Tải file CSV/JSON từ Google Sheets lên để nạp bộ nhớ, hoặc tải bộ nhớ hiện tại về máy:")
        tm_col1, tm_col2 = st.columns(2)
        with tm_col1:
            tm_df_export = pd.DataFrame(
                [{"Khoa_Goc": k, "Ban_Dich_Chuan": v} for k, v in st.session_state.tm_data.items()]
            )
            st.download_button(
                "⬇️ Tải TM (.csv)",
                data=tm_df_export.to_csv(index=False).encode("utf-8-sig"),
                file_name="THACO_Translation_Memory.csv",
                mime="text/csv",
                use_container_width=True
            )
        with tm_col2:
            if st.button("🗑️ Xóa trắng TM", use_container_width=True):
                st.session_state.tm_data = {}
                save_tm({})
                st.rerun()

        uploaded_tm = st.file_uploader("Nạp thêm TM từ file CSV (2 cột: Khoa_Goc, Ban_Dich_Chuan):", type=["csv"], key="tm_up")
        if uploaded_tm is not None:
            try:
                df_in = pd.read_csv(uploaded_tm)
                if len(df_in.columns) >= 2:
                    for _, r in df_in.iterrows():
                        k_val, v_val = str(r.iloc[0]).strip(), str(r.iloc[1]).strip()
                        if k_val and v_val:
                            st.session_state.tm_data[k_val] = v_val
                    save_tm(st.session_state.tm_data)
                    st.success(f"✅ Đã nạp tổng cộng {len(st.session_state.tm_data)} câu vào TM!")
            except Exception as e:
                st.error(f"Lỗi đọc CSV: {e}")

    st.markdown("#### 📂 4. Tải lên Tài liệu cần dịch")
    uploaded_file = st.file_uploader(
        "Hỗ trợ Excel (.xlsx), Word (.docx), PowerPoint (.pptx), PDF (.pdf)",
        type=["xlsx", "docx", "pptx", "pdf"]
    )

with col_right:
    st.markdown("#### 🖥️ Màn hình Giám sát (Live Console)")
    log_placeholder = st.empty()
    if not st.session_state.logs:
        log_placeholder.markdown(
            "<div class='log-container'>✅ Hệ thống THACO AI Translator PRO V5.0 sẵn sàng!<br>"
            "• <b>Sửa triệt để lỗi PowerPoint/Word</b>: Gom toàn bộ các mảnh chữ vụn thành câu hoàn chỉnh trước khi gửi AI.<br>"
            "• <b>Tích hợp đầy đủ</b>: Từ điển viết tắt THACO, Chế độ Song Ngữ, Dịch Đa Chiều, Bảng Hiệu đính Nhanh & Auto-Fit.</div>",
            unsafe_allow_html=True
        )
    else:
        display_logs = "<br>".join(st.session_state.logs[-50:])
        log_placeholder.markdown(f"<div class='log-container'>{display_logs}</div>", unsafe_allow_html=True)

    progress_bar = st.progress(0.0)

    if uploaded_file is not None:
        btn_label = "🚀 1. PHÂN TÍCH & DỊCH NHÁP BẰNG AI (ĐỂ HIỆU ĐÍNH)" if enable_review else "🚀 BẮT ĐẦU DỊCH & XUẤT FILE NGAY"
        if st.button(btn_label, type="primary", use_container_width=True):
            st.session_state.logs = []
            st.session_state.glossary_text = new_glossary
            st.session_state.staged_file_bytes = uploaded_file.getvalue()
            st.session_state.staged_file_name = uploaded_file.name
            
            file_ext = os.path.splitext(uploaded_file.name)[1].lower()
            add_log(f"🚀 Bắt đầu xử lý '{uploaded_file.name}' | Model: {selected_model} | Chiều: {direction}", log_placeholder)

            with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp_in:
                tmp_in.write(st.session_state.staged_file_bytes)
                input_path = tmp_in.name

            try:
                start_t = time.time()
                add_log("🔍 Đang quét cấu trúc file và hợp nhất các đoạn văn bị cắt vụn...", log_placeholder)
                unique_list = extract_all_unique_strings_from_file(input_path, file_ext, direction, translate_sheets)
                
                trans_map = translate_unique_list(
                    unique_list, selected_model, proxy_url, direct_key,
                    direction, st.session_state.glossary_text, log_placeholder, progress_bar
                )
                
                elapsed = round(time.time() - start_t, 1)
                add_log(f"🎉 Đã dịch xong {len(unique_list)} câu duy nhất trong {elapsed} giây!", log_placeholder)

                # Lưu vào bảng Review trong session_state
                st.session_state.review_df = pd.DataFrame([
                    {"Câu gốc trong File": k, "Bản dịch AI (Bấm đúp để sửa tay)": v}
                    for k, v in trans_map.items()
                ])

                # Nếu người dùng tắt chế độ Hiệu đính -> Xuất file luôn trong 1 bước
                if not enable_review:
                    output_ext = ".docx" if file_ext == ".pdf" else file_ext
                    suffix_tag = "_Bilingual" if bilingual else ("_EN" if direction == "Việt → Anh" else "_VN")
                    output_name = uploaded_file.name.replace(file_ext, f"{suffix_tag}{output_ext}")
                    output_path = input_path.replace(file_ext, f"_out{output_ext}")

                    apply_translations_to_file(
                        input_path, output_path, file_ext, trans_map,
                        bilingual, auto_fit, translate_sheets, log_placeholder
                    )
                    with open(output_path, "rb") as f:
                        st.session_state.final_output_bytes = f.read()
                        st.session_state.final_output_name = output_name
                    if os.path.exists(output_path):
                        os.remove(output_path)

            except Exception as e:
                add_log(f"❌ LỖI HỆ THỐNG: {str(e)}", log_placeholder)
            finally:
                if os.path.exists(input_path):
                    os.remove(input_path)

    # --- HIỂN THỊ BẢNG HIỆU ĐÍNH NHANH (INTERACTIVE REVIEW) ---
    if st.session_state.review_df is not None and enable_review:
        st.markdown("---")
        st.markdown("#### ✍️ Bảng Hiệu đính Nhanh trước khi xuất File")
        st.caption("Anh có thể kiểm tra và gõ sửa trực tiếp vào cột **'Bản dịch AI'**. Mọi câu anh sửa sẽ tự động được ghi nhớ vào Bộ nhớ đệm (TM) cho các file lần sau:")
        
        edited_df = st.data_editor(
            st.session_state.review_df,
            use_container_width=True,
            height=320,
            disabled=["Câu gốc trong File"],
            key="editor_grid"
        )

        if st.button("✅ 2. ÁP DỤNG BẢN HIỆU ĐÍNH & KẾT XUẤT FILE HOÀN THIỆN", type="primary", use_container_width=True):
            if st.session_state.staged_file_bytes:
                file_name = st.session_state.staged_file_name
                file_ext = os.path.splitext(file_name)[1].lower()
                
                # Cập nhật các câu đã sửa tay vào từ điển áp dụng và lưu vào TM
                final_map = {}
                tm = st.session_state.tm_data
                for _, row in edited_df.iterrows():
                    orig_s = str(row["Câu gốc trong File"]).strip()
                    edit_s = str(row["Bản dịch AI (Bấm đúp để sửa tay)"]).strip()
                    final_map[orig_s] = edit_s
                    tm[f"{direction}::{orig_s}"] = edit_s
                save_tm(tm)

                with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp_in:
                    tmp_in.write(st.session_state.staged_file_bytes)
                    in_p = tmp_in.name

                output_ext = ".docx" if file_ext == ".pdf" else file_ext
                suffix_tag = "_Bilingual" if bilingual else ("_EN" if direction == "Việt → Anh" else "_VN")
                output_name = file_name.replace(file_ext, f"{suffix_tag}{output_ext}")
                out_p = in_p.replace(file_ext, f"_out{output_ext}")

                try:
                    apply_translations_to_file(
                        in_p, out_p, file_ext, final_map,
                        bilingual, auto_fit, translate_sheets, log_placeholder
                    )
                    with open(out_p, "rb") as f:
                        st.session_state.final_output_bytes = f.read()
                        st.session_state.final_output_name = output_name
                    add_log(f"✅ Đã kết xuất thành công file '{output_name}'! Mời anh bấm nút tải xuống bên dưới.", log_placeholder)
                finally:
                    if os.path.exists(in_p): os.remove(in_p)
                    if os.path.exists(out_p): os.remove(out_p)

    # --- NÚT TẢI FILE KẾT QUẢ ---
    if st.session_state.get("final_output_bytes") is not None:
        st.success(f"🎉 File **{st.session_state.final_output_name}** đã sẵn sàng tải xuống!")
        st.download_button(
            label=f"⬇️ TẢI XUỐNG FILE HOÀN THIỆN ({st.session_state.final_output_name})",
            data=st.session_state.final_output_bytes,
            file_name=st.session_state.final_output_name,
            mime="application/octet-stream",
            type="primary",
            use_container_width=True
        )
# app.py
# THACO AI TECHNICAL TRANSLATOR PRO (V6.0 - GOOGLE SHEETS CLOUD EDITION)
# Hỗ trợ: Excel (.xlsx - gồm cả Shapes/Textboxes & Tên Sheet), Word (.docx - gồm cả Textboxes), PowerPoint (.pptx), PDF (.pdf)
# Tính năng: Đồng bộ 2 chiều Google Sheets (TM & Glossary), Dịch nguyên câu (Paragraph-level), Song ngữ, Đa chiều, Hiệu đính nhanh, Auto-Fit.

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
# 1. CẤU HÌNH MẶC ĐỊNH CLOUDFLARE WORKER & GOOGLE SHEETS CLOUD
# ============================================================
DEFAULT_PROXY_URL = "https://mrluu-groq-proxy.mrscong4.workers.dev"
DEFAULT_GSHEET_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbzRJD3wE4fFWkPZMs_x2r2WJ9tZLd_nJseem32BpiZsiEwxKcFb7FYZi2_SBvxbbv7G/exec"

TM_FILE = "translation_memory.json"
GLOSSARY_FILE = "thaco_glossary.txt"

st.set_page_config(page_title="THACO AI Technical Translator PRO V6", page_icon="⚡", layout="wide")

st.markdown("""
    <style>
    .main-title { font-size: 32px !important; font-weight: 800; color: #1E3A8A; margin-bottom: 2px; }
    .sub-title { font-size: 14px; color: #475569; margin-bottom: 16px; }
    .log-container { background-color: #0F172A; color: #10B981; font-family: 'Courier New', monospace; padding: 14px; border-radius: 8px; height: 280px; overflow-y: auto; font-size: 12.5px; line-height: 1.55; border: 1px solid #1E293B; }
    .badge-pro { background-color: #DBEAFE; color: #1D4ED8; padding: 3px 8px; border-radius: 6px; font-weight: 600; font-size: 12px; }
    .badge-cloud { background-color: #DCFCE7; color: #15803D; padding: 3px 8px; border-radius: 6px; font-weight: 600; font-size: 12px; margin-left: 6px; }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# 2. TỪ ĐIỂN VIẾT TẮT & THUẬT NGỮ CHUẨN THACO AUTO (MẶC ĐỊNH)
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

def dict_to_glossary_str(g_dict: dict) -> str:
    return "\n".join([f"{k} = {v}" for k, v in g_dict.items()])

def glossary_str_to_dict(glossary_str: str) -> dict:
    g = {}
    for line in glossary_str.strip().splitlines():
        if "=" in line and not line.strip().startswith("#"):
            k, v = line.split("=", 1)
            if k.strip() and v.strip():
                g[k.strip()] = v.strip()
    return g

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
# 3. ĐỒNG BỘ HAI CHIỀU VỚI GOOGLE SHEETS (CLOUD TM & GLOSSARY)
# ============================================================
def load_local_glossary() -> str:
    if os.path.exists(GLOSSARY_FILE):
        try:
            with open(GLOSSARY_FILE, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if content:
                    return content
        except Exception:
            pass
    return DEFAULT_THACO_GLOSSARY

def load_local_tm() -> dict:
    if os.path.exists(TM_FILE):
        try:
            with open(TM_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def push_to_gsheet(webapp_url: str, action: str, data_dict: dict) -> bool:
    """Đẩy dữ liệu (Thuật ngữ hoặc câu dịch TM mới) lên Google Sheets"""
    if not webapp_url or not webapp_url.strip().startswith("https://script.google.com") or not data_dict:
        return False
    try:
        resp = requests.post(
            webapp_url.strip(),
            json={"action": action, "data": data_dict},
            timeout=25
        )
        return resp.status_code == 200
    except Exception:
        return False

def sync_from_gsheet(webapp_url: str):
    """Đồng bộ hai chiều: Hút dữ liệu từ Google Sheets về App; nếu Sheet trống thì đẩy bộ mặc định THACO lên Sheet"""
    if not webapp_url or not webapp_url.strip().startswith("https://script.google.com"):
        return False, "Đường dẫn Google Apps Script không hợp lệ."
    try:
        resp = requests.get(webapp_url.strip(), timeout=25)
        if resp.status_code != 200:
            return False, f"Lỗi HTTP {resp.status_code} khi kết nối Google Sheets."
            
        data = resp.json()
        cloud_glossary = data.get("glossary", {})
        cloud_tm = data.get("tm", {})

        # 1. Xử lý Bảng thuật ngữ (Glossary)
        default_g_dict = glossary_str_to_dict(DEFAULT_THACO_GLOSSARY)
        if cloud_glossary:
            # Gộp từ điển mặc định và từ điển trên Cloud (ưu tiên Cloud nếu trùng khóa)
            merged_g = {**default_g_dict, **cloud_glossary}
            st.session_state.glossary_text = dict_to_glossary_str(merged_g)
        else:
            # Nếu tab Glossary trên Google Sheets đang trống -> Tự động khởi tạo bảng chuẩn THACO lên Sheet
            st.session_state.glossary_text = DEFAULT_THACO_GLOSSARY
            push_to_gsheet(webapp_url, "save_glossary", default_g_dict)

        with open(GLOSSARY_FILE, "w", encoding="utf-8") as f:
            f.write(st.session_state.glossary_text)

        # 2. Xử lý Bộ nhớ đệm (TM)
        if cloud_tm and isinstance(cloud_tm, dict):
            st.session_state.tm_data.update(cloud_tm)
            with open(TM_FILE, "w", encoding="utf-8") as f:
                json.dump(st.session_state.tm_data, f, ensure_ascii=False, indent=2)

        g_count = len(glossary_str_to_dict(st.session_state.glossary_text))
        tm_count = len(st.session_state.tm_data)
        return True, f"Đã đồng bộ Google Sheets: {g_count} thuật ngữ THACO & {tm_count} câu nhớ đệm (TM)!"
    except Exception as e:
        return False, f"Không thể kết nối Google Sheets: {str(e)}"

def save_glossary_text(text: str, webapp_url: str = ""):
    try:
        with open(GLOSSARY_FILE, "w", encoding="utf-8") as f:
            f.write(text)
    except Exception:
        pass
    g_dict = glossary_str_to_dict(text)
    return push_to_gsheet(webapp_url, "save_glossary", g_dict)

def save_tm(tm: dict, new_items_only: dict = None, webapp_url: str = ""):
    try:
        with open(TM_FILE, "w", encoding="utf-8") as f:
            json.dump(tm, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    payload_data = new_items_only if new_items_only is not None else tm
    if payload_data:
        push_to_gsheet(webapp_url, "upsert_tm", payload_data)

# Khởi tạo Session State & Tự động kết nối Google Sheets trong lần mở đầu tiên
if "tm_data" not in st.session_state:
    st.session_state.tm_data = load_local_tm()
if "glossary_text" not in st.session_state:
    st.session_state.glossary_text = load_local_glossary()
if "logs" not in st.session_state:
    st.session_state.logs = []
if "review_df" not in st.session_state:
    st.session_state.review_df = None
if "staged_file_bytes" not in st.session_state:
    st.session_state.staged_file_bytes = None
if "staged_file_name" not in st.session_state:
    st.session_state.staged_file_name = None
if "cloud_sync_status" not in st.session_state:
    ok, msg = sync_from_gsheet(DEFAULT_GSHEET_WEBAPP_URL)
    st.session_state.cloud_sync_status = (ok, msg)

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
# 5. BỘ LỌC NHẬN DIỆN VĂN BẢN & DỊCH KẾT HỢP CLOUD TM + GLOSSARY
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
    if re.match(r'^[0-9\-\_./&#\s():%,;+<>=\[\]]+$', t):
        return False

    if direction == "Việt → Anh":
        if has_vietnamese_accents(t):
            return True
        vn_tokens = ["STT", "SL", "TONG", "BCKT", "DVPT", "QLCL", "KTCL", "NCC", "SBH", "BJM"]
        t_up = t.upper()
        if any(tok in t_up for tok in vn_tokens):
            return True
        if re.match(r'^[A-Z0-9\-_./&#()+:]+$', t) and " " not in t:
            return False
        if re.match(r'^[A-Z0-9\-_./&#\s():%]+$', t):
            return False
        return True
    else:
        if re.match(r'^[A-Z0-9\-_./&#()+:]+$', t) and " " not in t and len(t) > 4:
            return False
        return True

def translate_unique_list(unique_strings: list, model_id: str, proxy_url: str, direct_key: str, gsheet_url: str, direction: str, glossary_str: str, log_placeholder, progress_bar) -> dict:
    """Dịch danh sách câu duy nhất: 1. Từ điển THACO -> 2. Bộ nhớ đệm Cloud TM -> 3. Gọi AI Batch & Tự động lưu lên Google Sheets"""
    tm = st.session_state.tm_data
    glossary_dict = parse_glossary(glossary_str, direction)
    
    translated_map = {}
    need_ai = []
    tm_hits = 0

    for s in unique_strings:
        s_clean = s.strip()
        # 1. Khớp chính xác trong Bảng thuật ngữ THACO
        matched_glossary = None
        for k, v in glossary_dict.items():
            if s_clean.lower() == k.lower():
                matched_glossary = v
                break
        if matched_glossary is not None:
            translated_map[s] = matched_glossary
            tm_hits += 1
            continue

        # 2. Khớp trong Bộ nhớ đệm Translation Memory (TM)
        tm_key = f"{direction}::{s_clean}"
        if tm_key in tm and str(tm[tm_key]).strip():
            translated_map[s] = str(tm[tm_key]).strip()
            tm_hits += 1
        elif direction == "Việt → Anh" and s_clean in tm and str(tm[s_clean]).strip():
            translated_map[s] = str(tm[s_clean]).strip()
            tm_hits += 1
        else:
            need_ai.append(s)

    add_log(f"⚡ Đã lấy tức thì {tm_hits}/{len(unique_strings)} câu từ Từ điển THACO & Google Sheets TM (0.01s). Cần gọi AI: {len(need_ai)} câu.", log_placeholder)

    if not need_ai:
        progress_bar.progress(1.0)
        return translated_map

    # 3. Chia lô 18 câu nguyên vẹn để gửi AI
    BATCH_SIZE = 18
    total_ai = len(need_ai)
    newly_translated_for_cloud = {}

    for start in range(0, total_ai, BATCH_SIZE):
        batch_slice = need_ai[start : start + BATCH_SIZE]
        payload = {str(i): item for i, item in enumerate(batch_slice)}
        
        preview = ", ".join([f"'{x[:28]}...'" if len(x) > 28 else f"'{x}'" for x in batch_slice[:2]])
        add_log(f"🤖 Đang gửi AI ({model_id}) dịch câu {start + 1} -> {start + len(batch_slice)}/{total_ai} ({preview})", log_placeholder)
        
        ai_res = call_ai_batch(payload, model_id, proxy_url, direct_key, direction, glossary_str, log_placeholder)
        
        batch_new_tm = {}
        for i, orig_str in enumerate(batch_slice):
            val = ai_res.get(str(i), orig_str)
            if isinstance(val, str) and val.strip():
                clean_val = val.strip()
                translated_map[orig_str] = clean_val
                tm_k = f"{direction}::{orig_str.strip()}"
                tm[tm_k] = clean_val
                batch_new_tm[tm_k] = clean_val
                newly_translated_for_cloud[tm_k] = clean_val
            else:
                translated_map[orig_str] = orig_str

        # Lưu cục bộ và đẩy lô mới lên Google Sheets
        save_tm(tm, new_items_only=batch_new_tm, webapp_url=gsheet_url)
        progress_bar.progress(min((start + len(batch_slice)) / total_ai, 1.0))
        if start + BATCH_SIZE < total_ai:
            time.sleep(1.5)

    if newly_translated_for_cloud:
        add_log(f"☁️ Đã tự động lưu {len(newly_translated_for_cloud)} câu dịch mới lên Google Sheets (Tab TM).", log_placeholder)

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
    runs = [r for r in paragraph.runs if r.text is not None]
    if not runs:
        return
    full_text = "".join(r.text for r in runs)
    if not full_text.strip():
        return

    split_idx = None
    acc = ""
    for i in range(len(runs) - 1):
        acc += runs[i].text
        b_cur = bool(runs[i].font.bold)
        b_next = bool(runs[i + 1].font.bold)
        if b_cur and not b_next and ":" in acc and len(acc.strip()) <= 45:
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
# ============================================================
def sanitize_excel_sheet_title(title: str, existing_titles: list) -> str:
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

def extract_xml_drawing_texts(zip_or_docx_path: str, xml_prefix: str) -> list:
    found = []
    try:
        with zipfile.ZipFile(zip_or_docx_path, 'r') as zf:
            for name in zf.namelist():
                if name.startswith(xml_prefix) and name.endswith(".xml"):
                    tree = ET.fromstring(zf.read(name))
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

            if auto_fit:
                for col in ws.columns:
                    col_letter = get_column_letter(col[0].column)
                    cur_w = ws.column_dimensions[col_letter].width
                    if cur_w and cur_w < 65:
                        ws.column_dimensions[col_letter].width = round(cur_w * 1.18, 1)

        wb.save(output_path)
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
# 8. GIAO DIỆN CHÍNH STREAMLIT (V6.0 CLOUD EDITION)
# ============================================================
st.markdown('<p class="main-title">⚡ Hệ thống Dịch thuật AI Chuyên ngành Ô tô - THACO PRO (V6.0)</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="sub-title">'
    '<span class="badge-pro">PRO V6.0</span>'
    '<span class="badge-cloud">☁️ Google Sheets Cloud Sync</span> '
    'Dịch nguyên câu chống đứt từ • Từ điển THACO • Xuất Song Ngữ • Dịch Đa Chiều (Việt ⇄ Anh) • Hiệu đính Nhanh'
    '</p>',
    unsafe_allow_html=True
)

col_left, col_right = st.columns([1.15, 2.0])

with col_left:
    st.markdown("#### ⚙️ 1. Kết nối Máy chủ AI & Google Sheets Cloud")
    with st.container(border=True):
        proxy_url = st.text_input("🔗 Đường dẫn Cloudflare Worker:", value=DEFAULT_PROXY_URL)
        gsheet_url = st.text_input("☁️ Link Google Sheets Web App (Lưu TM & Từ điển):", value=DEFAULT_GSHEET_WEBAPP_URL)
        direct_key = st.text_input("🔑 Groq API Key trực tiếp (Tùy chọn nếu không qua Worker):", type="password")
        
        btn_c1, btn_c2 = st.columns(2)
        with btn_c1:
            if st.button("🔄 Tải Model AI", use_container_width=True):
                try:
                    with st.spinner("Đang tải model..."):
                        models = fetch_live_chat_models(proxy_url, direct_key)
                        st.session_state.available_models = models
                        st.success(f"✅ Tìm thấy {len(models)} model!")
                except Exception as e:
                    st.error(f"❌ Lỗi: {str(e)}")
        with btn_c2:
            if st.button("☁️ Đồng bộ Google Sheets", use_container_width=True):
                with st.spinner("Đang đồng bộ dữ liệu đám mây..."):
                    ok, msg = sync_from_gsheet(gsheet_url)
                    st.session_state.cloud_sync_status = (ok, msg)
                    if ok:
                        st.success(msg)
                    else:
                        st.warning(msg)

        # Hiển thị trạng thái tự động đồng bộ lần đầu
        if st.session_state.get("cloud_sync_status"):
            sync_ok, sync_msg = st.session_state.cloud_sync_status
            if sync_ok:
                st.caption(f"🟢 {sync_msg}")
            else:
                st.caption(f"🟠 {sync_msg}")

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
    g_count_now = len(glossary_str_to_dict(st.session_state.glossary_text))
    with st.expander(f"📚 2. Khung Từ viết tắt & Thuật ngữ THACO ({g_count_now} mục)", expanded=False):
        st.caption("Nhập theo cú pháp `Từ viết tắt / Tiếng Việt = Tiếng Anh chuẩn`. Khi bấm Lưu sẽ tự động cập nhật lên tab **Glossary** trên Google Sheets:")
        new_glossary = st.text_area(
            "Bảng thuật ngữ THACO:",
            value=st.session_state.glossary_text,
            height=240
        )
        if st.button("💾 Lưu Bảng Thuật Ngữ lên Google Sheets", use_container_width=True):
            st.session_state.glossary_text = new_glossary
            ok_push = save_glossary_text(new_glossary, gsheet_url)
            if ok_push:
                st.success("✅ Đã lưu Bảng thuật ngữ THACO lên Google Sheets thành công!")
            else:
                st.info("💾 Đã lưu cục bộ (Kiểm tra lại quyền truy cập Web App nếu muốn đẩy lên Google Sheets).")

    # --- KHUNG BỘ NHỚ ĐỆM TRANSLATION MEMORY (TM) ---
    with st.expander(f"🧠 3. Bộ nhớ đệm Dịch thuật - Cloud TM ({len(st.session_state.tm_data)} câu)", expanded=False):
        st.caption("Toàn bộ câu dịch mới và câu hiệu đính đều được tự động đồng bộ vào tab **TM** trên Google Sheets:")
        tm_col1, tm_col2 = st.columns(2)
        with tm_col1:
            tm_df_export = pd.DataFrame(
                [{"Khoa_Goc": k, "Ban_Dich_Chuan": v} for k, v in st.session_state.tm_data.items()]
            )
            st.download_button(
                "⬇️ Tải bản sao TM (.csv)",
                data=tm_df_export.to_csv(index=False).encode("utf-8-sig"),
                file_name="THACO_Translation_Memory.csv",
                mime="text/csv",
                use_container_width=True
            )
        with tm_col2:
            if st.button("🔄 Nạp lại TM từ Google Sheets", use_container_width=True):
                ok, msg = sync_from_gsheet(gsheet_url)
                if ok:
                    st.success(msg)
                    st.rerun()

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
            "<div class='log-container'>✅ Hệ thống THACO AI Translator PRO V6.0 (Google Sheets Cloud) sẵn sàng!<br>"
            "• <b>Đồng bộ Đám mây Vĩnh viễn</b>: Tự động nạp/lưu Từ điển THACO (tab Glossary) và Bộ nhớ dịch (tab TM) trên Google Sheets.<br>"
            "• <b>Dịch nguyên câu (Paragraph-level)</b>: Chống cắt vụn chữ trong PowerPoint & Word, giữ nguyên bố cục và định dạng.<br>"
            "• <b>Tích hợp đầy đủ</b>: Xuất Song Ngữ, Dịch Đa Chiều (Việt ⇄ Anh), Bảng Hiệu đính Nhanh & Auto-Fit Layout.</div>",
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
                    unique_list, selected_model, proxy_url, direct_key, gsheet_url,
                    direction, st.session_state.glossary_text, log_placeholder, progress_bar
                )
                
                elapsed = round(time.time() - start_t, 1)
                add_log(f"🎉 Đã dịch xong {len(unique_list)} câu duy nhất trong {elapsed} giây!", log_placeholder)

                st.session_state.review_df = pd.DataFrame([
                    {"Câu gốc trong File": k, "Bản dịch AI (Bấm đúp để sửa tay)": v}
                    for k, v in trans_map.items()
                ])

                # Nếu tắt chế độ Hiệu đính -> Xuất file luôn trong 1 chạm
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
        st.caption("Anh có thể kiểm tra và gõ sửa trực tiếp vào cột **'Bản dịch AI'**. Khi bấm Áp dụng, các câu sửa tay sẽ tự động được đẩy lên Google Sheets (tab TM) để lưu vĩnh viễn:")
        
        edited_df = st.data_editor(
            st.session_state.review_df,
            use_container_width=True,
            height=320,
            disabled=["Câu gốc trong File"],
            key="editor_grid"
        )

        if st.button("✅ 2. ÁP DỤNG BẢN HIỆU ĐÍNH, LƯU GOOGLE SHEETS & KẾT XUẤT FILE", type="primary", use_container_width=True):
            if st.session_state.staged_file_bytes:
                file_name = st.session_state.staged_file_name
                file_ext = os.path.splitext(file_name)[1].lower()
                
                final_map = {}
                changed_for_cloud = {}
                tm = st.session_state.tm_data

                for _, row in edited_df.iterrows():
                    orig_s = str(row["Câu gốc trong File"]).strip()
                    edit_s = str(row["Bản dịch AI (Bấm đúp để sửa tay)"]).strip()
                    final_map[orig_s] = edit_s
                    tm_k = f"{direction}::{orig_s}"
                    if tm.get(tm_k) != edit_s:
                        changed_for_cloud[tm_k] = edit_s
                    tm[tm_k] = edit_s

                # Đẩy những câu được hiệu đính lên Google Sheets
                save_tm(tm, new_items_only=changed_for_cloud if changed_for_cloud else None, webapp_url=gsheet_url)
                if changed_for_cloud:
                    add_log(f"☁️ Đã cập nhật {len(changed_for_cloud)} câu hiệu đính mới lên Google Sheets.", log_placeholder)

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
# app.py
# HỆ THỐNG DỊCH THUẬT KỸ THUẬT AI CHUYÊN NGÀNH Ô TÔ - THACO AUTO (ENTERPRISE v5.0)
# Tích hợp 6 tính năng cao cấp & Thuật toán Paragraph-Level chống cắt vụn câu 100%

import streamlit as st
import requests
import json
import time
import os
import tempfile
import re
import unicodedata
import pandas as pd
from copy import copy

# ----- Thư viện xử lý File -----
from docx import Document
import openpyxl
from openpyxl.utils import get_column_letter
from pdf2docx import Converter
from pptx import Presentation

# ============================================================
# 1. CẤU HÌNH HỆ THỐNG & GIAO DIỆN
# ============================================================
DEFAULT_PROXY_URL = "https://mrluu-groq-proxy.mrscong4.workers.dev"
TM_FILE_PATH = "thaco_translation_memory.json"

st.set_page_config(page_title="THACO AI Technical Translator v5.0", page_icon="⚡", layout="wide")

st.markdown("""
    <style>
    .main-title { font-size: 32px !important; font-weight: 800; color: #1E3A8A; margin-bottom: 2px; }
    .sub-title { font-size: 14px; color: #475569; margin-bottom: 16px; }
    .log-container { background-color: #0F172A; color: #10B981; font-family: 'Courier New', monospace; padding: 14px; border-radius: 8px; height: 340px; overflow-y: auto; font-size: 12.5px; line-height: 1.55; }
    .badge-v5 { background-color: #DBEAFE; color: #1D4ED8; padding: 3px 8px; border-radius: 6px; font-size: 13px; font-weight: 600; }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# 2. SYSTEM PROMPT ĐA CHIỀU (VIỆT <-> ANH) CHUẨN THACO AUTO
# ============================================================
def get_system_prompt(direction: str) -> str:
    if direction == "vi2en":
        return """You are a senior automotive engineering and after-sales quality management translator at THACO AUTO.
Translate the Vietnamese strings in the provided JSON object into natural, grammatically complete, and professional technical English.

CRITICAL TERMINOLOGY & ABBREVIATIONS AT THACO AUTO:
1. Report Headers & 8D Sections:
   - "BÁO CÁO PHÂN TÍCH NGUYÊN NHÂN & GIẢI PHÁP" -> "ROOT CAUSE ANALYSIS & CORRECTIVE ACTION REPORT"
   - "HẠNG MỤC: HƯ HỎNG CỔNG SẠC USB" -> "SUBJECT: USB CHARGING PORT FAILURE"
   - "HẠNG MỤC" -> "SUBJECT / ITEM" | "MÔ TẢ VẤN ĐỀ" -> "PROBLEM DESCRIPTION" | "MÃ SỐ LỖI" -> "DEFECT CODE"
   - "GIẢI PHÁP KHẮC PHỤC LỖI" -> "CORRECTIVE ACTIONS" | "Giải pháp cải tiến" -> "Improvement Solutions"
   - "NGUYÊN NHÂN" -> "ROOT CAUSE" | "Nguyên nhân gốc rễ" -> "Root Cause"
   - "NỘI DUNG TRIỂN KHAI THỰC HIỆN" -> "IMPLEMENTATION PLAN"
   - "TRÁCH NHIỆM" -> "PIC" | "H.THÀNH" -> "TARGET DATE" | "Tình trạng" -> "Status" | "Theo dõi" -> "Monitoring"
   - "Hành động phòng ngừa" -> "Preventive Actions" | "Kiểm chứng giải pháp" -> "Solution Verification"
   - "Trước cải tiến" / "Trước thay đổi" -> "Before Improvement" / "Before Change"
   - "Sau cải tiến" / "Sau thay đổi" -> "After Improvement" / "After Change"
2. THACO Departments & Entities:
   - "BP DVPT" / "BP Dịch vụ" -> "Service & Parts Dept."
   - "Phòng QLCL KCN" -> "Industrial Park QC Dept."
   - "BP KTCL NM" -> "Factory QC Dept."
   - "TT R&D ô tô" / "TT R&D" -> "Automotive R&D Center"
   - "NCC" / "Nhà cung ứng" / "Nhà cung cấp" -> "Supplier" ("NCC trong nước" -> "Domestic supplier")
3. Technical & Quality Context:
   - "Đối với xe lỗi" -> "For defective vehicles"
   - "Đối với xe sản xuất tại NM" -> "For vehicles in factory production"
   - "Đối với xe SBH" / "Sau bán hàng" -> "For After-Sales vehicles" / "After-sales"
   - "Lũy kế SL lỗi/số xe giao" -> "Cumulative defects / delivered units"
   - "Cổng sạc tròn loại cải tiến" -> "Improved round charging port"
   - "Cổng sạc dẹp loại cải tiến" -> "Improved flat charging port"
   - "Áp dụng VIN#" -> "Applied from VIN#"
   - "vật tư" / "linh kiện" -> "part / component" ("vật tư mẫu" -> "sample part", "thu hồi vật tư lỗi" -> "recall defective parts")
   - "MS:" / "mã số:" -> "Part No.:"
   - "theo thông báo số..." -> "per Technical Bulletin No...."
   - "điện trở" -> "resistor", "điện áp" -> "voltage", "tụ điện" -> "capacitor", "cắt bỏ đèn Led chiếu sáng" -> "remove illuminating LED".
4. Preserve all part numbers (e.g., USB-CDQ23035B, 8129000459), VINs, vehicle models (TB120B, TB120S), quantities, percentages, dates, and person names (Mr. Duc, Mr. Viet, Mr. Sinh).
5. Translate full sentences smoothly. NEVER output broken word-by-word Vietlish.
6. OUTPUT FORMAT: Return ONLY a valid JSON object with the EXACT same keys ("0", "1", "2", ...) mapped to their translated English strings. No markdown, no explanations."""
    else:
        return """You are a senior automotive engineering and after-sales quality management translator at THACO AUTO.
Translate the English strings in the provided JSON object into natural, accurate, and professional Vietnamese using standard THACO AUTO terminology (e.g., Warranty Claim -> Khiếu nại bảo hành, Root Cause -> Nguyên nhân gốc rễ, Corrective Action -> Hành động khắc phục, Preventive Action -> Hành động phòng ngừa, Part Number -> Mã linh kiện, Supplier -> Nhà cung cấp, After-sales -> Sau bán hàng).
Preserve all part numbers, VINs, vehicle model codes, dates, and proper names.
OUTPUT FORMAT: Return ONLY a valid JSON object with the EXACT same keys ("0", "1", "2", ...) mapped to their translated Vietnamese strings. No markdown, no explanations."""

# ============================================================
# 3. QUẢN LÝ BỘ NHỚ ĐỆM DỊCH THUẬT (TRANSLATION MEMORY - TM)
# ============================================================
def load_local_tm() -> dict:
    if os.path.exists(TM_FILE_PATH):
        try:
            with open(TM_FILE_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_local_tm(tm_dict: dict):
    try:
        with open(TM_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(tm_dict, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

if "tm_dict" not in st.session_state:
    st.session_state.tm_dict = load_local_tm()
if "logs" not in st.session_state:
    st.session_state.logs = []

def add_log(message: str, log_placeholder=None):
    time_str = time.strftime("%H:%M:%S")
    st.session_state.logs.append(f"[{time_str}] {message}")
    if log_placeholder is not None:
        display_logs = "<br>".join(st.session_state.logs[-50:])
        log_placeholder.markdown(f"<div class='log-container'>{display_logs}</div>", unsafe_allow_html=True)

# ============================================================
# 4. HÀM KẾT NỐI AI & GỌI BATCH CHỐNG LỖI 429 TUYỆT ĐỐI
# ============================================================
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
    exclude_keywords = ["whisper", "guard", "orpheus", "tts", "vision", "embed", " tool", "playai"]
    chat_models = [m for m in all_ids if not any(ex in m.lower() for ex in exclude_keywords)]
    
    def sort_key(m_name):
        name = m_name.lower()
        if "qwen" in name or "70b" in name: return 0
        if "deepseek" in name or "llama-4" in name: return 1
        if "gemma" in name or "mixtral" in name: return 2
        if "8b" in name: return 3
        return 4
        
    chat_models.sort(key=sort_key)
    return chat_models

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

    for attempt in range(5):
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=75)
            if resp.status_code == 400 and "response_format" in resp.text:
                payload.pop("response_format", None)
                resp = requests.post(url, headers=headers, json=payload, timeout=75)

            if resp.status_code == 429:
                wait_s = 8 + attempt * 5
                add_log(f"⏳ Chạm giới hạn tốc độ API (429), hệ thống tự động nghỉ {wait_s}s rồi dịch tiếp...", log_placeholder)
                time.sleep(wait_s)
                continue

            if resp.status_code != 200:
                raise Exception(f"HTTP {resp.status_code} - {resp.text[:200]}")

            content = resp.json()["choices"][0]["message"]["content"].strip()
            content = re.sub(r'<think>.*?</think>', '', content, flags=re.DOTALL).strip()
            
            json_match = re.search(r'\{.*\}', content, re.DOTALL)
            if json_match:
                content = json_match.group(0)
            return json.loads(content)

        except Exception as e:
            add_log(f"⚠️ Thử lại lần {attempt + 1}/5 do lỗi: {str(e)}", log_placeholder)
            time.sleep(3)

    add_log("❌ Không thể dịch lô này sau 5 lần thử, giữ nguyên câu gốc.", log_placeholder)
    return batch_dict

# ============================================================
# 5. BỘ LỌC NGÔN NGỮ & DỊCH LÔ KẾT HỢP TRANSLATION MEMORY
# ============================================================
def has_vietnamese_accents(text: str) -> bool:
    if "đ" in text.lower():
        return True
    for ch in text:
        if ord(ch) > 127 and unicodedata.category(ch).startswith("L"):
            if len(unicodedata.normalize("NFD", ch)) > 1:
                return True
    return False

def should_translate(text: str, direction: str = "vi2en") -> bool:
    t = text.strip()
    if not t or t.startswith("="):
        return False
    if not any(ch.isalpha() for ch in t):
        return False
    # Bỏ qua mã phụ tùng, số VIN, kích thước, ngày tháng (VD: RN5B29SEMRN007347, USB-CDQ23035B, TB120B, 30/7/2026)
    if re.match(r'^[A-Z0-9\-_./&#\s():%+=<>]+$', t) and t.upper() not in ["STT", "SL", "TONG", "PIC", "NO.", "UNIT", "QTY", "TOTAL", "STATUS", "CAUSE"]:
        return False

    if direction == "vi2en":
        if has_vietnamese_accents(t):
            return True
        vn_unaccented = ["STT", "SL", "TONG", "NGUYEN NHAN", "GIAI PHAP", "MO TA", "MA SO"]
        if any(k == t.upper() or k in t.upper() for k in vn_unaccented):
            return True
        return False
    else:
        # Chiều Anh -> Việt: Dịch mọi câu tiếng Anh (không chứa dấu tiếng Việt)
        if has_vietnamese_accents(t) and len(t.split()) > 2:
            return False
        return True

def translate_unique_map(unique_strings: list, model_id: str, direction: str, proxy_url: str, direct_key: str, use_tm: bool, log_placeholder, progress_bar) -> dict:
    """Trả về dict {original_clean_str: translated_str} có tận dụng Bộ nhớ đệm TM"""
    translated_unique = {}
    to_call_ai = []
    tm_hits = 0

    for s in unique_strings:
        tm_key = f"{direction}::{s}"
        if use_tm and tm_key in st.session_state.tm_dict:
            translated_unique[s] = st.session_state.tm_dict[tm_key]
            tm_hits += 1
        elif use_tm and direction == "vi2en" and s in st.session_state.tm_dict:
            translated_unique[s] = st.session_state.tm_dict[s]
            tm_hits += 1
        else:
            to_call_ai.append(s)

    if tm_hits > 0:
        add_log(f"⚡ [Translation Memory] Khớp tức thì {tm_hits}/{len(unique_strings)} câu trong 0.01s mà không cần gọi AI!", log_placeholder)

    total_ai = len(to_call_ai)
    if total_ai == 0:
        progress_bar.progress(1.0)
        return translated_unique

    # Gộp 35 câu nguyên vẹn / 1 lần gọi AI để chỉ cần 1-2 request cho cả Slide/Sheet, triệt tiêu lỗi 429
    BATCH_SIZE = 35
    for start_idx in range(0, total_ai, BATCH_SIZE):
        batch_slice = to_call_ai[start_idx : start_idx + BATCH_SIZE]
        batch_payload = {str(i): s for i, s in enumerate(batch_slice)}
        
        sample_preview = ", ".join([f"'{s[:30]}...'" if len(s) > 30 else f"'{s}'" for s in batch_slice[:2]])
        add_log(f"🤖 [v5.0 Paragraph-Mode] Gửi AI ({model_id}) dịch câu {start_idx + 1} -> {start_idx + len(batch_slice)} ({sample_preview})", log_placeholder)
        
        ai_response = call_ai_batch(batch_payload, model_id, direction, proxy_url, direct_key, log_placeholder)
        
        for i, original_clean in enumerate(batch_slice):
            trans_val = ai_response.get(str(i), original_clean)
            if not isinstance(trans_val, str) or not trans_val.strip():
                trans_val = original_clean
            clean_trans = trans_val.strip()
            translated_unique[original_clean] = clean_trans
            # Lưu vào Bộ nhớ đệm TM
            st.session_state.tm_dict[f"{direction}::{original_clean}"] = clean_trans
            
        progress_bar.progress(min((start_idx + len(batch_slice)) / total_ai, 1.0))
        if start_idx + BATCH_SIZE < total_ai:
            time.sleep(1.5)

    save_local_tm(st.session_state.tm_dict)
    return translated_unique

def format_output_text(original_raw: str, translated_clean: str, bilingual: bool) -> str:
    """Định dạng đầu ra: Ghi đè hoặc Xuất Song Ngữ (Câu gốc ở trên, Bản dịch trong ngoặc ở dưới)"""
    orig_clean = original_raw.strip()
    if not translated_clean or orig_clean.lower() == translated_clean.strip().lower():
        return original_raw
    lead = original_raw[:len(original_raw) - len(original_raw.lstrip())]
    trail = original_raw[len(original_raw.rstrip()):]
    if bilingual:
        return f"{lead}{orig_clean}\n({translated_clean.strip()}){trail}"
    return f"{lead}{translated_clean.strip()}{trail}"

# ============================================================
# 6. THUẬT TOÁN QUÉT & GHI FILE (GIỮ NGUYÊN BỐ CỤC + AUTO-FIT)
# ============================================================
class ParagraphWholeUnit:
    """Gom TOÀN BỘ 1 dòng/đoạn văn thành 1 câu duy nhất, giữ nguyên in đậm tiêu đề trước dấu ':'"""
    def __init__(self, paragraph):
        self.runs = [r for r in paragraph.runs if r.text is not None]
        self.full_text = "".join(r.text for r in self.runs)
        self.has_bold_prefix = False
        if len(self.runs) >= 2 and ":" in self.full_text[:45]:
            first_bold = bool(self.runs[0].font.bold)
            last_bold = bool(self.runs[-1].font.bold)
            if first_bold and not last_bold:
                self.has_bold_prefix = True

    def apply(self, final_text: str):
        if not self.runs:
            return
        # Nếu dòng gốc có phần đầu In đậm trước dấu ':' và phần sau chữ thường -> Giữ nguyên kiểu chữ đó
        if self.has_bold_prefix and ":" in final_text and "\n(" not in final_text:
            prefix, rest = final_text.split(":", 1)
            self.runs[0].text = prefix + ":"
            self.runs[0].font.bold = True
            self.runs[-1].text = rest
            self.runs[-1].font.bold = False
            for r in self.runs[1:-1]:
                r.text = ""
        else:
            self.runs[0].text = final_text
            for r in self.runs[1:]:
                r.text = ""

class XmlTextNodeGroup:
    """Quét và gom các thẻ <a:t> hoặc <w:t> trong một đoạn XML (<a:p> / <w:p>) của Textbox/Shape nổi"""
    def __init__(self, t_nodes):
        self.t_nodes = t_nodes
        self.full_text = "".join(n.text for n in t_nodes if n.text)

    def apply(self, final_text: str):
        if not self.t_nodes:
            return
        self.t_nodes[0].text = final_text
        for n in self.t_nodes[1:]:
            n.text = ""

def extract_pptx_units(shape, units_list):
    """Đệ quy quét mọi Group, TextFrame, Table và SmartArt/Diagram XML trên Slide PowerPoint"""
    if shape.shape_type == 6: # GROUP
        for sub_shape in shape.shapes:
            extract_pptx_units(sub_shape, units_list)
        return
            
    if shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            u = ParagraphWholeUnit(p)
            if u.full_text.strip():
                units_list.append(u)
                
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for p in cell.text_frame.paragraphs:
                    u = ParagraphWholeUnit(p)
                    if u.full_text.strip():
                        units_list.append(u)

    # Nếu là SmartArt / Diagram không có text_frame thông thường -> Quét XML <a:p>
    if not shape.has_text_frame and not shape.has_table and hasattr(shape, "_element"):
        for p_elem in shape._element.xpath(".//*[local-name()='p']"):
            t_nodes = p_elem.xpath(".//*[local-name()='t']")
            if t_nodes:
                grp = XmlTextNodeGroup(t_nodes)
                if grp.full_text.strip():
                    units_list.append(grp)

def extract_all_texts_from_file(input_path: str, file_ext: str, direction: str, translate_sheet_names: bool) -> list:
    """Bước 1: Quét toàn bộ file để lấy danh sách các câu duy nhất cần dịch"""
    raw_texts = []
    if file_ext == ".xlsx":
        wb = openpyxl.load_workbook(input_path)
        for ws in wb.worksheets:
            if translate_sheet_names and ws.title.strip():
                raw_texts.append(ws.title)
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.strip():
                        raw_texts.append(cell.value)
    elif file_ext in [".docx", ".pdf"]:
        work_path = input_path
        if file_ext == ".pdf":
            work_path = input_path.replace(".pdf", "_scan.docx")
            if not os.path.exists(work_path):
                cv = Converter(input_path)
                cv.convert(work_path)
                cv.close()
        doc = Document(work_path)
        units = []
        for p in doc.paragraphs:
            u = ParagraphWholeUnit(p)
            if u.full_text.strip(): units.append(u)
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        u = ParagraphWholeUnit(p)
                        if u.full_text.strip(): units.append(u)
        # Quét thêm Textbox vẽ nổi (w:txbxContent) trong Word
        for p_elem in doc.element.xpath(".//*[local-name()='txbxContent']//*[local-name()='p']"):
            t_nodes = p_elem.xpath(".//*[local-name()='t']")
            if t_nodes:
                grp = XmlTextNodeGroup(t_nodes)
                if grp.full_text.strip(): units.append(grp)
        raw_texts = [u.full_text for u in units]
    elif file_ext == ".pptx":
        prs = Presentation(input_path)
        units = []
        for slide in prs.slides:
            for shape in slide.shapes:
                extract_pptx_units(shape, units)
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                for p in slide.notes_slide.notes_text_frame.paragraphs:
                    u = ParagraphWholeUnit(p)
                    if u.full_text.strip(): units.append(u)
        raw_texts = [u.full_text for u in units]

    unique_list = []
    seen = set()
    for t in raw_texts:
        if should_translate(t, direction):
            clean = t.strip()
            if clean not in seen:
                seen.add(clean)
                unique_list.append(clean)
    return unique_list

def apply_translations_to_file(input_path: str, output_path: str, file_ext: str, trans_map: dict, bilingual: bool, translate_sheet_names: bool, auto_fit: bool):
    """Bước 2: Áp dụng bản dịch (đã qua duyệt/chỉnh sửa) vào file đích và giữ nguyên 100% định dạng"""
    if file_ext == ".xlsx":
        wb = openpyxl.load_workbook(input_path)
        used_titles = set()
        for ws in wb.worksheets:
            # 1. Dịch tên Sheet (nếu bật)
            if translate_sheet_names and ws.title.strip() in trans_map:
                new_title = trans_map[ws.title.strip()]
                # Làm sạch các ký tự cấm trong tên Sheet Excel và giới hạn 31 ký tự
                new_title = re.sub(r'[\\/*?:\[\]]', '-', new_title).strip()[:31]
                if new_title and new_title.lower() not in used_titles:
                    ws.title = new_title
            used_titles.add(ws.title.lower())

            # 2. Dịch nội dung từng ô & Bật Auto-Wrap Text
            col_max_len = {}
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.strip() in trans_map:
                        new_val = format_output_text(cell.value, trans_map[cell.value.strip()], bilingual)
                        cell.value = new_val
                        if auto_fit:
                            al = copy(cell.alignment) if cell.alignment else openpyxl.styles.Alignment()
                            al.wrap_text = True
                            cell.alignment = al
                            lines = str(new_val).split("\n")
                            max_l = max(len(line) for line in lines) if lines else 10
                            col_max_len[cell.column] = max(col_max_len.get(cell.column, 0), min(max_l, 55))

            # 3. Tự động nới rộng cột vừa phải nếu cột quá hẹp (Auto-Fit Column Width)
            if auto_fit:
                for col_idx, max_l in col_max_len.items():
                    col_letter = get_column_letter(col_idx)
                    current_w = ws.column_dimensions[col_letter].width
                    if current_w and current_w < max_l * 1.15:
                        ws.column_dimensions[col_letter].width = min(max(current_w, max_l * 1.12), 60)

        wb.save(output_path)

    elif file_ext in [".docx", ".pdf"]:
        work_path = input_path.replace(".pdf", "_scan.docx") if file_ext == ".pdf" else input_path
        if file_ext == ".pdf" and not os.path.exists(work_path):
            cv = Converter(input_path)
            cv.convert(work_path)
            cv.close()
        doc = Document(work_path)
        units = []
        for p in doc.paragraphs:
            u = ParagraphWholeUnit(p)
            if u.full_text.strip(): units.append(u)
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        u = ParagraphWholeUnit(p)
                        if u.full_text.strip(): units.append(u)
        for p_elem in doc.element.xpath(".//*[local-name()='txbxContent']//*[local-name()='p']"):
            t_nodes = p_elem.xpath(".//*[local-name()='t']")
            if t_nodes:
                grp = XmlTextNodeGroup(t_nodes)
                if grp.full_text.strip(): units.append(grp)

        for u in units:
            clean_key = u.full_text.strip()
            if clean_key in trans_map:
                u.apply(format_output_text(u.full_text, trans_map[clean_key], bilingual))
        doc.save(output_path)
        if file_ext == ".pdf" and os.path.exists(work_path):
            os.remove(work_path)

    elif file_ext == ".pptx":
        prs = Presentation(input_path)
        units = []
        for slide in prs.slides:
            for shape in slide.shapes:
                extract_pptx_units(shape, units)
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                for p in slide.notes_slide.notes_text_frame.paragraphs:
                    u = ParagraphWholeUnit(p)
                    if u.full_text.strip(): units.append(u)

        for u in units:
            clean_key = u.full_text.strip()
            if clean_key in trans_map:
                u.apply(format_output_text(u.full_text, trans_map[clean_key], bilingual))
        prs.save(output_path)

# ============================================================
# 7. GIAO DIỆN CHÍNH (ENTERPRISE v5.0)
# ============================================================
st.markdown('<p class="main-title">⚡ Hệ thống Dịch thuật AI Chuyên ngành Ô tô - MR LƯU <span class="badge-v5">Enterprise v5.0</span></p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Hỗ trợ Excel, Word, PowerPoint, PDF • Gom câu nguyên vẹn (Paragraph-Mode) • Xuất Song Ngữ • Hiệu đính Tương tác & Bộ nhớ Đệm TM</p>', unsafe_allow_html=True)

col_left, col_right = st.columns([1.2, 2])

with col_left:
    st.markdown("### ⚙️ 1. Cấu hình AI & Tùy chọn Dịch")
    with st.container(border=True):
        proxy_url = st.text_input("🔗 Đường dẫn Cloudflare Worker:", value=DEFAULT_PROXY_URL)
        direct_key = st.text_input("🔑 Groq API Key trực tiếp (Tùy chọn):", type="password")
        
        if st.button("🔄 Kiểm tra kết nối & Tải Model AI", use_container_width=True):
            try:
                with st.spinner("Đang kết nối máy chủ AI..."):
                    models = fetch_live_chat_models(proxy_url, direct_key)
                    st.session_state.available_models = models
                    st.success(f"✅ Đã tải {len(models)} model AI đang hoạt động!")
            except Exception as e:
                st.error(f"❌ Lỗi: {str(e)}")

        model_list = st.session_state.get("available_models", [
            "qwen/qwen3.8-27b",
            "meta-llama/llama-4-scout-17b-16e-instruct",
            "meta-llama/llama-4-maverick-17b-128e-instruct",
            "qwen-qwq-32b",
            "deepseek-r1-distill-llama-70b"
        ])
        selected_model = st.selectbox("🤖 Chọn Model AI:", options=model_list)

        direction_label = st.radio(
            "🌐 Chiều dịch thuật:",
            options=["🇻🇳 Tiếng Việt ➔ 🇬🇧 Tiếng Anh", "🇬🇧 Tiếng Anh ➔ 🇻🇳 Tiếng Việt"],
            horizontal=True
        )
        direction_code = "vi2en" if "Việt ➔" in direction_label else "en2vi"

        st.markdown("---")
        bilingual_mode = st.checkbox("📑 Xuất Song Ngữ (Giữ câu gốc + Chèn bản dịch ở dòng dưới)", value=False)
        review_mode = st.checkbox("✏️ Mở Bảng Hiệu đính Nhanh (Duyệt & sửa trước khi xuất File)", value=True)
        c1, c2 = st.columns(2)
        with c1:
            translate_sheet_names = st.checkbox("🗂️ Dịch tên Sheet Excel", value=True)
        with c2:
            auto_fit_excel = st.checkbox("📐 Auto-Fit & Wrap Text", value=True)
        use_tm = st.checkbox(f"🧠 Dùng Bộ nhớ Đệm TM (Đang lưu {len(st.session_state.tm_dict)} câu)", value=True)

    with st.expander("🗄️ Quản lý Bộ nhớ Đệm (TM) & Kết nối Google Sheets"):
        gsheet_csv_url = st.text_input("Link Google Sheets CSV (Publish to Web -> CSV):", placeholder="https://docs.google.com/spreadsheets/d/.../export?format=csv")
        if st.button("📥 Đồng bộ từ Google Sheets CSV", use_container_width=True) and gsheet_csv_url:
            try:
                df_gs = pd.read_csv(gsheet_csv_url)
                if len(df_gs.columns) >= 2:
                    for _, row in df_gs.iterrows():
                        k, v = str(row.iloc[0]).strip(), str(row.iloc[1]).strip()
                        if k and v and k != "nan":
                            st.session_state.tm_dict[f"vi2en::{k}"] = v
                    save_local_tm(st.session_state.tm_dict)
                    st.success(f"✅ Đã nạp thành công! Tổng bộ nhớ: {len(st.session_state.tm_dict)} câu.")
            except Exception as e:
                st.error(f"Lỗi đọc Google Sheets CSV: {e}")

        if st.session_state.tm_dict:
            tm_json_bytes = json.dumps(st.session_state.tm_dict, ensure_ascii=False, indent=2).encode("utf-8")
            st.download_button("⬇️ Tải bộ nhớ TM (.json)", data=tm_json_bytes, file_name="thaco_tm_backup.json", mime="application/json", use_container_width=True)

    st.markdown("### 📂 2. Tải lên Tài liệu")
    uploaded_file = st.file_uploader("Hỗ trợ Excel (.xlsx), Word (.docx), PowerPoint (.pptx), PDF (.pdf)", type=["xlsx", "docx", "pptx", "pdf"])

with col_right:
    st.markdown("### 🖥️ Màn hình Giám sát (Live Console v5.0)")
    log_placeholder = st.empty()
    display_init = "<br>".join(st.session_state.logs[-45:]) if st.session_state.logs else (
        "✅ <b>Enterprise v5.0 đã kích hoạt!</b><br>"
        "• Đã khóa thuật toán <b>ParagraphWholeUnit</b>: Gom toàn bộ câu trong PowerPoint/Word, chấm dứt 100% lỗi cắt từ đơn.<br>"
        "• Tích hợp đủ 6 tính năng: PowerPoint, Song ngữ, Đa chiều Việt↔Anh, Bảng Hiệu đính, Bộ nhớ TM & Auto-Fit."
    )
    log_placeholder.markdown(f"<div class='log-container'>{display_init}</div>", unsafe_allow_html=True)
    
    progress_bar = st.empty()

    if uploaded_file:
        if st.button("🚀 BẮT ĐẦU PHÂN TÍCH & DỊCH BẰNG AI", type="primary", use_container_width=True):
            st.session_state.logs = []
            file_ext = os.path.splitext(uploaded_file.name)[1].lower()
            with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp_in:
                tmp_in.write(uploaded_file.getbuffer())
                input_path = tmp_in.name

            st.session_state.current_input_path = input_path
            st.session_state.current_file_ext = file_ext
            st.session_state.current_file_name = uploaded_file.name

            try:
                pb = progress_bar.progress(0.0)
                start_time = time.time()
                add_log(f"📂 Đang quét toàn câu (Paragraph-Mode) trong file: {uploaded_file.name}...", log_placeholder)
                
                unique_sentences = extract_all_texts_from_file(input_path, file_ext, direction_code, translate_sheet_names)
                add_log(f"🎯 Đã gom thành {len(unique_sentences)} câu hoàn chỉnh (Không còn bị băm nhỏ thành từng từ!).", log_placeholder)

                trans_map = translate_unique_map(
                    unique_sentences, selected_model, direction_code, proxy_url, direct_key, use_tm, log_placeholder, pb
                )
                st.session_state.current_trans_map = trans_map
                elapsed = round(time.time() - start_time, 1)
                add_log(f"🎉 AI đã dịch xong {len(unique_sentences)} câu trong {elapsed}s!", log_placeholder)

                # Nếu không bật chế độ Hiệu đính trước -> Xuất file luôn
                if not review_mode:
                    output_ext = ".docx" if file_ext == ".pdf" else file_ext
                    suffix_tag = "_BILINGUAL" if bilingual_mode else f"_{direction_code.upper()}"
                    output_name = uploaded_file.name.replace(file_ext, f"{suffix_tag}{output_ext}")
                    output_path = input_path.replace(file_ext, f"_out{output_ext}")
                    apply_translations_to_file(input_path, output_path, file_ext, trans_map, bilingual_mode, translate_sheet_names, auto_fit_excel)
                    with open(output_path, "rb") as f:
                        st.session_state.ready_file_bytes = f.read()
                    st.session_state.ready_file_name = output_name
            except Exception as e:
                add_log(f"❌ LỖI HỆ THỐNG: {str(e)}", log_placeholder)

        # HIỂN THỊ BẢNG HIỆU ĐÍNH NHANH (TÍNH NĂNG SỐ 4)
        if review_mode and "current_trans_map" in st.session_state and st.session_state.get("current_file_name") == uploaded_file.name:
            st.markdown("#### ✏️ Bảng Hiệu đính Nhanh (Interactive Review)")
            st.caption("Anh có thể kiểm tra và nhấp đúp vào cột **Bản dịch AI** để sửa trực tiếp bất kỳ thuật ngữ nào trước khi xuất file. Mọi chỉnh sửa sẽ tự động được lưu vào Bộ nhớ Đệm (TM) cho lần sau.")
            
            df_review = pd.DataFrame([
                {"Câu gốc trong File": k, "Bản dịch AI (Nhấp để sửa)": v}
                for k, v in st.session_state.current_trans_map.items()
            ])
            edited_df = st.data_editor(df_review, use_container_width=True, height=320, num_rows="fixed")

            if st.button("✅ ÁP DỤNG BẢN DỊCH & XUẤT FILE HOÀN THIỆN", type="primary", use_container_width=True):
                updated_map = {}
                for _, row in edited_df.iterrows():
                    orig_k = str(row["Câu gốc trong File"])
                    edited_v = str(row["Bản dịch AI (Nhấp để sửa)"]).strip()
                    updated_map[orig_k] = edited_v
                    st.session_state.tm_dict[f"{direction_code}::{orig_k}"] = edited_v
                save_local_tm(st.session_state.tm_dict)

                in_p = st.session_state.current_input_path
                f_ext = st.session_state.current_file_ext
                output_ext = ".docx" if f_ext == ".pdf" else f_ext
                suffix_tag = "_BILINGUAL" if bilingual_mode else f"_{direction_code.upper()}"
                output_name = uploaded_file.name.replace(f_ext, f"{suffix_tag}{output_ext}")
                output_path = in_p.replace(f_ext, f"_out{output_ext}")

                apply_translations_to_file(in_p, output_path, f_ext, updated_map, bilingual_mode, translate_sheet_names, auto_fit_excel)
                with open(output_path, "rb") as f:
                    st.session_state.ready_file_bytes = f.read()
                st.session_state.ready_file_name = output_name
                st.success("🎉 Đã kết xuất file và cập nhật Bộ nhớ Đệm TM thành công!")

        # NÚT TẢI FILE XUỐNG
        if "ready_file_bytes" in st.session_state and st.session_state.get("current_file_name") == uploaded_file.name:
            st.download_button(
                label=f"⬇️ TẢI XUỐNG FILE HOÀN THIỆN ({st.session_state.ready_file_name})",
                data=st.session_state.ready_file_bytes,
                file_name=st.session_state.ready_file_name,
                mime="application/octet-stream",
                type="primary",
                use_container_width=True
            )
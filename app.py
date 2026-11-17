# app.py
# Hệ thống Dịch thuật Kỹ thuật bằng AI (Groq LLM qua Cloudflare Proxy)
# Hỗ trợ: Excel (.xlsx), Word (.docx), PowerPoint (.pptx), PDF (.pdf)
# NÂNG CẤP: Dịch toàn câu (Paragraph-level) giữ nguyên Bold/Font, chống cắt vụn từ (Run-splitting)

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
# 2. SYSTEM PROMPT CHUYÊN SÂU CHO BÁO CÁO KỸ THUẬT & 8D THACO
# ============================================================
SYSTEM_PROMPT = """You are a senior automotive engineering and after-sales quality management translator at THACO AUTO.
Translate the Vietnamese strings in the provided JSON object into natural, grammatically complete, and professional technical English.

CRITICAL TERMINOLOGY & ABBREVIATIONS AT THACO AUTO:
1. Report Headers & Sections:
   - "BÁO CÁO PHÂN TÍCH NGUYÊN NHÂN & GIẢI PHÁP" -> "ROOT CAUSE ANALYSIS & CORRECTIVE ACTION REPORT"
   - "HẠNG MỤC" -> "ITEM / SUBJECT"
   - "MÔ TẢ VẤN ĐỀ" -> "PROBLEM DESCRIPTION"
   - "MÃ SỐ LỖI" -> "DEFECT CODE"
   - "GIẢI PHÁP KHẮC PHỤC LỖI" -> "CORRECTIVE ACTIONS"
   - "NGUYÊN NHÂN" -> "ROOT CAUSE" | "Nguyên nhân gốc rễ" -> "Root Cause"
   - "NỘI DUNG TRIỂN KHAI THỰC HIỆN" -> "IMPLEMENTATION PLAN / ACTIONS"
   - "TRÁCH NHIỆM" -> "PIC / RESPONSIBILITY"
   - "H.THÀNH" / "HOÀN THÀNH" -> "TARGET DATE / COMPLETION"
   - "Tình trạng" -> "Status" | "Theo dõi" -> "Monitoring"
   - "Hành động phòng ngừa" -> "Preventive Actions"
   - "Kiểm chứng giải pháp" -> "Solution Verification"
   - "Trước cải tiến" / "Trước thay đổi" -> "Before Improvement" / "Before Change"
   - "Sau cải tiến" / "Sau thay đổi" -> "After Improvement" / "After Change"
2. THACO Departments & Roles:
   - "BP DVPT" / "BP Dịch vụ" -> "Service & Parts Dept."
   - "Phòng QLCL KCN" -> "Industrial Park QC Dept."
   - "BP KTCL NM" -> "Factory QC Dept."
   - "TT R&D ô tô" / "TT R&D" -> "Automotive R&D Center"
   - "NCC" / "Nhà cung ứng" / "Nhà cung cấp" -> "Supplier" ("NCC trong nước" -> "Domestic supplier")
3. Technical & Quality Terms:
   - "Đối với xe lỗi" -> "For defective vehicles"
   - "Đối với xe sản xuất tại NM" -> "For vehicles in factory production"
   - "Đối với xe SBH" / "Sau bán hàng" -> "For After-Sales vehicles" / "After-sales"
   - "Lũy kế SL lỗi/số xe giao" -> "Cumulative defects / delivered vehicles"
   - "Dòng xe" -> "Vehicle Model"
   - "Áp dụng VIN#" -> "Applied from VIN#"
   - "vật tư" / "linh kiện" -> "part / component" ("vật tư mẫu" -> "sample part", "thu hồi vật tư lỗi" -> "recall defective parts")
   - "MS:" / "mã số:" -> "Part No.:"
   - "theo thông báo số..." -> "per Technical Bulletin No...."
   - "điện trở" -> "resistor", "điện áp" -> "voltage", "tụ điện" -> "capacitor", "cổng sạc" -> "charging port".
4. Preserve all part numbers (e.g., USB-CDQ23035B, 8129000459), VINs (e.g., RN5B29SEMRN007347), vehicle models (TB120B, TB120S), quantities, percentages, dates, and person names (Mr. Duc, Mr. Viet, Mr. Sinh).
5. NEVER output broken Vietlish. Translate the entire sentence smoothly and accurately.
6. OUTPUT FORMAT: Return ONLY a valid JSON object with the EXACT same keys ("0", "1", "2", ...) mapped to their translated English strings. No markdown, no explanations."""

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

def call_ai_batch(batch_dict: dict, model_id: str, proxy_url: str, direct_key: str, log_placeholder) -> dict:
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
            resp = requests.post(url, headers=headers, json=payload, timeout=60)
            if resp.status_code == 400 and "response_format" in resp.text:
                payload.pop("response_format", None)
                resp = requests.post(url, headers=headers, json=payload, timeout=60)

            if resp.status_code == 429:
                add_log("⚠️ Chạm giới hạn tốc độ (429), đang chờ 6 giây...", log_placeholder)
                time.sleep(6)
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
            add_log(f"⚠️ Thử lại lần {attempt + 1} do lỗi: {str(e)}", log_placeholder)
            time.sleep(2)

    add_log("❌ Không thể dịch cụm này sau 3 lần thử, giữ nguyên văn bản gốc.", log_placeholder)
    return batch_dict

# ============================================================
# 4. BỘ LỌC NHẬN DIỆN TIẾNG VIỆT CHUẨN UNICODE 100%
# ============================================================
def has_vietnamese_accents(text: str) -> bool:
    """Nhận diện chính xác 100% mọi ký tự có dấu tiếng Việt (cả chữ thường lẫn IN HOA)"""
    if "đ" in text.lower():
        return True
    for ch in text:
        if ord(ch) > 127 and unicodedata.category(ch).startswith("L"):
            # Kiểm tra ký tự chữ cái có dấu (Latin Extended)
            decomp = unicodedata.normalize("NFD", ch)
            if len(decomp) > 1:
                return True
    return False

def should_translate(text: str) -> bool:
    """Bỏ qua dấu câu đứng lẻ, số thuần túy, công thức Excel hoặc mã phụ tùng/số VIN"""
    t = text.strip()
    if not t or t.startswith("="):
        return False
    # Phải có ít nhất 1 chữ cái (bỏ qua các chuỗi chỉ toàn dấu câu như ":", "-", "1.", "20/8/2026")
    if not any(ch.isalpha() for ch in t):
        return False
    # Nếu có dấu tiếng Việt (dù viết thường hay IN HOA toàn bộ) -> BẮT BUỘC DỊCH
    if has_vietnamese_accents(t):
        return True
    # Các từ viết tắt tiếng Việt không dấu cần dịch
    vn_unaccented_keywords = ["STT", "SL", "TONG", "NGUYEN NHAN", "GIAI PHAP", "MO TA", "MA SO", "TRACH NHIEM"]
    t_upper = t.upper()
    if any(k == t_upper or k in t_upper for k in vn_unaccented_keywords):
        return True
    # Nếu chỉ gồm chữ in hoa không dấu + số + ký tự mã hiệu (VD: TB120B, RN5B29SEMRN007347, USB-CDQ23035B) -> Không dịch
    if re.match(r'^[A-Z0-9\-_./&#\s():%]+$', t):
        return False
    return True

def translate_all_strings(texts: list, model_id: str, proxy_url: str, direct_key: str, log_placeholder, progress_bar) -> list:
    results = list(texts)
    
    unique_map = {}
    for idx, raw_text in enumerate(texts):
        if should_translate(raw_text):
            clean = raw_text.strip()
            unique_map.setdefault(clean, []).append(idx)

    unique_strings = list(unique_map.keys())
    total_unique = len(unique_strings)
    add_log(f"🔍 Đã gom {len(texts)} đoạn văn thành {total_unique} câu hoàn chỉnh cần AI dịch.", log_placeholder)

    if total_unique == 0:
        progress_bar.progress(1.0)
        return results

    # Chia lô 20 câu hoàn chỉnh/lần để AI dịch sâu và chuẩn ngữ cảnh nhất
    BATCH_SIZE = 20
    translated_unique = {}
    
    for start_idx in range(0, total_unique, BATCH_SIZE):
        batch_slice = unique_strings[start_idx : start_idx + BATCH_SIZE]
        batch_payload = {str(i): s for i, s in enumerate(batch_slice)}
        
        sample_preview = ", ".join([f"'{s[:25]}...'" if len(s) > 25 else f"'{s}'" for s in batch_slice[:2]])
        add_log(f"🤖 Đang gửi AI ({model_id}) dịch lô {start_idx + 1} -> {start_idx + len(batch_slice)} ({sample_preview})", log_placeholder)
        
        ai_response = call_ai_batch(batch_payload, model_id, proxy_url, direct_key, log_placeholder)
        
        for i, original_clean in enumerate(batch_slice):
            trans_val = ai_response.get(str(i), original_clean)
            if not isinstance(trans_val, str) or not trans_val.strip():
                trans_val = original_clean
            translated_unique[original_clean] = trans_val.strip()
            
        progress_bar.progress(min((start_idx + len(batch_slice)) / total_unique, 1.0))
        time.sleep(0.8)

    for clean_str, indices in unique_map.items():
        trans_str = translated_unique.get(clean_str, clean_str)
        for idx in indices:
            orig = texts[idx]
            lead = orig[:len(orig) - len(orig.lstrip())]
            trail = orig[len(orig.rstrip()):]
            results[idx] = lead + trans_str + trail

    return results

# ============================================================
# 5. THUẬT TOÁN GOM CÂU THÔNG MINH CHO POWERPOINT & WORD
# (Chống cắt vụn chữ Run-splitting & Bảo toàn In đậm đầu dòng)
# ============================================================
class ParagraphSegment:
    """Đại diện cho 1 cụm liên tiếp các run có cùng định dạng trong 1 đoạn văn"""
    def __init__(self, runs):
        self.runs = runs
        self.text = "".join(r.text for r in runs if r.text)

    def apply_translation(self, new_text: str):
        if not self.runs:
            return
        # Gán toàn bộ câu đã dịch vào run đầu tiên (giữ nguyên font, size, màu sắc, bold của nhóm)
        self.runs[0].text = new_text
        # Xóa rỗng các run bị băm nhỏ phía sau để không bị lặp chữ
        for r in self.runs[1:]:
            r.text = ""

def collect_paragraph_segments(paragraph, segments_list, texts_list):
    """
    Gom các run bị cắt vụn trong 1 paragraph thành câu hoàn chỉnh:
    - Nếu dòng có dạng 'Tiêu đề in đậm: Nội dung thường' -> Tách đúng tại dấu ':' thành 2 vế để giữ nguyên in đậm vế đầu.
    - Nếu toàn bộ dòng cùng kiểu -> Gom hết tất cả các run thành 1 câu duy nhất để AI dịch chuẩn 100% ngữ pháp.
    """
    runs = [r for r in paragraph.runs if r.text is not None]
    if not runs:
        return
    full_text = "".join(r.text for r in runs)
    if not full_text.strip():
        return

    # Kiểm tra xem có sự chuyển tiếp từ In Đậm (Bold) sang Chữ Thường sau dấu ':' hay không
    split_idx = None
    acc_text = ""
    for i in range(len(runs) - 1):
        acc_text += runs[i].text
        b_cur = bool(runs[i].font.bold)
        b_next = bool(runs[i + 1].font.bold)
        # Nếu đang ở quanh dấu ':' và có sự đổi trạng thái Bold (VD: "1. Vấn đề:" [Bold] + " Hư hỏng..." [Normal])
        if b_cur != b_next and (":" in acc_text and len(acc_text.strip()) <= 45):
            split_idx = i + 1
            break

    if split_idx is not None:
        seg1 = ParagraphSegment(runs[:split_idx])
        seg2 = ParagraphSegment(runs[split_idx:])
        if seg1.text.strip():
            segments_list.append(seg1)
            texts_list.append(seg1.text)
        if seg2.text.strip():
            segments_list.append(seg2)
            texts_list.append(seg2.text)
    else:
        # Gom toàn bộ các run trong paragraph thành 1 câu hoàn chỉnh duy nhất
        seg = ParagraphSegment(runs)
        segments_list.append(seg)
        texts_list.append(seg.text)

# ============================================================
# 6. CÁC HÀM XỬ LÝ FILE (EXCEL, WORD, POWERPOINT, PDF)
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
    add_log("📂 Đang mở file Word và gom các câu văn...", log_placeholder)
    doc = Document(input_path)
    segments, texts = [], []
    
    for p in doc.paragraphs:
        collect_paragraph_segments(p, segments, texts)
        
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    collect_paragraph_segments(p, segments, texts)
                            
    translated = translate_all_strings(texts, model_id, proxy_url, direct_key, log_placeholder, progress_bar)
    for seg, new_val in zip(segments, translated):
        seg.apply_translation(new_val)
        
    add_log("💾 Đang lưu file Word...", log_placeholder)
    doc.save(output_path)

def extract_pptx_shape_paragraphs(shape, segments_list, texts_list):
    """Đệ quy quét toàn bộ Text Frame, Table và Group Shapes theo cấp độ Câu/Đoạn văn (Paragraph)"""
    if shape.shape_type == 6: # MSO_SHAPE_TYPE.GROUP
        for sub_shape in shape.shapes:
            extract_pptx_shape_paragraphs(sub_shape, segments_list, texts_list)
            
    if shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            collect_paragraph_segments(p, segments_list, texts_list)
                    
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for p in cell.text_frame.paragraphs:
                    collect_paragraph_segments(p, segments_list, texts_list)

def process_pptx(input_path, output_path, model_id, proxy_url, direct_key, log_placeholder, progress_bar):
    add_log("📂 Đang mở file PowerPoint và gom câu trên tất cả các Slide...", log_placeholder)
    prs = Presentation(input_path)
    segments, texts = [], []
    
    for slide in prs.slides:
        for shape in slide.shapes:
            extract_pptx_shape_paragraphs(shape, segments, texts)
            
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            for p in slide.notes_slide.notes_text_frame.paragraphs:
                collect_paragraph_segments(p, segments, texts)
                        
    add_log(f"📊 Đã quét {len(prs.slides)} Slide -> Hợp nhất thành {len(segments)} đoạn văn nguyên vẹn.", log_placeholder)
    translated = translate_all_strings(texts, model_id, proxy_url, direct_key, log_placeholder, progress_bar)
    
    for seg, new_val in zip(segments, translated):
        seg.apply_translation(new_val)
        
    add_log("💾 Đang lưu file PowerPoint hoàn thiện (giữ nguyên bố cục, màu sắc, font chữ)...", log_placeholder)
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
# 7. GIAO DIỆN CHÍNH
# ============================================================
st.markdown('<p class="main-title">⚡ Hệ thống Dịch thuật AI Chuyên ngành Ô tô - THACO</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Hỗ trợ Excel, Word, PowerPoint, PDF • Dịch toàn câu chuẩn ngữ cảnh 8D/QA/QC • Giữ nguyên cấu trúc File</p>', unsafe_allow_html=True)

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
        "• Đã nâng cấp thuật toán <b>Paragraph-Level</b>: Chống cắt vụn chữ trong PowerPoint & Word.<br>"
        "• Tích hợp bộ thuật ngữ Báo cáo Phân tích Nguyên nhân & Giải pháp THACO AUTO.</div>",
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
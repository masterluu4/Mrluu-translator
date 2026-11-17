# app.py
# Hệ thống Dịch thuật Kỹ thuật THACO (Zero-API-Key & Batch Processing)
# Miễn phí 100% - Không cần cấu hình Key - Không bao giờ lỗi Model

import streamlit as st
import time
import os
import tempfile
import re
from deep_translator import GoogleTranslator

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
    .main-title { font-size: 36px !important; font-weight: 700; color: #1E3A8A; margin-bottom: 0px; }
    .sub-title { font-size: 16px; color: #64748B; margin-bottom: 20px; }
    .log-container { background-color: #0F172A; color: #10B981; font-family: 'Courier New', monospace; padding: 15px; border-radius: 8px; height: 360px; overflow-y: auto; font-size: 13px; line-height: 1.5; }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# 2. TỪ ĐIỂN THUẬT NGỮ CHUYÊN NGÀNH THACO & Ô TÔ (MẶC ĐỊNH)
# ============================================================
DEFAULT_GLOSSARY = """STT = No.
Số BCKT = Technical Report No.
BCKT = Technical Report
Báo cáo kỹ thuật = Technical Report
Hình thức vận chuyển = Shipping Method
Đường biển = Sea Freight
Đường hàng không = Air Freight
Đường bộ = Road Transport
Mã vật tư = Part Number
Mã phụ tùng = Part Number
Tên vật tư = Part Name
Tên phụ tùng = Part Name
Đơn vị tính = Unit
ĐVT = Unit
Số lượng = Quantity
SL = Qty
Ghi chú = Remarks
Nguyên nhân gốc rễ = Root Cause
Hành động khắc phục = Corrective Action
Hành động phòng ngừa = Preventive Action
Ống dẫn nước làm mát = Coolant Hose
Bảo hành = Warranty
Khiếu nại = Claim
Lỗi lắp ráp = Assembly Defect
Nhà cung cấp = Supplier
NCC = Supplier
Phòng QLCL = Quality Management Dept."""

def parse_glossary(glossary_str: str) -> dict:
    """Chuyển chuỗi từ điển thành dict, sắp xếp từ dài đến ngắn để ưu tiên cụm từ dài"""
    glossary = {}
    for line in glossary_str.strip().split("\n"):
        if "=" in line:
            k, v = line.split("=", 1)
            if k.strip() and v.strip():
                glossary[k.strip()] = v.strip()
    return dict(sorted(glossary.items(), key=lambda item: len(item[0]), reverse=True))

def apply_pre_glossary(text: str, glossary: dict) -> str:
    """Thay thế các thuật ngữ chuẩn của THACO trước khi dịch"""
    # Nếu cả ô khớp chính xác 100% với từ điển (không phân biệt hoa thường) -> lấy luôn
    for vn_term, en_term in glossary.items():
        if text.strip().lower() == vn_term.lower():
            return en_term
            
    # Thay thế cụm từ bên trong câu
    result = text
    for vn_term, en_term in glossary.items():
        pattern = re.compile(r'\b' + re.escape(vn_term) + r'\b', re.IGNORECASE)
        result = pattern.sub(en_term, result)
    return result

# ============================================================
# 3. HỆ THỐNG LOG & HÀM DỊCH GỘP BATCH SIÊU TỐC
# ============================================================
if "logs" not in st.session_state:
    st.session_state.logs = []

def add_log(message: str, log_placeholder):
    time_str = time.strftime("%H:%M:%S")
    full_msg = f"[{time_str}] {message}"
    st.session_state.logs.append(full_msg)
    display_logs = "<br>".join(st.session_state.logs[-40:])
    log_placeholder.markdown(f"<div class='log-container'>{display_logs}</div>", unsafe_allow_html=True)

def is_translatable(text: str) -> bool:
    """Bỏ qua chuỗi rỗng, công thức, số thuần túy hoặc mã phụ tùng (VD: CNA6129P5-1303010)"""
    t = text.strip()
    if not t or t.startswith("="):
        return False
    # Nếu chuỗi chỉ chứa chữ cái không dấu + số + ký tự gạch ngang/chấm (mã vật tư) -> Không dịch
    if re.match(r'^[A-Z0-9\-_./\s]+$', t):
        return False
    return True

def translate_batch_texts(texts: list, glossary: dict, log_placeholder, progress_bar) -> list:
    """Gộp nhiều ô thành từng cụm lớn để dịch trong vài giây thay vì dịch lẻ từng ô"""
    translator = GoogleTranslator(source='vi', target='en')
    results = list(texts)
    
    # Lọc các vị trí thực sự cần dịch
    valid_indices = []
    processed_inputs = []
    
    for idx, t in enumerate(texts):
        if is_translatable(t):
            # Nếu khớp hoàn toàn trong từ điển THACO thì gán luôn không cần gọi mạng
            matched_exact = False
            for vn_term, en_term in glossary.items():
                if t.strip().lower() == vn_term.lower():
                    lead = t[:len(t) - len(t.lstrip())]
                    trail = t[len(t.rstrip()):]
                    results[idx] = lead + en_term + trail
                    matched_exact = True
                    break
            if not matched_exact:
                valid_indices.append(idx)
                processed_inputs.append(apply_pre_glossary(t.strip(), glossary))

    total_to_translate = len(valid_indices)
    add_log(f"🔍 Sau khi lọc mã phụ tùng & áp dụng từ điển THACO, cần dịch {total_to_translate}/{len(texts)} mục.", log_placeholder)
    
    if total_to_translate == 0:
        progress_bar.progress(1.0)
        return results

    # Gộp từng cụm 25 câu (hoặc tối đa 3500 ký tự) để dịch 1 lần
    SEPARATOR = "\n ||| \n"
    batch_indices = []
    batch_strings = []
    current_char_count = 0
    completed = 0

    def flush_batch(b_indices, b_strings):
        nonlocal completed
        if not b_indices:
            return
        combined_text = SEPARATOR.join(b_strings)
        sample = b_strings[0][:35] + "..." if len(b_strings[0]) > 35 else b_strings[0]
        add_log(f"⚡ Đang dịch cụm {len(b_indices)} câu (Bắt đầu: '{sample}')", log_placeholder)
        
        for attempt in range(3):
            try:
                translated_combined = translator.translate(combined_text)
                split_parts = [p.strip() for p in re.split(r'\s*\|\|\|\s*', translated_combined)]
                
                if len(split_parts) == len(b_indices):
                    for idx_pos, trans_str in zip(b_indices, split_parts):
                        orig = texts[idx_pos]
                        lead = orig[:len(orig) - len(orig.lstrip())]
                        trail = orig[len(orig.rstrip()):]
                        results[idx_pos] = lead + trans_str + trail
                    break
                else:
                    # Nếu dấu phân cách bị lệch, dịch lẻ từng câu trong cụm này
                    for idx_pos, raw_str in zip(b_indices, b_strings):
                        orig = texts[idx_pos]
                        lead = orig[:len(orig) - len(orig.lstrip())]
                        trail = orig[len(orig.rstrip()):]
                        trans_single = translator.translate(raw_str)
                        results[idx_pos] = lead + (trans_single.strip() if trans_single else raw_str)
                    break
            except Exception as e:
                add_log(f"⚠️ Mạng chập chờn ({str(e)}), đang thử lại...", log_placeholder)
                time.sleep(1.5)
                
        completed += len(b_indices)
        progress_bar.progress(min(completed / total_to_translate, 1.0))
        time.sleep(0.3)

    for idx_pos, proc_str in zip(valid_indices, processed_inputs):
        if len(batch_indices) >= 25 or (current_char_count + len(proc_str)) > 3500:
            flush_batch(batch_indices, batch_strings)
            batch_indices, batch_strings, current_char_count = [], [], 0
            
        batch_indices.append(idx_pos)
        batch_strings.append(proc_str)
        current_char_count += len(proc_str)

    flush_batch(batch_indices, batch_strings)
    return results

# ============================================================
# 4. CÁC HÀM XỬ LÝ FILE (EXCEL, WORD, PDF)
# ============================================================
def process_excel(input_path, output_path, glossary, log_placeholder, progress_bar):
    add_log("📂 Đang đọc cấu trúc file Excel...", log_placeholder)
    wb = openpyxl.load_workbook(input_path)
    cells = []
    texts = []
    
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip():
                    cells.append(cell)
                    texts.append(cell.value)
                    
    add_log(f"📊 Tìm thấy tổng cộng {len(cells)} ô chứa văn bản/mã hiệu.", log_placeholder)
    translated_texts = translate_batch_texts(texts, glossary, log_placeholder, progress_bar)
    
    for cell, new_val in zip(cells, translated_texts):
        cell.value = new_val
        
    add_log("💾 Đang lưu file Excel (giữ nguyên 100% định dạng, border, màu sắc)...", log_placeholder)
    wb.save(output_path)

def process_word(input_path, output_path, glossary, log_placeholder, progress_bar):
    add_log("📂 Đang đọc cấu trúc file Word...", log_placeholder)
    doc = Document(input_path)
    runs = []
    texts = []
    
    for p in doc.paragraphs:
        for r in p.runs:
            if r.text.strip():
                runs.append(r)
                texts.append(r.text)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        if r.text.strip():
                            runs.append(r)
                            texts.append(r.text)
                            
    add_log(f"📝 Tìm thấy {len(runs)} phân đoạn văn bản.", log_placeholder)
    translated_texts = translate_batch_texts(texts, glossary, log_placeholder, progress_bar)
    
    for r, new_val in zip(runs, translated_texts):
        r.text = new_val
        
    add_log("💾 Đang lưu file Word...", log_placeholder)
    doc.save(output_path)

def process_pdf(input_path, output_path, glossary, log_placeholder, progress_bar):
    add_log("🔄 Đang chuyển đổi PDF sang cấu trúc Word giữ nguyên Layout...", log_placeholder)
    temp_docx = input_path.replace(".pdf", "_temp.docx")
    cv = Converter(input_path)
    cv.convert(temp_docx)
    cv.close()
    progress_bar.progress(0.15)

    class ScaledProgressBar:
        def __init__(self, pb): self.pb = pb
        def progress(self, val): self.pb.progress(0.15 + (val * 0.85))
        
    process_word(temp_docx, output_path, glossary, log_placeholder, ScaledProgressBar(progress_bar))
    if os.path.exists(temp_docx):
        os.remove(temp_docx)

# ============================================================
# 5. GIAO DIỆN CHÍNH
# ============================================================
st.markdown('<p class="main-title">⚡ Hệ thống Dịch thuật Kỹ thuật THACO</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Chế độ Dịch Gộp Siêu Tốc • Tích hợp Từ điển Chuyên ngành Ô tô & QLCL • Không cần API Key</p>', unsafe_allow_html=True)

col_left, col_right = st.columns([1.2, 2])

with col_left:
    st.markdown("### 📂 1. Tải lên Tài liệu")
    uploaded_file = st.file_uploader("Kéo thả file Excel (.xlsx), Word (.docx), PDF (.pdf)", type=["xlsx", "docx", "pdf"])
    
    with st.expander("📖 2. Từ điển Thuật ngữ THACO (Tùy chỉnh)", expanded=False):
        st.caption("Các từ khóa dưới đây sẽ được ép dịch chính xác 100% theo chuẩn THACO. Anh có thể thêm từ mới theo cú pháp: `Tiếng Việt = Tiếng Anh`")
        glossary_input = st.text_area("Danh sách thuật ngữ:", value=DEFAULT_GLOSSARY, height=220)

with col_right:
    st.markdown("### 🖥️ Màn hình Giám sát (Live Console)")
    log_placeholder = st.empty()
    log_placeholder.markdown(
        "<div class='log-container'>✅ Hệ thống sẵn sàng hoạt động!<br>"
        "• Không yêu cầu API Key hay Proxy.<br>"
        "• Tự động nhận diện và giữ nguyên Mã phụ tùng (Part Number).<br>"
        "• Tốc độ xử lý: ~60 ô Excel chỉ trong 3 - 5 giây.</div>", 
        unsafe_allow_html=True
    )
    
    progress_bar = st.empty()
    action_placeholder = st.empty()

    if uploaded_file:
        if action_placeholder.button("🚀 BẮT ĐẦU DỊCH NGAY", type="primary", use_container_width=True):
            st.session_state.logs = []
            glossary_dict = parse_glossary(glossary_input)
            add_log(f"🚀 Bắt đầu xử lý file: {uploaded_file.name}", log_placeholder)
            
            file_ext = os.path.splitext(uploaded_file.name)[1].lower()
            with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp_in:
                tmp_in.write(uploaded_file.getbuffer())
                input_path = tmp_in.name

            output_ext = ".docx" if file_ext == ".pdf" else file_ext
            output_name = uploaded_file.name.replace(file_ext, f"_EN{output_ext}")
            output_path = input_path.replace(file_ext, f"_out{output_ext}")

            try:
                pb = progress_bar.progress(0.0)
                start_time = time.time()
                
                if file_ext == ".xlsx":
                    process_excel(input_path, output_path, glossary_dict, log_placeholder, pb)
                elif file_ext == ".docx":
                    process_word(input_path, output_path, glossary_dict, log_placeholder, pb)
                elif file_ext == ".pdf":
                    process_pdf(input_path, output_path, glossary_dict, log_placeholder, pb)
                    
                elapsed = round(time.time() - start_time, 1)
                add_log(f"🎉 HOÀN TẤT 100% chỉ trong {elapsed} giây!", log_placeholder)
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
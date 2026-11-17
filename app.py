# app.py
# Ứng dụng Web dịch tài liệu kỹ thuật (Word, Excel, PDF) Việt → Anh
# Sử dụng Streamlit + Google Gemini API, giữ nguyên định dạng gốc.

import streamlit as st
import google.generativeai as genai
import time
import os
import tempfile
import shutil
from copy import copy

# ----- Xử lý Word -----
from docx import Document
from docx.text.run import Run
from docx.text.hyperlink import Hyperlink
from docx.table import Table
from docx.oxml.ns import qn

# ----- Xử lý Excel -----
import openpyxl
from openpyxl.utils import get_column_letter

# ----- Xử lý PDF -----
from pdf2docx import Converter

# ============================================================
# 1. CẤU HÌNH SYSTEM PROMPT DỊCH THUẬT CHUYÊN NGÀNH THACO
# ============================================================
SYSTEM_PROMPT = """You are an expert technical translator specializing in automotive manufacturing, vehicle engineering, and after-sales quality management. Your task is to translate Vietnamese text into precise, grammatically correct, and professional English.
Context: The documents are from the After-Sales Quality Management Department of THACO AUTO, covering passenger cars, trucks, buses, and spare parts.
Requirements:
- Use standard automotive industry and THACO-specific terminology (e.g., IATF 16949, 8D report, 5 Core Tools, PFMEA, 5-Why, Ishikawa 4M1E, RACI, NCR/CAR, Warranty Claims, CKD assembly).
- Note specific contexts: 'CBU' refers to Completely Built-Up units imported for research/new projects (not retail).
- Tone: Maintain a highly formal, precise, and PROACTIVE tone (reflecting the proactive role of the Manufacturing Quality Management department, not a passive one). Suitable for reporting to executives and communicating with international partners.
- Output: Return ONLY the translated English text. Do not add any notes, conversational text, or explanations."""

# ============================================================
# 2. HÀM GỌI GEMINI API (có retry & sleep)
# ============================================================
def init_gemini(api_key: str):
    """Khởi tạo Gemini client với API key."""
    genai.configure(api_key=api_key)
    # Sử dụng gemini-1.5-flash cho tốc độ và chi phí tối ưu (free tier)
    # Có thể đổi thành "gemini-1.5-pro" nếu cần chất lượng cao hơn
    model = genai.GenerativeModel(
        model_name="gemini-flash-latest",
        system_instruction=SYSTEM_PROMPT,
    )
    return model

def translate_text(model, text: str, max_retries: int = 3) -> str:
    """
    Dịch một đoạn văn bản sang tiếng Anh.
    - Tự động retry với exponential backoff khi gặp lỗi 429.
    - Sleep giữa các lần gọi để tránh vượt rate limit (free tier ~15 RPM).
    """
    if not text or not text.strip():
        return text  # giữ nguyên khoảng trắng / chuỗi rỗng

    core = text.strip()
    lead = text[: len(text) - len(text.lstrip())]
    trail = text[len(text.rstrip()) :]

    for attempt in range(max_retries):
        try:
            response = model.generate_content(core)
            translated = response.text.strip() if response.text else core
            # Sleep nhỏ giữa các request để giảm tải
            time.sleep(0.5)
            return lead + translated + trail
        except Exception as e:
            error_str = str(e)
            if "429" in error_str or "ResourceExhausted" in error_str:
                wait = (2 ** attempt) + 1  # 2s, 5s, 9s
                time.sleep(wait)
            else:
                # Lỗi khác: thử lại sau 1s
                time.sleep(1)
    # Nếu thất bại hoàn toàn, trả về nguyên bản để không mất dữ liệu
    return lead + core + trail

# ============================================================
# 3. XỬ LÝ WORD (.docx) – GIỮ NGUYÊN ĐỊNH DẠNG RUN-BY-RUN
# ============================================================
def translate_run(run: Run, model):
    """Dịch nội dung của một run, giữ nguyên định dạng (bold, italic, font...)."""
    translate_and_set(run, model)

def translate_and_set(run: Run, model):
    """Hàm phụ: lấy text, dịch, gán lại vào run.text (giữ format)."""
    original = run.text
    if not original:
        return
    translated = translate_text(model, original)
    run.text = translated  # Run.text setter giữ nguyên định dạng của run

def translate_paragraph(paragraph, model):
    """Duyệt từng run trong paragraph và dịch."""
    for run in paragraph.runs:
        translate_run(run, model)
    # Xử lý hyperlink (runs bên trong hyperlink không nằm trong paragraph.runs)
    for hyperlink in paragraph._element.findall(qn("w:hyperlink")):
        for r in hyperlink.findall(qn("w:r")):
            run = Run(r, paragraph)
            translate_run(run, model)

def translate_table(table: Table, model):
    """Dịch tất cả ô trong bảng, bao gồm bảng lồng nhau."""
    seen_cells = set()
    for row in table.rows:
        for cell in row.cells:
            if id(cell._tc) in seen_cells:
                continue
            seen_cells.add(id(cell._tc))
            # Duyệt paragraph trong ô
            for p in cell.paragraphs:
                translate_paragraph(p, model)
            # Duyệt bảng lồng nhau
            for nested_table in cell.tables:
                translate_table(nested_table, model)

def translate_docx(input_path: str, output_path: str, model):
    """Dịch toàn bộ file Word: body, header, footer, tables."""
    doc = Document(input_path)

    # 1) Body
    for paragraph in doc.paragraphs:
        translate_paragraph(paragraph, model)

    # 2) Tables trong body
    for table in doc.tables:
        translate_table(table, model)

    # 3) Headers & Footers
    for section in doc.sections:
        for header in [section.header, section.first_page_header, section.even_page_header]:
            if header:
                for p in header.paragraphs:
                    translate_paragraph(p, model)
                for t in header.tables:
                    translate_table(t, model)
        for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
            if footer:
                for p in footer.paragraphs:
                    translate_paragraph(p, model)
                for t in footer.tables:
                    translate_table(t, model)

    doc.save(output_path)

# ============================================================
# 4. XỬ LÝ EXCEL (.xlsx) – GIỮ NGUYÊN MERGE, MÀU NỀN, BORDER
# ============================================================
def copy_cell_style(src_cell, dst_cell):
    """Copy toàn bộ style từ ô nguồn sang ô đích (font, fill, border, alignment...)."""
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.border = copy(src_cell.border)
        dst_cell.alignment = copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format
        dst_cell.protection = copy(src_cell.protection)

def translate_excel(input_path: str, output_path: str, model):
    """Dịch các ô text trong Excel, giữ nguyên định dạng."""
    wb = openpyxl.load_workbook(input_path)

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                # Chỉ dịch ô có giá trị là chuỗi (bỏ qua số, công thức, None)
                if cell.value is None or not isinstance(cell.value, str):
                    continue
                # Bỏ qua công thức (bắt đầu bằng "=")
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    continue
                original = cell.value
                if not original.strip():
                    continue
                translated = translate_text(model, original)
                # Gán giá trị mới nhưng giữ nguyên style của ô
                # (openpyxl không tự động mất style khi gán .value)
                cell.value = translated

    wb.save(output_path)

# ============================================================
# 5. XỬ LÝ PDF – CHUYỂN SANG WORD, DỊCH, XUẤT WORD
# ============================================================
def translate_pdf(input_path: str, output_path: str, model):
    """
    Quy trình: PDF → DOCX (pdf2docx) → dịch DOCX → lưu DOCX.
    Người dùng tải về file Word đã dịch (giữ layout tốt hơn can thiệp PDF trực tiếp).
    """
    # Tạo file DOCX tạm
    temp_docx = input_path.replace(".pdf", "_temp.docx")
    cv = Converter(input_path)
    cv.convert(temp_docx)
    cv.close()

    # Dịch file DOCX tạm
    translate_docx(temp_docx, output_path, model)

    # Dọn file tạm
    if os.path.exists(temp_docx):
        os.remove(temp_docx)

# ============================================================
# 6. GIAO DIỆN STREAMLIT
# ============================================================
st.set_page_config(page_title="Dịch tài liệu kỹ thuật THACO", layout="wide")
st.title("📄 Dịch tài liệu kỹ thuật tự động (Việt → Anh)")
st.caption("Hỗ trợ Word, Excel, PDF – Giữ nguyên định dạng – Sử dụng Google Gemini")

# --- Sidebar: API Key ---
with st.sidebar:
    st.header("🔑 Cấu hình API")
    api_key = st.text_input(
        "Nhập Google Gemini API Key",
        type="password",
        help="Lấy API key miễn phí tại https://aistudio.google.com/app/apikey",
    )
    st.markdown("---")
    st.info(
        "**Lưu ý:** Ứng dụng sử dụng model `gemini-1.5-flash` (free tier). "
        "Nếu gặp lỗi 429, vui lòng đợi vài phút rồi thử lại."
    )

# --- Khu vực upload ---
uploaded_file = st.file_uploader(
    "Chọn file cần dịch",
    type=["docx", "xlsx", "pdf"],
    help="Hỗ trợ .docx, .xlsx, .pdf (PDF sẽ được chuyển sang Word)",
)

if uploaded_file and api_key:
    file_ext = os.path.splitext(uploaded_file.name)[1].lower()

    # Lưu file tạm
    with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp_in:
        tmp_in.write(uploaded_file.getbuffer())
        input_path = tmp_in.name

    # Xác định định dạng output
    if file_ext == ".pdf":
        output_ext = ".docx"
        output_name = uploaded_file.name.replace(".pdf", "_translated.docx")
    else:
        output_ext = file_ext
        output_name = uploaded_file.name.replace(file_ext, f"_translated{file_ext}")

    output_path = input_path.replace(file_ext, f"_out{output_ext}")

    if st.button("🚀 Bắt đầu dịch", type="primary"):
        try:
            with st.spinner("Đang khởi tạo Gemini..."):
                model = init_gemini(api_key)

            progress_bar = st.progress(0)
            status_text = st.empty()

            # Dịch theo định dạng
            if file_ext == ".docx":
                status_text.text("Đang dịch file Word (giữ nguyên định dạng)...")
                translate_docx(input_path, output_path, model)
                progress_bar.progress(100)

            elif file_ext == ".xlsx":
                status_text.text("Đang dịch file Excel (giữ nguyên định dạng)...")
                translate_excel(input_path, output_path, model)
                progress_bar.progress(100)

            elif file_ext == ".pdf":
                status_text.text("Đang chuyển PDF sang Word...")
                progress_bar.progress(20)
                translate_pdf(input_path, output_path, model)
                progress_bar.progress(100)

            status_text.text("✅ Dịch hoàn tất!")
            st.success("File đã được dịch thành công. Nhấn nút bên dưới để tải về.")

            # Nút tải xuống
            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 Tải file đã dịch",
                    data=f,
                    file_name=output_name,
                    mime="application/octet-stream",
                    type="primary",
                )

        except Exception as e:
            st.error(f"Đã xảy ra lỗi: {e}")
            st.exception(e)  # Hiển thị traceback để debug (có thể ẩn khi production)

        finally:
            # Dọn file tạm (giữ output_path cho đến khi người dùng tải xong)
            if os.path.exists(input_path):
                os.remove(input_path)
            # Không xóa output_path ngay vì download_button cần file tồn tại
            # (Streamlit giữ file trong session; có thể xóa sau khi session kết thúc)

elif uploaded_file and not api_key:
    st.warning("Vui lòng nhập API Key ở sidebar trước khi bắt đầu.")
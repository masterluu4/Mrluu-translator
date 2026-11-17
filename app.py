# app.py
# LUU PRO - HỆ THỐNG DỊCH THUẬT AI CHUYÊN NGÀNH Ô TÔ (V7.0 - DIGITAL COCKPIT UI)
# Giữ nguyên 100% Logic Dịch thuật & Đồng bộ Google Sheets + Giao diện Enterprise SaaS Hiện đại

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

st.set_page_config(
    page_title="LUU PRO | Automotive AI Translator",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"  # <-- 1. Tự động đóng gọn Sidebar khi mở trang
)

# ============================================================
# 2. HỆ THỐNG GIAO DIỆN KỸ THUẬT SỐ (FULL-SCREEN MODE)
# ============================================================
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap');

    /* 2. Ẩn thanh công cụ Share, Star, GitHub, Deploy và vạch màu trên cùng */
    .stAppToolbar,
    div[data-testid="stToolbarActions"],
    div[data-testid="stAppDeployButton"],
    div[data-testid="stDecoration"],
    #MainMenu,
    footer {
        display: none !important;
    }

    /* Làm trong suốt thanh Header trên cùng nhưng GIỮ LẠI nút bấm mở Sidebar ở góc trái */
    header[data-testid="stHeader"] {
        background: transparent !important;
        height: 2.5rem !important;
    }

    /* Làm nổi bật nút mũi tên mở Sidebar ở góc trên bên trái để dễ bấm khi cần */
    button[data-testid="stExpandSidebarButton"],
    div[data-testid="collapsedControl"] button {
        background-color: #1E3A8A !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        box-shadow: 0 4px 10px rgba(30, 58, 138, 0.3) !important;
    }
    button[data-testid="stExpandSidebarButton"] svg,
    div[data-testid="collapsedControl"] svg {
        fill: #FFFFFF !important;
        color: #FFFFFF !important;
    }

    /* Kéo giao diện chính sát lên đỉnh trang để bung Full màn hình */
    .stApp {
        background: linear-gradient(160deg, #F1F5F9 0%, #E2E8F0 100%);
        font-family: 'Inter', sans-serif;
    }
    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 1.5rem !important;
        max-width: 98% !important;
    }

    /* Thanh Sidebar phong cách Dark Industrial */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0F172A 0%, #1E293B 100%) !important;
        border-right: 1px solid #334155;
    }
    section[data-testid="stSidebar"] * {
        color: #F8FAFC !important;
    }
    section[data-testid="stSidebar"] .stTextInput input,
    section[data-testid="stSidebar"] .stTextArea textarea,
    section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
        background-color: #1E293B !important;
        color: #F8FAFC !important;
        border: 1px solid #475569 !important;
        border-radius: 8px !important;
    }
    section[data-testid="stSidebar"] details {
        background-color: rgba(30, 41, 59, 0.75) !important;
        border: 1px solid #334155 !important;
        border-radius: 10px !important;
        margin-bottom: 10px !important;
    }
    section[data-testid="stSidebar"] .stButton > button {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
        border: 1px solid #3B82F6 !important;
        color: #FFFFFF !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
    }

    /* HERO BANNER ĐẲNG CẤP TRÊN CÙNG */
    .hero-banner {
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 55%, #1D4ED8 100%);
        padding: 20px 28px;
        border-radius: 16px;
        color: #FFFFFF;
        box-shadow: 0 12px 28px -6px rgba(15, 23, 42, 0.28);
        margin-bottom: 20px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        border: 1px solid rgba(255, 255, 255, 0.14);
    }
    .hero-title {
        font-size: 25px;
        font-weight: 800;
        letter-spacing: -0.3px;
        margin: 0 0 6px 0;
        color: #FFFFFF;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .hero-sub {
        font-size: 13.5px;
        color: #CBD5E1;
        margin: 0;
        font-weight: 400;
    }
    .hero-badges {
        display: flex;
        gap: 10px;
        align-items: center;
        flex-wrap: wrap;
    }
    .cyber-badge {
        background: rgba(255, 255, 255, 0.12);
        backdrop-filter: blur(8px);
        border: 1px solid rgba(255, 255, 255, 0.25);
        color: #F8FAFC;
        padding: 6px 13px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 600;
    }
    .cyber-badge-green {
        background: rgba(16, 185, 129, 0.22);
        border: 1px solid #34D399;
        color: #A7F3D0;
        padding: 6px 13px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 700;
    }

    /* TIÊU ĐỀ PHÂN KHU TRONG CARD */
    .zone-header-blue {
        background: linear-gradient(90deg, #EFF6FF 0%, #DBEAFE 100%);
        border-left: 5px solid #2563EB;
        padding: 10px 14px;
        border-radius: 8px;
        margin-bottom: 12px;
        color: #1E3A8A;
        font-weight: 700;
        font-size: 15px;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .zone-header-teal {
        background: linear-gradient(90deg, #F0FDFA 0%, #CCFBF1 100%);
        border-left: 5px solid #0D9488;
        padding: 10px 14px;
        border-radius: 8px;
        margin-top: 8px;
        margin-bottom: 12px;
        color: #115E59;
        font-weight: 700;
        font-size: 15px;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .step-pill {
        background: #1E3A8A;
        color: #FFFFFF;
        font-size: 11px;
        padding: 2px 9px;
        border-radius: 99px;
        font-weight: 700;
    }
    .step-pill-teal {
        background: #0F766E;
        color: #FFFFFF;
        font-size: 11px;
        padding: 2px 9px;
        border-radius: 99px;
        font-weight: 700;
    }

    /* KHUNG CONTAINER TRẮNG NỔI BẬT */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #FFFFFF !important;
        border-radius: 14px !important;
        border: 1px solid #CBD5E1 !important;
        box-shadow: 0 6px 18px -4px rgba(15, 23, 42, 0.07) !important;
        padding: 6px !important;
    }

    /* LÀM NỔI BẬT KHUNG KÉO THẢ FILE (DROPZONE) */
    section[data-testid="stFileUploaderDropzone"] {
        background: linear-gradient(180deg, #F8FAFC 0%, #EFF6FF 100%) !important;
        border: 2px dashed #3B82F6 !important;
        border-radius: 12px !important;
        padding: 18px !important;
        transition: all 0.25s ease;
    }
    section[data-testid="stFileUploaderDropzone"]:hover {
        border-color: #1D4ED8 !important;
        background: #DBEAFE !important;
    }

    /* 3 THẺ THỐNGNG KÊ KPI KỸ THUẬT SỐ */
    .kpi-card {
        background: #FFFFFF;
        border-radius: 12px;
        padding: 12px 16px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 4px 10px rgba(15, 23, 42, 0.04);
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .kpi-blue { border-top: 4px solid #2563EB; }
    .kpi-purple { border-top: 4px solid #7C3AED; }
    .kpi-emerald { border-top: 4px solid #10B981; }
    .kpi-icon {
        font-size: 22px;
        background: #F1F5F9;
        width: 42px;
        height: 42px;
        display: flex;
        align-items: center;
        justify-content: center;
        border-radius: 10px;
    }
    .kpi-label {
        font-size: 11.5px;
        font-weight: 600;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.4px;
    }
    .kpi-value {
        font-size: 15.5px;
        font-weight: 800;
        color: #0F172A;
        margin-top: 2px;
    }

    /* MÀN HÌNH TERMINAL MAC-OS */
    .terminal-wrapper {
        border-radius: 12px;
        overflow: hidden;
        border: 1px solid #1E293B;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.35);
        margin-bottom: 8px;
    }
    .terminal-bar {
        background: #1E293B;
        padding: 8px 14px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        border-bottom: 1px solid #334155;
    }
    .term-dots {
        display: flex;
        gap: 7px;
    }
    .dot-r { width: 11px; height: 11px; border-radius: 50%; background: #EF4444; display: inline-block; }
    .dot-y { width: 11px; height: 11px; border-radius: 50%; background: #F59E0B; display: inline-block; }
    .dot-g { width: 11px; height: 11px; border-radius: 50%; background: #10B981; display: inline-block; }
    .term-title {
        color: #94A3B8;
        font-family: 'JetBrains Mono', monospace;
        font-size: 11.5px;
        font-weight: 600;
    }
    .log-container {
        background-color: #090D16;
        color: #34D399;
        font-family: 'JetBrains Mono', monospace;
        padding: 14px 16px;
        height: 290px;
        overflow-y: auto;
        font-size: 12.5px;
        line-height: 1.65;
    }

    /* NÚT BẤM CHÍNH (PRIMARY BUTTON) NỔI BẬT */
    div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 50%, #1E40AF 100%) !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 0.65rem 1.2rem !important;
        font-weight: 700 !important;
        font-size: 14.5px !important;
        letter-spacing: 0.2px !important;
        box-shadow: 0 6px 16px rgba(37, 99, 235, 0.35) !important;
        transition: all 0.2s ease !important;
    }
    div.stButton > button[kind="primary"]:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 8px 22px rgba(29, 78, 216, 0.48) !important;
    }
    div.stDownloadButton > button {
        background: linear-gradient(135deg, #059669 0%, #047857 100%) !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 0.7rem 1.2rem !important;
        font-weight: 700 !important;
        box-shadow: 0 6px 16px rgba(5, 150, 105, 0.35) !important;
    }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# 3. TỪ ĐIỂN VIẾT TẮT & THUẬT NGỮ CHUẨN Ô TÔ (MẶC ĐỊNH)
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
# 4. ĐỒNG BỘ HAI CHIỀU VỚI GOOGLE SHEETS (CLOUD TM & GLOSSARY)
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
    if not webapp_url or not webapp_url.strip().startswith("https://script.google.com"):
        return False, "Đường dẫn Google Apps Script không hợp lệ."
    try:
        resp = requests.get(webapp_url.strip(), timeout=25)
        if resp.status_code != 200:
            return False, f"Lỗi HTTP {resp.status_code} khi kết nối Google Sheets."
            
        data = resp.json()
        cloud_glossary = data.get("glossary", {})
        cloud_tm = data.get("tm", {})

        default_g_dict = glossary_str_to_dict(DEFAULT_THACO_GLOSSARY)
        if cloud_glossary:
            merged_g = {**default_g_dict, **cloud_glossary}
            st.session_state.glossary_text = dict_to_glossary_str(merged_g)
        else:
            st.session_state.glossary_text = DEFAULT_THACO_GLOSSARY
            push_to_gsheet(webapp_url, "save_glossary", default_g_dict)

        with open(GLOSSARY_FILE, "w", encoding="utf-8") as f:
            f.write(st.session_state.glossary_text)

        if cloud_tm and isinstance(cloud_tm, dict):
            st.session_state.tm_data.update(cloud_tm)
            with open(TM_FILE, "w", encoding="utf-8") as f:
                json.dump(st.session_state.tm_data, f, ensure_ascii=False, indent=2)

        g_count = len(glossary_str_to_dict(st.session_state.glossary_text))
        tm_count = len(st.session_state.tm_data)
        return True, f"Đã đồng bộ Cloud: {g_count} thuật ngữ & {tm_count} câu TM"
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

# Khởi tạo Session State
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

def render_terminal_html(logs_list: list) -> str:
    body = "<br>".join(logs_list[-50:]) if logs_list else (
        "✅ <b>LUU PRO V7.0 (Digital Cockpit Edition)</b> sẵn sàng hoạt động!<br>"
        "• <b>Bước 1</b>: Kéo thả tài liệu (.xlsx, .docx, .pptx, .pdf) vào Phân khu 01 bên trái.<br>"
        "• <b>Bước 2</b>: Chọn chế độ ở Phân khu 02 và nhấn nút <b>Bắt đầu</b> để AI xử lý."
    )
    return f"""
    <div class="terminal-wrapper">
        <div class="terminal-bar">
            <div class="term-dots">
                <span class="dot-r"></span><span class="dot-y"></span><span class="dot-g"></span>
            </div>
            <div class="term-title">⚡ AI_TRANSLATION_ENGINE // LIVE_CONSOLE</div>
            <div style="color:#10B981; font-size:11px; font-family:monospace;">● ONLINE</div>
        </div>
        <div class="log-container">{body}</div>
    </div>
    """

def add_log(message: str, log_placeholder=None):
    time_str = time.strftime("%H:%M:%S")
    st.session_state.logs.append(f"<span style='color:#64748B;'>[{time_str}]</span> {message}")
    if log_placeholder is not None:
        log_placeholder.markdown(render_terminal_html(st.session_state.logs), unsafe_allow_html=True)

# ============================================================
# 5. KẾT NỐI AI & GỌI JSON BATCH THÔNG MINH (CHỐNG LỖI 429)
# ============================================================
def build_system_prompt(direction: str, glossary_str: str) -> str:
    dir_instruction = (
        "Translate every value in the input JSON object from Vietnamese into professional, grammatically complete automotive technical English."
        if direction == "Việt → Anh"
        else "Translate every value in the input JSON object from English into professional, accurate automotive technical Vietnamese."
    )
    return f"""You are a senior automotive engineering and after-sales quality management translator.
{dir_instruction}

MANDATORY GLOSSARY & ABBREVIATIONS (Always follow these exact translations when encountered):
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
# 6. BỘ LỌC NHẬN DIỆN VĂN BẢN & DỊCH KẾT HỢP CLOUD TM + GLOSSARY
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
    tm = st.session_state.tm_data
    glossary_dict = parse_glossary(glossary_str, direction)
    
    translated_map = {}
    need_ai = []
    tm_hits = 0

    for s in unique_strings:
        s_clean = s.strip()
        matched_glossary = None
        for k, v in glossary_dict.items():
            if s_clean.lower() == k.lower():
                matched_glossary = v
                break
        if matched_glossary is not None:
            translated_map[s] = matched_glossary
            tm_hits += 1
            continue

        tm_key = f"{direction}::{s_clean}"
        if tm_key in tm and str(tm[tm_key]).strip():
            translated_map[s] = str(tm[tm_key]).strip()
            tm_hits += 1
        elif direction == "Việt → Anh" and s_clean in tm and str(tm[s_clean]).strip():
            translated_map[s] = str(tm[s_clean]).strip()
            tm_hits += 1
        else:
            need_ai.append(s)

    add_log(f"⚡ Đã lấy tức thì {tm_hits}/{len(unique_strings)} câu từ Từ điển & Google Sheets TM (0.01s). Cần gọi AI: {len(need_ai)} câu.", log_placeholder)

    if not need_ai:
        progress_bar.progress(1.0)
        return translated_map

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

        save_tm(tm, new_items_only=batch_new_tm, webapp_url=gsheet_url)
        progress_bar.progress(min((start + len(batch_slice)) / total_ai, 1.0))
        if start + BATCH_SIZE < total_ai:
            time.sleep(1.5)

    if newly_translated_for_cloud:
        add_log(f"☁️ Đã tự động lưu {len(newly_translated_for_cloud)} câu dịch mới lên Google Sheets (Tab TM).", log_placeholder)

    return translated_map

# ============================================================
# 7. THUẬT TOÁN GOM CÂU THÔNG MINH (CHỐNG CẮT VỤN RUN POWERPOINT/WORD)
# ============================================================
class ParagraphSegment:
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
# 8. TRÍCH XUẤT & KẾT XUẤT CHO EXCEL, WORD, POWERPOINT, PDF
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
# 9. THANH BÊN (SIDEBAR): CẤU HÌNH HỆ THỐNG, TỪ ĐIỂN & CLOUD TM
# ============================================================
with st.sidebar:
    st.markdown("### 🎛️ Trung tâm Điều khiển")
    st.caption("Cấu hình lõi AI, Từ điển chuyên ngành & Đồng bộ Đám mây")
    
    model_list = st.session_state.get("available_models", [
        "qwen/qwen3.8-27b",
        "meta-llama/llama-4-scout-17b-16e-instruct",
        "meta-llama/llama-4-maverick-17b-128e-instruct",
        "deepseek-r1-distill-llama-70b",
        "gemma2-9b-it"
    ])
    selected_model = st.selectbox("🤖 Động cơ Model AI:", options=model_list)

    with st.expander("🔌 Kết nối Server AI & Cloud", expanded=False):
        proxy_url = st.text_input("🔗 Cloudflare Worker:", value=DEFAULT_PROXY_URL)
        gsheet_url = st.text_input("☁️ Google Sheets Web App:", value=DEFAULT_GSHEET_WEBAPP_URL)
        direct_key = st.text_input("🔑 Groq API Key (Tùy chọn):", type="password")
        
        sb_c1, sb_c2 = st.columns(2)
        with sb_c1:
            if st.button("🔄 Tải Model", use_container_width=True):
                try:
                    models = fetch_live_chat_models(proxy_url, direct_key)
                    st.session_state.available_models = models
                    st.success(f"Đã tải {len(models)} model!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Lỗi: {e}")
        with sb_c2:
            if st.button("☁️ Đồng bộ", use_container_width=True):
                ok, msg = sync_from_gsheet(gsheet_url)
                st.session_state.cloud_sync_status = (ok, msg)
                if ok: st.success(msg)
                else: st.warning(msg)

    g_count_now = len(glossary_str_to_dict(st.session_state.glossary_text))
    with st.expander(f"📚 Từ điển Chuyên ngành ({g_count_now} mục)", expanded=False):
        st.caption("Cú pháp: `Tiếng Việt = Tiếng Anh chuẩn`")
        new_glossary = st.text_area(
            "Danh sách thuật ngữ:",
            value=st.session_state.glossary_text,
            height=220,
            label_visibility="collapsed"
        )
        if st.button("💾 Lưu lên Google Sheets", use_container_width=True):
            st.session_state.glossary_text = new_glossary
            if save_glossary_text(new_glossary, gsheet_url):
                st.success("✅ Đã lưu lên Google Sheets!")
            else:
                st.info("💾 Đã lưu cục bộ.")

    with st.expander(f"🧠 Bộ nhớ Cloud TM ({len(st.session_state.tm_data)} câu)", expanded=False):
        tm_df_export = pd.DataFrame(
            [{"Khoa_Goc": k, "Ban_Dich_Chuan": v} for k, v in st.session_state.tm_data.items()]
        )
        st.download_button(
            "⬇️ Xuất file TM (.csv)",
            data=tm_df_export.to_csv(index=False).encode("utf-8-sig"),
            file_name="Translation_Memory_Backup.csv",
            mime="text/csv",
            use_container_width=True
        )
        if st.button("🔄 Làm mới từ Google Sheets", use_container_width=True):
            ok, msg = sync_from_gsheet(gsheet_url)
            if ok:
                st.success(msg)
                st.rerun()

    st.divider()
    if st.session_state.get("cloud_sync_status"):
        sync_ok, sync_msg = st.session_state.cloud_sync_status
        st.caption(f"{'🟢' if sync_ok else '🟠'} {sync_msg}")

# ============================================================
# 10. MÀN HÌNH CHÍNH: HERO BANNER + WORKSPACE ĐẲNG CẤP
# ============================================================
st.markdown("""
    <div class="hero-banner">
        <div>
            <div class="hero-title">⚡ HỆ THỐNG DỊCH THUẬT AI CHUYÊN NGÀNH Ô TÔ — LUU PRO</div>
            <div class="hero-sub">Nền tảng Dịch thuật Kỹ thuật Số • Bảo toàn 100% Cấu trúc Excel, Word, PowerPoint, PDF • Đồng bộ Đám mây Thời gian thực</div>
        </div>
        <div class="hero-badges">
            <span class="cyber-badge">🚀 ENGINE V7.0</span>
            <span class="cyber-badge-green">☁️ CLOUD TM ACTIVE</span>
        </div>
    </div>
""", unsafe_allow_html=True)

col_work, col_monitor = st.columns([1.15, 1.85], gap="large")

# --- CỘT TRÁI: TRẠM TIẾP NHẬN FILE & CẤU HÌNH DỊCH NHANH ---
with col_work:
    with st.container(border=True):
        st.markdown(
            '<div class="zone-header-blue">'
            '<span>📂 1. TIẾP NHẬN TÀI LIỆU KỸ THUẬT</span>'
            '<span class="step-pill">BƯỚC 1</span>'
            '</div>',
            unsafe_allow_html=True
        )
        uploaded_file = st.file_uploader(
            "Kéo thả hoặc chọn file Excel (.xlsx), Word (.docx), PowerPoint (.pptx), PDF (.pdf):",
            type=["xlsx", "docx", "pptx", "pdf"]
        )

        st.markdown(
            '<div class="zone-header-teal">'
            '<span>⚙️ 2. THIẾT LẬP CHẾ ĐỘ XỬ LÝ AI</span>'
            '<span class="step-pill-teal">BƯỚC 2</span>'
            '</div>',
            unsafe_allow_html=True
        )
        direction = st.radio("🌐 Chiều dịch thuật:", ["Việt → Anh", "Anh → Việt"], horizontal=True)
        
        opt_c1, opt_c2 = st.columns(2)
        with opt_c1:
            bilingual = st.checkbox("📖 Xuất Song Ngữ\n(Gốc + Dịch dòng dưới)", value=False)
            translate_sheets = st.checkbox("📑 Dịch tên Sheet Excel", value=True)
        with opt_c2:
            auto_fit = st.checkbox("📐 Auto-Fit Layout\n(Tự xuống dòng & giãn cột)", value=True)
            enable_review = st.checkbox("✍️ Duyệt & Hiệu đính\ntrước khi xuất File", value=True)

        st.write("")
        btn_label = "🚀 BẮT ĐẦU PHÂN TÍCH & DỊCH NHÁP (ĐỂ HIỆU ĐÍNH)" if enable_review else "🚀 BẮT ĐẦU DỊCH & XUẤT FILE NGAY"
        start_clicked = st.button(
            btn_label,
            type="primary",
            use_container_width=True,
            disabled=(uploaded_file is None)
        )
        if uploaded_file is None:
            st.caption("💡 *Hãy thả tài liệu vào Phân khu 01 phía trên để mở khóa nút kích hoạt AI.*")

    if st.session_state.get("final_output_bytes") is not None:
        with st.container(border=True):
            st.success(f"🎉 Xử lý hoàn tất! Tài liệu **{st.session_state.final_output_name}** đã sẵn sàng:")
            st.download_button(
                label=f"⬇️ TẢI XUỐNG TÀI LIỆU HOÀN THIỆN ({st.session_state.final_output_name})",
                data=st.session_state.final_output_bytes,
                file_name=st.session_state.final_output_name,
                mime="application/octet-stream",
                type="primary",
                use_container_width=True
            )

# --- CỘT PHẢI: 3 THẺ KPI KỸ THUẬT SỐ + TERMINAL GIÁM SÁT & BẢNG HIỆU ĐÍNH ---
with col_monitor:
    kpi1, kpi2, kpi3 = st.columns(3)
    with kpi1:
        st.markdown(f"""
            <div class="kpi-card kpi-blue">
                <div class="kpi-icon">🤖</div>
                <div>
                    <div class="kpi-label">Model AI Đang Chạy</div>
                    <div class="kpi-value">{selected_model.split('/')[-1]}</div>
                </div>
            </div>
        """, unsafe_allow_html=True)
    with kpi2:
        st.markdown(f"""
            <div class="kpi-card kpi-purple">
                <div class="kpi-icon">📚</div>
                <div>
                    <div class="kpi-label">Từ Điển Chuyên Ngành</div>
                    <div class="kpi-value">{g_count_now} Thuật ngữ</div>
                </div>
            </div>
        """, unsafe_allow_html=True)
    with kpi3:
        st.markdown(f"""
            <div class="kpi-card kpi-emerald">
                <div class="kpi-icon">🧠</div>
                <div>
                    <div class="kpi-label">Bộ Nhớ Đệm Cloud TM</div>
                    <div class="kpi-value">{len(st.session_state.tm_data)} Câu chuẩn</div>
                </div>
            </div>
        """, unsafe_allow_html=True)

    st.write("")
    with st.container(border=True):
        tab_console, tab_review = st.tabs([
            "🖥️ TRẠM GIÁM SÁT TIẾN TRÌNH (LIVE CONSOLE)",
            f"✍️ BẢNG HIỆU ĐÍNH TƯƠNG TÁC ({len(st.session_state.review_df) if st.session_state.review_df is not None else 0} CÂU)"
        ])

        with tab_console:
            log_placeholder = st.empty()
            log_placeholder.markdown(render_terminal_html(st.session_state.logs), unsafe_allow_html=True)
            progress_bar = st.progress(0.0)

        # Xử lý sự kiện khi bấm nút BẮT ĐẦU DỊCH
        if uploaded_file is not None and start_clicked:
            st.session_state.logs = []
            st.session_state.final_output_bytes = None
            st.session_state.glossary_text = new_glossary
            st.session_state.staged_file_bytes = uploaded_file.getvalue()
            st.session_state.staged_file_name = uploaded_file.name
            
            file_ext = os.path.splitext(uploaded_file.name)[1].lower()
            add_log(f"🚀 Khởi chạy dịch tài liệu <b>{uploaded_file.name}</b> | Model: <b>{selected_model}</b> | Chiều: <b>{direction}</b>", log_placeholder)

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
                add_log(f"🎉 Hoàn tất dịch {len(unique_list)} câu duy nhất chỉ trong <b>{elapsed} giây</b>!", log_placeholder)

                st.session_state.review_df = pd.DataFrame([
                    {"Câu gốc trong File": k, "Bản dịch AI (Bấm đúp để sửa tay)": v}
                    for k, v in trans_map.items()
                ])

                if enable_review:
                    add_log("👉 Mời anh bấm sang tab <b>'✍️ BẢNG HIỆU ĐÍNH TƯƠNG TÁC'</b> bên cạnh để duyệt và xuất file!", log_placeholder)
                else:
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
                    st.rerun()

            except Exception as e:
                add_log(f"❌ LỖI HỆ THỐNG: {str(e)}", log_placeholder)
            finally:
                if os.path.exists(input_path):
                    os.remove(input_path)

        with tab_review:
            if st.session_state.review_df is not None:
                st.caption("Bấm đúp vào cột **'Bản dịch AI'** để chỉnh sửa trực tiếp. Hệ thống sẽ tự động cập nhật câu sửa lên Google Sheets Cloud TM:")
                edited_df = st.data_editor(
                    st.session_state.review_df,
                    use_container_width=True,
                    height=290,
                    disabled=["Câu gốc trong File"],
                    key="editor_grid"
                )

                if st.button("✅ ÁP DỤNG BẢN HIỆU ĐÍNH, ĐỒNG BỘ CLOUD & KẾT XUẤT FILE", type="primary", use_container_width=True):
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

                        save_tm(tm, new_items_only=changed_for_cloud if changed_for_cloud else None, webapp_url=gsheet_url)
                        if changed_for_cloud:
                            add_log(f"☁️ Đã cập nhật {len(changed_for_cloud)} câu hiệu đính lên Google Sheets.", log_placeholder)

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
                            add_log(f"✅ Đã kết xuất thành công file <b>{output_name}</b>!", log_placeholder)
                            st.rerun()
                        finally:
                            if os.path.exists(in_p): os.remove(in_p)
                            if os.path.exists(out_p): os.remove(out_p)
            else:
                st.info("📋 Bảng hiệu đính tương tác sẽ hiển thị tại đây ngay sau khi anh tải file lên và nhấn nút **Bắt đầu Phân tích & Dịch nháp**.")
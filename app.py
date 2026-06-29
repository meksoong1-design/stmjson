# ============================================================
# STM Document AI Bank Statement Parser v2 - JSON Structure Edition
# วิธีใหม่: อ่าน Table Structure + Bounding Box จาก Document AI JSON
# ไม่ต้องใช้ keyword list / BANK_CONFIGS / เดาจากข้อความ
# Output sheets: เจ้าของบัญชี, รายละเอียด เดบิต-เครดิต, สรุปยอด,
#                BANK STATEMENT 1, check, ocr_review, raw_text, document_json
# ============================================================

import os
import re
import io
import json
import math
import tempfile
from datetime import datetime

import pandas as pd
import streamlit as st

from google.api_core.client_options import ClientOptions
from google.cloud import documentai_v1 as documentai
from google.oauth2 import service_account
from google.protobuf.json_format import MessageToDict

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill


# ============================================================
# 0) PAGE CONFIG
# ============================================================
st.set_page_config(page_title="STM Document AI Parser v2", layout="centered")

st.markdown("""
<style>
.block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 980px; }
h1 { font-size: 2.35rem !important; font-weight: 800 !important; }
.section-title { display:flex; align-items:center; gap:8px; font-weight:800; font-size:1.15rem; margin:18px 0 10px 0; }
.section-title .step-number { display:inline-flex; align-items:center; justify-content:center; min-width:30px; height:30px;
  border-radius:999px; background:#4c82fb; color:white; font-weight:800; }
div[data-testid="stButton"] button, div[data-testid="stDownloadButton"] button { border-radius:10px; font-weight:750; }
div[data-testid="stFileUploader"] section { border-radius:10px; }
</style>
""", unsafe_allow_html=True)


# ============================================================
# 1) CONSTANTS
# ============================================================
BALANCE_TOLERANCE = 0.05
FORCE_YEAR = datetime.now().year

BANK_LABELS = {
    "AUTO":     "ตรวจจับอัตโนมัติ",
    "KBANK":    "กสิกรไทย (KBANK)",
    "BBL":      "กรุงเทพ (BBL)",
    "SCB":      "ไทยพาณิชย์ (SCB)",
    "KRUNGSRI": "กรุงศรี (Krungsri)",
    "DEFAULT":  "ไม่ระบุ / อื่นๆ",
}

# ── column-index hints per bank (0-based, can be overridden by auto-detect)
BANK_COL_HINTS = {
    # bank: {date, desc, debit, credit, balance}
    "KBANK":    {"date": 0, "desc": 1, "debit": 2, "credit": 3, "balance": 4},
    "BBL":      {"date": 0, "desc": 1, "debit": 2, "credit": 3, "balance": 4},
    "SCB":      {"date": 0, "desc": 1, "debit": 2, "credit": 3, "balance": 4},
    "KRUNGSRI": {"date": 0, "desc": 1, "debit": 2, "credit": 3, "balance": 4},
    "DEFAULT":  {"date": 0, "desc": 1, "debit": 2, "credit": 3, "balance": 4},
}

OPENING_KEYWORDS = [
    "ยอดยกมา", "ยอด ยก มา", "ยอด ยก", "ยก มา",
    "BALANCE BROUGHT FORWARD", "BROUGHT FORWARD", "BALANCE B/F", "B/F",
    "ยอดเงินคงเหลือยกมา",
]


# ============================================================
# 2) LOGIN
# ============================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "results" not in st.session_state:
    st.session_state.results = None


def login_page():
    st.title("Login")
    st.caption("กรุณาเข้าสู่ระบบก่อนใช้งาน STM Document AI Parser")
    password = st.text_input("Password", type="password")
    if st.button("เข้าสู่ระบบ", type="primary", use_container_width=True):
        correct_pw = st.secrets.get("APP_PASSWORD", "stm2025")
        if password == correct_pw:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("รหัสผ่านไม่ถูกต้อง")
    st.stop()


# ============================================================
# 3) DOCUMENT AI  (คืน document object + dict)
# ============================================================
def get_documentai_config():
    try:
        return (
            str(st.secrets["DOCUMENTAI_PROJECT_ID"]),
            str(st.secrets["DOCUMENTAI_LOCATION"]),
            str(st.secrets["DOCUMENTAI_PROCESSOR_ID"]),
        )
    except Exception as e:
        st.error(f"ไม่พบค่า DOCUMENTAI config ใน Secrets: {e}")
        st.stop()


def get_documentai_client():
    if "gcp_service_account" not in st.secrets:
        st.error("ไม่พบ [gcp_service_account] ใน Streamlit Secrets")
        st.stop()
    try:
        info = dict(st.secrets["gcp_service_account"])
        if "private_key" in info and isinstance(info["private_key"], str):
            info["private_key"] = info["private_key"].replace("\\n", "\n")
        creds = service_account.Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/cloud-platform"]
        )
    except Exception as e:
        st.error(f"อ่าน [gcp_service_account] ไม่สำเร็จ: {e}")
        st.stop()

    project_id, location, processor_id = get_documentai_config()
    opts = ClientOptions(api_endpoint=f"{location}-documentai.googleapis.com")
    client = documentai.DocumentProcessorServiceClient(credentials=creds, client_options=opts)
    processor_name = client.processor_path(project_id, location, processor_id)
    return client, processor_name


def get_mime_type(filename):
    name = filename.lower()
    if name.endswith(".pdf"):   return "application/pdf"
    if name.endswith((".jpg", ".jpeg")): return "image/jpeg"
    if name.endswith(".png"):   return "image/png"
    raise ValueError("รองรับเฉพาะ PDF, JPG, JPEG, PNG")


def call_document_ai(uploaded_file):
    """ส่งไฟล์ไปยัง Document AI แล้วคืน (document_proto, doc_dict, raw_text)"""
    client, processor_name = get_documentai_client()
    content = uploaded_file.read()
    uploaded_file.seek(0)
    raw_doc = documentai.RawDocument(content=content, mime_type=get_mime_type(uploaded_file.name))
    req     = documentai.ProcessRequest(name=processor_name, raw_document=raw_doc)
    result  = client.process_document(request=req)
    document = result.document
    doc_dict = MessageToDict(document._pb, preserving_proto_field_name=True)
    raw_text = document.text or ""
    return document, doc_dict, raw_text


# ============================================================
# 4) TEXT / MONEY / DATE HELPERS
# ============================================================
def clean_text(value):
    if value is None: return ""
    return re.sub(r"\s+", " ", str(value).replace("\x00", " ").strip())


def estimate_text_width(value):
    text = clean_text(value)
    w = 0.0
    for ch in text:
        if "\u0e00" <= ch <= "\u0e7f": w += 1.4
        elif ch.isupper():              w += 1.1
        elif ch.isdigit() or ch in ".,:-/()": w += 0.85
        elif ch == " ":                 w += 0.4
        else:                           w += 0.95
    return w


def fix_ocr_chars(s):
    s = str(s)
    for bad, good in [("O","0"),("o","0"),("I","1"),("l","1"),("|","1"),
                      ("S","5"),("s","5"),("B","8"),("g","9"),("q","9"),
                      ("Z","2"),("z","2")]:
        s = s.replace(bad, good)
    return s


def parse_money(s):
    """แปลง string → float | None  (รองรับ 1,234.56 / 1.234,56 / 1234.56)"""
    if s is None: return None
    s = fix_ocr_chars(str(s)).strip()
    s = re.sub(r"\s+", "", s)
    # European format: 1.234,56
    if re.match(r"^\d{1,3}(\.\d{3})+,\d{2}$", s):
        s = s.replace(".", "").replace(",", ".")
    # comma as decimal separator only: 1234,56
    elif re.match(r"^\d+,\d{2}$", s):
        s = s.replace(",", ".")
    else:
        s = s.replace(",", "")  # remove thousand-separator commas
    s = re.sub(r"[^0-9.\-]", "", s)
    if not s: return None
    try:   return float(s)
    except: return None


def extract_money_values(text):
    """ดึง float ทุกตัวที่เป็น amount/balance จากข้อความ"""
    text = fix_ocr_chars(str(text))
    patterns = [
        r"-?\d{1,3}(?:,\d{3})+\.\d{2}",
        r"-?\d{1,3}(?:\.\d{3})+,\d{2}",
        r"-?\d+\.\d{2}",
        r"-?\d+,\d{2}",
    ]
    seen, vals = set(), []
    for pat in patterns:
        for m in re.findall(pat, text):
            v = parse_money(m)
            if v is not None and v not in seen:
                seen.add(v); vals.append(v)
    return vals


def fix_ocr_date(s):
    s = str(s).strip().replace("O","0").replace("o","0").replace("I","1").replace("l","1")
    return re.sub(r"[./]", "-", s)


def normalize_date(s):
    s = fix_ocr_date(s)
    m1 = re.match(r"^(\d{1,2})-(\d{1,2})-(\d{2,4})$", s)
    m2 = re.match(r"^(\d{1,2})-(\d{1,2})$", s)
    if m1:
        day, month = int(m1.group(1)), int(m1.group(2))
        y = m1.group(3)
        year = 2000 + int(y) if len(y) == 2 else int(y)
    elif m2:
        day, month, year = int(m2.group(1)), int(m2.group(2)), FORCE_YEAR
    else:
        return None
    try:    return datetime(year, month, day).strftime("%d/%m/%Y")
    except: return None


def extract_date(text):
    for tok in str(text).split():
        d = normalize_date(tok)
        if d: return d
    return None


def extract_time(text):
    m = re.search(r"\b\d{1,2}:\d{2}(?::\d{2})?\b", str(text))
    return m.group(0) if m else ""


def is_opening_balance(text):
    up = clean_text(text).upper()
    return any(kw.upper() in up for kw in OPENING_KEYWORDS)


def is_valid_number(x):
    try:    return x is not None and not math.isnan(float(x))
    except: return False


def remove_money_and_date(text):
    """ลบ amount / date / time ออกจาก description"""
    t = str(text)
    for pat in [r"-?\d{1,3}(?:,\d{3})+\.\d{2}", r"-?\d+\.\d{2}", r"-?\d+,\d{2}",
                r"\b\d{1,2}:\d{2}(?::\d{2})?\b",
                r"\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b",
                r"\b\d{1,2}[-/]\d{1,2}\b"]:
        t = re.sub(pat, " ", t)
    return clean_text(t)


# ============================================================
# 5) BANK DETECTION  (จาก raw_text เท่านั้น, ไม่ใช้ keyword สำหรับ classify)
# ============================================================
BANK_DETECT = {
    "KBANK":    ["KASIKORN", "KBANK", "K PLUS", "กสิกร"],
    "BBL":      ["BANGKOK BANK", "BANGKOKBANK", "ธนาคารกรุงเทพ", "BBL", "BUALUANG"],
    "SCB":      ["SCB", "ไทยพาณิชย์", "SIAM COMMERCIAL"],
    "KRUNGSRI": ["KRUNGSRI", "กรุงศรี", "AYUDHYA", "BANK OF AYUDHYA"],
}

def detect_bank(raw_text):
    text = clean_text(raw_text).upper()
    best, best_score = "DEFAULT", 0
    for bank, kws in BANK_DETECT.items():
        score = sum(10 for kw in kws if kw.upper() in text)
        if score > best_score:
            best_score, best = score, bank
    return best


# ============================================================
# 6) ACCOUNT INFO
# ============================================================
def extract_account_info(raw_text):
    owner, acc_no = "", ""
    lines = [clean_text(x) for x in raw_text.splitlines() if clean_text(x)]
    for i, line in enumerate(lines[:60]):
        if not owner:
            if re.search(r"(ชื่อบัญชี|Account Name|ชื่อ\s*-\s*นามสกุล|ชื่อ\s*สกุล)", line, re.I):
                owner = re.sub(r"^.*?(ชื่อบัญชี|Account\s*Name|ชื่อ\s*-\s*นามสกุล|ชื่อ\s*สกุล)\s*[:：]?\s*", "", line, flags=re.I).strip()
                if not owner and i+1 < len(lines):
                    owner = lines[i+1]
            elif re.match(r"^(นาย|นาง|น\.ส\.|นางสาว|MR\.|MRS\.|MS\.)\s+", line, re.I):
                owner = line
        if not acc_no:
            m = re.search(r"(\d{3}[- ]?\d{1}[- ]?\d{5}[- ]?\d{1}|\d{10,12})", line)
            if m: acc_no = m.group(1).replace(" ", "")
    return owner or "ไม่ระบุ", acc_no or "ไม่ระบุ"


# ============================================================
# 7) CORE: อ่าน TABLE จาก Document AI JSON
# ============================================================

def get_cell_text(cell, full_text):
    """ดึง text จาก cell โดยใช้ text_anchor ใน Document AI JSON"""
    segments = (
        cell.get("layout", {}).get("text_anchor", {}).get("text_segments", [])
        or cell.get("layout", {}).get("textAnchor", {}).get("textSegments", [])
    )
    if not segments or not full_text:
        return ""
    parts = []
    for seg in segments:
        start = int(seg.get("start_index", seg.get("startIndex", 0)) or 0)
        end   = int(seg.get("end_index",   seg.get("endIndex",   0)) or 0)
        parts.append(full_text[start:end])
    return clean_text("".join(parts))


def bounding_box_x(layout):
    """คืน normalized x-center ของ bounding box (0–1)"""
    verts = (
        layout.get("bounding_poly", {}).get("normalized_vertices", [])
        or layout.get("boundingPoly", {}).get("normalizedVertices", [])
    )
    if not verts: return None
    xs = [v.get("x", 0) for v in verts]
    return sum(xs) / len(xs) if xs else None


def auto_detect_col_map(header_row, full_text):
    """
    วิเคราะห์ header row เพื่อ map column index → role
    Returns dict: {role: col_index}  roles: date, desc, debit, credit, balance
    """
    DATE_KW    = ["วันที่", "DATE", "วนท", "วัน"]
    DESC_KW    = ["รายการ", "DESCRIPTION", "DESC", "DETAIL", "รายละเอียด", "TRANSACTION"]
    DEBIT_KW   = ["เดบิต", "DEBIT", "ถอน", "จ่าย", "引出", "WITHDRAWAL", "DR"]
    CREDIT_KW  = ["เครดิต", "CREDIT", "ฝาก", "รับ", "DEPOSIT", "CR"]
    BALANCE_KW = ["ยอดคงเหลือ", "BALANCE", "คงเหลือ", "BAL", "ยอด"]

    col_map = {}
    for idx, cell in enumerate(header_row.get("cells", [])):
        text = get_cell_text(cell, full_text).upper()
        if any(k.upper() in text for k in BALANCE_KW):
            col_map.setdefault("balance", idx)
        elif any(k.upper() in text for k in DEBIT_KW):
            col_map.setdefault("debit", idx)
        elif any(k.upper() in text for k in CREDIT_KW):
            col_map.setdefault("credit", idx)
        elif any(k.upper() in text for k in DATE_KW):
            col_map.setdefault("date", idx)
        elif any(k.upper() in text for k in DESC_KW):
            col_map.setdefault("desc", idx)
    return col_map


def fallback_col_map_from_xpos(header_row, full_text, n_cols):
    """
    เมื่อ header ไม่มีข้อความ ใช้ตำแหน่ง x ของ column
    สมมติ: col ซ้ายสุด = date, ขวาสุด = balance, กลางๆ = debit/credit
    """
    if n_cols < 3: return {"date": 0, "desc": 1, "balance": n_cols - 1}
    if n_cols == 3: return {"date": 0, "desc": 1, "balance": 2}
    if n_cols == 4: return {"date": 0, "desc": 1, "debit": 2, "balance": 3}
    # 5+ cols: date, desc, debit, credit, balance
    return {"date": 0, "desc": 1, "debit": n_cols - 3, "credit": n_cols - 2, "balance": n_cols - 1}


def parse_table(table, full_text, bank_hint="DEFAULT"):
    """
    อ่านตาราง Document AI → list of row dicts
    คืน: list[{date, time, desc, debit, credit, balance, raw, is_opening, check_flag}]
    """
    header_rows = table.get("header_rows", table.get("headerRows", []))
    body_rows   = table.get("body_rows",   table.get("bodyRows",   []))

    # detect column map จาก header
    col_map = {}
    if header_rows:
        col_map = auto_detect_col_map(header_rows[0], full_text)
    if len(col_map) < 3:
        # ลองหา header ใน body rows แรก
        n_cols = max((len(r.get("cells", [])) for r in body_rows[:3]), default=5)
        col_map = fallback_col_map_from_xpos(header_rows[0] if header_rows else {}, full_text, n_cols)

    rows = []
    for body_row in body_rows:
        cells = body_row.get("cells", [])
        def cell_text(role):
            idx = col_map.get(role)
            if idx is None or idx >= len(cells): return ""
            return get_cell_text(cells[idx], full_text)

        date_raw  = cell_text("date")
        desc_raw  = cell_text("desc")
        debit_raw = cell_text("debit")
        cred_raw  = cell_text("credit")
        bal_raw   = cell_text("balance")

        # fallback: ถ้า col_map ไม่มี debit/credit → ดึง amount ทุกตัวจาก row
        all_money = []
        if "debit" not in col_map and "credit" not in col_map:
            all_texts = " ".join(get_cell_text(c, full_text) for c in cells)
            all_money = extract_money_values(all_texts)

        date = extract_date(date_raw) or extract_date(desc_raw)
        time = extract_time(date_raw) or extract_time(desc_raw)
        debit  = parse_money(debit_raw)
        credit = parse_money(cred_raw)
        balance = parse_money(bal_raw)

        raw_line = clean_text(" | ".join(get_cell_text(c, full_text) for c in cells))
        opening = is_opening_balance(raw_line)

        check_flag = "OK"
        if not date: check_flag = "NO_DATE_REVIEW"

        rows.append({
            "date":       date or "",
            "time":       time,
            "desc":       clean_text(desc_raw) or remove_money_and_date(raw_line),
            "debit":      debit,
            "credit":     credit,
            "balance":    balance,
            "all_money":  all_money,
            "raw":        raw_line,
            "is_opening": opening,
            "check_flag": check_flag,
            "source":     "table",
        })
    return rows


# ============================================================
# 8) FALLBACK: bounding-box layout เมื่อไม่มี table
# ============================================================
def parse_by_bounding_box(doc_dict, full_text):
    """
    เมื่อ Document AI ไม่ส่ง table กลับมา
    ใช้ bounding box x-position แบ่ง column
    zone:  x < 0.25 → date    | 0.25-0.45 → desc
           0.45-0.60 → debit  | 0.60-0.75 → credit  | > 0.75 → balance
    """
    pages = doc_dict.get("pages", [])
    rows_by_y = {}   # y_bucket → {zone: text}

    for page in pages:
        page_w = page.get("dimension", {}).get("width", 1) or 1
        page_h = page.get("dimension", {}).get("height", 1) or 1

        for token in page.get("tokens", []):
            layout = token.get("layout", {})
            verts  = (layout.get("bounding_poly", {}).get("normalized_vertices", [])
                     or layout.get("boundingPoly", {}).get("normalizedVertices", []))
            if not verts: continue

            xs = [v.get("x", 0) for v in verts]
            ys = [v.get("y", 0) for v in verts]
            x_center = sum(xs) / len(xs)
            y_center = sum(ys) / len(ys)
            y_bucket = round(y_center * 100)  # group rows by y (1% bins)

            segs = (layout.get("text_anchor", {}).get("text_segments", [])
                   or layout.get("textAnchor", {}).get("textSegments", []))
            token_text = ""
            for seg in segs:
                s = int(seg.get("start_index", seg.get("startIndex", 0)) or 0)
                e = int(seg.get("end_index",   seg.get("endIndex",   0)) or 0)
                token_text += full_text[s:e]
            token_text = clean_text(token_text)
            if not token_text: continue

            if x_center < 0.25:         zone = "date"
            elif x_center < 0.45:       zone = "desc"
            elif x_center < 0.60:       zone = "debit"
            elif x_center < 0.75:       zone = "credit"
            else:                        zone = "balance"

            if y_bucket not in rows_by_y:
                rows_by_y[y_bucket] = {}
            rows_by_y[y_bucket].setdefault(zone, []).append(token_text)

    parsed = []
    for y_bucket in sorted(rows_by_y.keys()):
        zones = rows_by_y[y_bucket]
        date_text  = " ".join(zones.get("date",    []))
        desc_text  = " ".join(zones.get("desc",    []))
        debit_text = " ".join(zones.get("debit",   []))
        cred_text  = " ".join(zones.get("credit",  []))
        bal_text   = " ".join(zones.get("balance", []))

        date    = extract_date(date_text) or extract_date(desc_text)
        time    = extract_time(date_text) or extract_time(desc_text)
        debit   = parse_money(debit_text)
        credit  = parse_money(cred_text)
        balance = parse_money(bal_text)
        raw     = clean_text(f"{date_text} {desc_text} {debit_text} {cred_text} {bal_text}")
        opening = is_opening_balance(raw)

        if not date and not any([debit, credit, balance]): continue

        parsed.append({
            "date":       date or "",
            "time":       time,
            "desc":       clean_text(desc_text) or remove_money_and_date(raw),
            "debit":      debit,
            "credit":     credit,
            "balance":    balance,
            "all_money":  [],
            "raw":        raw,
            "is_opening": opening,
            "check_flag": "OK" if date else "NO_DATE_REVIEW",
            "source":     "bbox",
        })
    return parsed


# ============================================================
# 9) BALANCE-CHAIN VERIFICATION
# ============================================================
def verify_and_fill_amounts(parsed_rows):
    """
    วิ่งตาม balance chain:
    • ถ้า debit/credit ชัดเจนอยู่แล้ว → ตรวจว่า balance ถูกต้อง
    • ถ้าไม่มี debit/credit → คำนวณจาก prev_balance/balance diff
    • ถ้า all_money มีให้ → ลองเดา debit/credit จากนั้น
    """
    result = []
    prev_balance = None

    for row in parsed_rows:
        debit   = row.get("debit")
        credit  = row.get("credit")
        balance = row.get("balance")
        is_open = row.get("is_opening", False)
        all_money = row.get("all_money", [])
        check   = row.get("check_flag", "OK")

        expected_balance = None
        diff_py          = None
        final_debit      = 0.0
        final_credit     = 0.0
        balance_check    = check

        if is_open:
            balance_check = "OPENING_BALANCE"
            if is_valid_number(balance):
                prev_balance = float(balance)
        else:
            # ── CASE A: debit XOR credit มีค่า (ชัดเจนจาก table column)
            if is_valid_number(debit) and not is_valid_number(credit):
                final_debit = float(debit)
                if is_valid_number(prev_balance) and is_valid_number(balance):
                    expected_balance = round(prev_balance - final_debit, 2)
                    diff_py = round(float(balance) - expected_balance, 2)
                    balance_check = "OK_DEBIT" if abs(diff_py) <= BALANCE_TOLERANCE else "DEBIT_DIFF_REVIEW"
                else:
                    balance_check = "OK_DEBIT_NO_PREV"

            elif is_valid_number(credit) and not is_valid_number(debit):
                final_credit = float(credit)
                if is_valid_number(prev_balance) and is_valid_number(balance):
                    expected_balance = round(prev_balance + final_credit, 2)
                    diff_py = round(float(balance) - expected_balance, 2)
                    balance_check = "OK_CREDIT" if abs(diff_py) <= BALANCE_TOLERANCE else "CREDIT_DIFF_REVIEW"
                else:
                    balance_check = "OK_CREDIT_NO_PREV"

            # ── CASE B: มีทั้งสอง column → หาว่าช่องไหนไม่ว่าง
            elif is_valid_number(debit) and is_valid_number(credit):
                # ทั้งสองมีค่า → ผิดปกติ; ตรวจ chain
                if is_valid_number(prev_balance) and is_valid_number(balance):
                    diff = round(float(balance) - float(prev_balance), 2)
                    if diff < 0:
                        final_debit  = abs(diff); balance_check = "BOTH_COL_DEBIT_INFERRED"
                    else:
                        final_credit = abs(diff); balance_check = "BOTH_COL_CREDIT_INFERRED"
                else:
                    final_debit = float(debit); balance_check = "BOTH_COL_REVIEW"

            # ── CASE C: ไม่มี debit/credit column → ใช้ balance diff
            else:
                if is_valid_number(prev_balance) and is_valid_number(balance):
                    diff = round(float(balance) - float(prev_balance), 2)
                    if abs(diff) > BALANCE_TOLERANCE:
                        if diff < 0:
                            final_debit  = abs(diff)
                            balance_check = "INFER_DEBIT_FROM_BAL"
                        else:
                            final_credit = abs(diff)
                            balance_check = "INFER_CREDIT_FROM_BAL"
                    else:
                        balance_check = "ZERO_CHANGE"
                elif all_money:
                    # fallback: amount = all_money[0], balance = all_money[-1]
                    if len(all_money) >= 2:
                        amt = all_money[0]
                        bal_guess = all_money[-1]
                        if is_valid_number(prev_balance):
                            diff = round(bal_guess - float(prev_balance), 2)
                            if abs(diff - amt) <= BALANCE_TOLERANCE:
                                final_credit = amt; balance_check = "FALLBACK_CREDIT"
                            elif abs(diff + amt) <= BALANCE_TOLERANCE:
                                final_debit  = amt; balance_check = "FALLBACK_DEBIT"
                            else:
                                balance_check = "FALLBACK_REVIEW"
                        else:
                            balance_check = "NO_PREV_REVIEW"
                    else:
                        balance_check = "SINGLE_MONEY_REVIEW"

            if is_valid_number(balance):
                prev_balance = float(balance)

        result.append({
            **row,
            "final_debit":        final_debit,
            "final_credit":       final_credit,
            "expected_balance":   expected_balance,
            "diff_py":            diff_py,
            "balance_check":      balance_check,
        })
    return result


# ============================================================
# 10) PIPELINE: doc_dict → DataFrames
# ============================================================
def process_document(doc_dict, full_text, bank, filename):
    pages = doc_dict.get("pages", [])

    # รวม rows จากทุก page ทุก table
    all_rows = []
    has_tables = False
    for page in pages:
        tables = page.get("tables", [])
        for table in tables:
            rows = parse_table(table, full_text, bank)
            if rows:
                has_tables = True
                all_rows.extend(rows)

    # fallback bounding box ถ้าไม่มี table
    if not has_tables or len(all_rows) == 0:
        all_rows = parse_by_bounding_box(doc_dict, full_text)

    # กรอง row ที่ไม่มีข้อมูลเลย
    all_rows = [r for r in all_rows if r.get("date") or is_valid_number(r.get("balance"))]

    # ตรวจ balance chain
    verified = verify_and_fill_amounts(all_rows)

    # สร้าง df_txn
    tx_rows = []
    for v in verified:
        if v.get("is_opening") or (v["final_debit"] == 0.0 and v["final_credit"] == 0.0):
            if not v.get("is_opening"): continue  # skip zero-amount non-opening
        tx_rows.append({
            "ลำดับ":        len(tx_rows) + 1,
            "วันที่เดือนปี": v["date"],
            "เวลา":          v["time"],
            "รายการ":        v["desc"][:60],
            "เดบิต":         v["final_debit"],
            "เครดิต":        v["final_credit"],
            "ยอดคงเหลือ":    float(v["balance"]) if is_valid_number(v["balance"]) else 0.0,
            "รายละเอียด":    v["desc"],
            "ช่องทาง":       "Document AI",
            "source":        v.get("source", ""),
            "ไฟล์ต้นฉบับ":  filename,
        })

    # สร้าง check_df
    check_rows = []
    for i, v in enumerate(verified, start=1):
        check_rows.append({
            "seq":                i,
            "วันที่เดือนปี":     v["date"],
            "เวลา":              v["time"],
            "เดบิต (อ่านได้)":   v.get("debit"),
            "เครดิต (อ่านได้)":  v.get("credit"),
            "เดบิต (สุดท้าย)":   v["final_debit"],
            "เครดิต (สุดท้าย)":  v["final_credit"],
            "ยอดคงเหลือ":        v.get("balance"),
            "expected_balance":  v.get("expected_balance"),
            "diff_py":           v.get("diff_py"),
            "balance_check":     v["balance_check"],
            "source":            v.get("source", ""),
            "raw_line_text":     v["raw"],
            "ไฟล์ต้นฉบับ":      filename,
        })

    df_txn   = pd.DataFrame(tx_rows)
    check_df = pd.DataFrame(check_rows)
    return df_txn, check_df


# ============================================================
# 11) SALARY DETECTION + SUMMARY
# ============================================================
def detect_salary(df_txn):
    if df_txn.empty: return pd.DataFrame()
    SALARY_KW = ["SALARY","PAYROLL","SAL","PAYR","เงินเดือน"]
    BONUS_KW  = ["BONUS","INCENTIVE","COMMISSION","โบนัส","อินเซนทีฟ","ค่าคอม"]
    rows = []
    for _, row in df_txn.iterrows():
        cr = float(row.get("เครดิต", 0) or 0)
        if cr <= 0: continue
        desc = str(row.get("รายละเอียด","")).upper()
        if any(k.upper() in desc for k in BONUS_KW):
            rows.append({"กลุ่ม":"รายได้พิเศษ/โบนัส","วันที่":row["วันที่เดือนปี"],"จำนวนเงิน":cr})
        elif any(k.upper() in desc for k in SALARY_KW):
            rows.append({"กลุ่ม":"เงินเดือน","วันที่":row["วันที่เดือนปี"],"จำนวนเงิน":cr})
    return pd.DataFrame(rows)


def build_summary_df(df_txn, owner, acc_no, bank_name):
    if df_txn.empty:
        last_bal, s_date, e_date = 0.0, "", ""
    else:
        valid   = pd.to_numeric(df_txn["ยอดคงเหลือ"], errors="coerce").dropna()
        last_bal = float(valid.iloc[-1]) if len(valid) else 0.0
        s_date  = df_txn["วันที่เดือนปี"].iloc[0]
        e_date  = df_txn["วันที่เดือนปี"].iloc[-1]
    rows = [
        {"รายการ":"เจ้าของบัญชี",        "ข้อมูล": owner},
        {"รายการ":"เลขที่บัญชีเงินฝาก",  "ข้อมูล": acc_no},
        {"รายการ":"ธนาคาร",               "ข้อมูล": bank_name},
        {"รายการ":"จำนวนรายการ",          "ข้อมูล": len(df_txn)},
        {"รายการ":"รวมเดบิต",             "ข้อมูล": float(df_txn["เดบิต"].sum()) if not df_txn.empty else 0.0},
        {"รายการ":"รวมเครดิต",            "ข้อมูล": float(df_txn["เครดิต"].sum()) if not df_txn.empty else 0.0},
        {"รายการ":"ยอดคงเหลือล่าสุด",    "ข้อมูล": last_bal},
        {"รายการ":"วันที่เริ่มต้น",       "ข้อมูล": s_date},
        {"รายการ":"วันที่สิ้นสุด",        "ข้อมูล": e_date},
    ]
    df_income = detect_salary(df_txn)
    for grp in ["เงินเดือน","รายได้พิเศษ/โบนัส"]:
        rows.append({"รายการ":"หมวดหมู่รายได้","ข้อมูล": grp})
        if not df_income.empty and grp in df_income["กลุ่ม"].values:
            g = df_income[df_income["กลุ่ม"]==grp].copy()
            total = float(g["จำนวนเงิน"].sum())
            rows += [
                {"รายการ":f"{grp} - สถานะ",            "ข้อมูล":"พบ"},
                {"รายการ":f"{grp} - จำนวนรายการ",      "ข้อมูล": len(g)},
                {"รายการ":f"{grp} - ยอดรวม",           "ข้อมูล": total},
                {"รายการ":f"{grp} - ค่าเฉลี่ยต่อเดือน","ข้อมูล": total/len(g)},
            ]
            for i, r in enumerate(g.itertuples(), 1):
                rows += [{"รายการ":f"{grp} - รายการที่ {i}","ข้อมูล": r.วันที่},
                         {"รายการ":f"{grp} - ยอด","ข้อมูล": r.จำนวนเงิน}]
        else:
            rows += [
                {"รายการ":f"{grp} - สถานะ","ข้อมูล":"ไม่พบ"},
                {"รายการ":f"{grp} - จำนวนรายการ","ข้อมูล":0},
                {"รายการ":f"{grp} - ยอดรวม","ข้อมูล":0.0},
            ]
        rows.append({"รายการ":"","ข้อมูล":""})
    rows.append({"รายการ":"สร้างไฟล์เมื่อ","ข้อมูล": datetime.now().strftime("%d/%m/%Y %H:%M:%S")})
    return pd.DataFrame(rows)


# ============================================================
# 12) BANK STATEMENT 1 SHEET  (เหมือนเดิม)
# ============================================================
def calculate_30day_blocks(df_txn, period_days=30, periods=3):
    empty = {"daily": pd.DataFrame(columns=["วันที่","ยอดคงเหลือสิ้นวัน","ช่วงที่"]), "blocks": []}
    if df_txn.empty: return empty
    work = df_txn.copy()
    work["_date"] = pd.to_datetime(
        work["วันที่เดือนปี"].astype(str).str.replace("-","/"), dayfirst=True, errors="coerce")
    work = work.dropna(subset=["_date"])
    if work.empty: return empty
    work["ยอดคงเหลือ"] = pd.to_numeric(work["ยอดคงเหลือ"], errors="coerce").fillna(0.0)
    work = work.reset_index(drop=True)
    work["_row_order"] = work.index
    work = work.sort_values(["_date", "_row_order"])
    work["_day"] = work["_date"].dt.normalize()
    daily_txn = work.groupby("_day").agg(ยอดคงเหลือสิ้นวัน=("ยอดคงเหลือ","last")).sort_index()
    full_days = pd.date_range(start=daily_txn.index.min(), end=daily_txn.index.max(), freq="D")
    daily = daily_txn.reindex(full_days)
    daily["ยอดคงเหลือสิ้นวัน"] = daily["ยอดคงเหลือสิ้นวัน"].ffill().fillna(0.0)
    daily = daily.tail(period_days*periods).reset_index().rename(columns={"index":"วันที่"})
    daily["ลำดับ"] = range(1, len(daily)+1)
    daily["ช่วงที่"] = ((daily["ลำดับ"]-1) // period_days) + 1
    blocks = []
    for b in range(1, periods+1):
        blk = daily[daily["ช่วงที่"]==b].copy()
        blocks.append({"ช่วงที่":b,"จำนวนวัน":len(blk),
                        "ผลรวม":float(blk["ยอดคงเหลือสิ้นวัน"].sum()) if len(blk) else 0.0,
                        "ค่าเฉลี่ย":float(blk["ยอดคงเหลือสิ้นวัน"].mean()) if len(blk) else 0.0,
                        "data":blk})
    return {"daily":daily,"blocks":blocks}


def add_bank_statement_sheet(path, df_txn, owner, acc_no, bank_name):
    result = calculate_30day_blocks(df_txn)
    blocks = result["blocks"]
    wb = load_workbook(path)
    ws = wb.create_sheet("BANK STATEMENT 1")
    title_blue  = PatternFill("solid", fgColor="6FA1E8")
    header_blue = PatternFill("solid", fgColor="A9C4F5")
    light_blue  = PatternFill("solid", fgColor="DDEBF7")
    yellow      = PatternFill("solid", fgColor="FFFF00")
    gray        = PatternFill("solid", fgColor="D9D9D9")
    heavy = Border(left=Side(style="medium",color="000000"), right=Side(style="medium",color="000000"),
                   top=Side(style="medium",color="000000"),  bottom=Side(style="medium",color="000000"))
    thin  = Border(left=Side(style="thin",color="D9D9D9"),   right=Side(style="thin",color="D9D9D9"),
                   top=Side(style="thin",color="D9D9D9"),    bottom=Side(style="thin",color="D9D9D9"))
    ws.merge_cells("A1:I1"); ws["A1"]="BANK STATEMENT 1"
    ws["A1"].font=Font(bold=True,size=13); ws["A1"].alignment=Alignment(horizontal="center",vertical="center")
    ws["A1"].fill=title_blue
    for idx in range(3):
        r=2+idx
        ws.cell(r,1,1); ws.cell(r,2,f"เดือนที่ {idx+1}")
        ws.cell(r,3,f"=ROUND({['C37','F37','I37'][idx]},2)")
    ws.cell(2,5,"เฉลี่ยต่อเดือน")
    ws.cell(2,6,"=IF((C39>0)+(F39>0)+(I39>0)=0,0,ROUND((IF(C39>0,C37,0)+IF(F39>0,F37,0)+IF(I39>0,I37,0))/((C39>0)+(F39>0)+(I39>0)),2))")
    ws.cell(3,5,"จำนวนวันรวม"); ws.cell(3,6,"=SUM(C39,F39,I39)")
    ws.cell(4,5,"ประเมินวงเงินได้"); ws.cell(4,6,"=ROUND(F2*G3*G4,2)")
    ws.merge_cells("H2:I2"); ws["H2"]="Account Detail"
    ws.cell(3,8,"ธนาคาร"); ws.cell(3,9,"เลขที่บัญชี")
    ws.cell(4,8,bank_name); ws.cell(4,9,acc_no)
    for r in range(2,5):
        for c in range(1,10):
            cell=ws.cell(r,c); cell.fill=header_blue; cell.border=heavy
            cell.font=Font(bold=True); cell.alignment=Alignment(horizontal="center",vertical="center")
    for addr in ["C2","C3","C4","F2","F3"]: ws[addr].fill=light_blue
    for idx, block in enumerate(blocks):
        sc=([1,4,7])[idx]; nc,dc,ac=sc,sc+1,sc+2
        for col,val in [(nc,"N"),(dc,"DATE"),(ac,"AMOUNT")]:
            cell=ws.cell(5,col,val); cell.fill=header_blue; cell.font=Font(bold=True)
            cell.border=heavy; cell.alignment=Alignment(horizontal="center",vertical="center")
        data=block.get("data",pd.DataFrame())
        for i in range(1,31):
            er=6+i-1
            ws.cell(er,nc,i).fill=light_blue
            ws.cell(er,dc).fill=light_blue
            if i<=len(data):
                rd=data.iloc[i-1]
                dv=rd.get("วันที่",""); av=float(rd.get("ยอดคงเหลือสิ้นวัน",0))
                ws.cell(er,dc,dv); ws.cell(er,ac,av)
                if i>1:
                    pv=float(data.iloc[i-2].get("ยอดคงเหลือสิ้นวัน",0))
                    if round(av,2)!=round(pv,2):
                        ws.cell(er,dc).fill=yellow; ws.cell(er,ac).fill=yellow
            for col in [nc,dc,ac]:
                cell=ws.cell(er,col); cell.border=thin
                cell.alignment=Alignment(horizontal="center" if col!=ac else "right",vertical="center")
            ws.cell(er,dc).number_format="d mmm yy"
            ws.cell(er,ac).number_format="#,##0.00"
        al=ws.cell(5,ac).column_letter
        ws.cell(36,nc,"รวม"); ws.cell(36,ac,f"=SUM({al}6:{al}35)")
        ws.cell(37,ac,f"=IF({al}39=0,0,{al}36/{al}39)")
        ws.cell(38,ac,"จำนวนวัน"); ws.cell(39,ac,f"=COUNT({al}6:{al}35)")
        for r in range(36,40):
            for c in [nc,dc,ac]:
                cell=ws.cell(r,c); cell.border=heavy
                cell.alignment=Alignment(horizontal="center" if c!=ac else "right",vertical="center")
                if c==ac: cell.number_format="#,##0.00"
        ws.cell(38,ac).fill=gray; ws.cell(38,ac).font=Font(bold=True)
    ws.sheet_view.showGridLines=False; ws.freeze_panes="A5"
    for letter,width in {"A":5,"B":14,"C":16,"D":5,"E":14,"F":16,"G":5,"H":14,"I":16}.items():
        ws.column_dimensions[letter].width=width
    wb.save(path)


# ============================================================
# 13) EXCEL EXPORT
# ============================================================
def auto_fit_worksheet(ws):
    ws.freeze_panes = "A2"
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    headers = {c: clean_text(ws.cell(row=1,column=c).value) for c in range(1, ws.max_column+1)}
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=False)
            if cell.row == 1:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")
    for col_cells in ws.columns:
        letter = col_cells[0].column_letter
        col_idx = col_cells[0].column
        width = estimate_text_width(headers.get(col_idx, ""))
        for cell in col_cells:
            if cell.value is not None:
                width = max(width, estimate_text_width(str(cell.value)))
        ws.column_dimensions[letter].width = max(8, min(width+1.5, 40))


def format_summary_sheet(ws):
    ws.sheet_view.showGridLines = False
    h_fill = PatternFill("solid", fgColor="1F4E79")
    s_fill = PatternFill("solid", fgColor="DDEBF7")
    sub_fill = PatternFill("solid", fgColor="EAF4EA")
    for cell in ws[1]:
        cell.fill = h_fill; cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for r in range(2, ws.max_row+1):
        label = clean_text(ws.cell(r,1).value)
        if label == "หมวดหมู่รายได้":
            for c in range(1,3): ws.cell(r,c).fill=s_fill; ws.cell(r,c).font=Font(bold=True)
        if any(k in label for k in ["สถานะ","จำนวนรายการ","ยอดรวม","ค่าเฉลี่ยต่อเดือน"]):
            ws.cell(r,1).fill=sub_fill; ws.cell(r,2).fill=sub_fill; ws.cell(r,1).font=Font(bold=True)
        if isinstance(ws.cell(r,2).value,(int,float)):
            ws.cell(r,2).number_format="#,##0.00"


def build_output_excel(df_account, df_txn, owner, acc_no, bank_name,
                        raw_text, doc_json_text, check_df):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp_path = tmp.name

    review_df = pd.DataFrame()
    if check_df is not None and not check_df.empty and "balance_check" in check_df.columns:
        review_df = check_df[
            check_df["balance_check"].astype(str).str.contains("REVIEW|BROKEN", regex=True, na=False)
        ].copy()

    with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
        df_account.to_excel(writer, sheet_name="เจ้าของบัญชี", index=False)
        df_txn.to_excel(writer, sheet_name="รายละเอียด เดบิต-เครดิต", index=False)
        build_summary_df(df_txn, owner, acc_no, bank_name).to_excel(writer, sheet_name="สรุปยอด", index=False)
        check_df.to_excel(writer, sheet_name="check", index=False)
        review_df.to_excel(writer, sheet_name="ocr_review", index=False)
        pd.DataFrame({"raw_text":[raw_text]}).to_excel(writer, sheet_name="raw_text", index=False)
        pd.DataFrame({"document_json":[doc_json_text]}).to_excel(writer, sheet_name="document_json", index=False)

    add_bank_statement_sheet(tmp_path, df_txn, owner, acc_no, bank_name)

    # format
    wb = load_workbook(tmp_path)
    for ws in wb.worksheets:
        if ws.title != "BANK STATEMENT 1":
            auto_fit_worksheet(ws)
            if ws.title == "สรุปยอด": format_summary_sheet(ws)
    wb.save(tmp_path)

    with open(tmp_path, "rb") as f: data = f.read()
    os.unlink(tmp_path)
    return data


# ============================================================
# 14) MAIN PIPELINE PER FILE
# ============================================================
def process_one_file(uploaded_file, selected_bank):
    document, doc_dict, raw_text = call_document_ai(uploaded_file)
    if not raw_text.strip() and not doc_dict.get("pages"):
        raise ValueError("Document AI อ่านข้อความไม่เจอ")

    doc_json_text = json.dumps(doc_dict, ensure_ascii=False, indent=2)
    active_bank   = selected_bank if selected_bank != "AUTO" else detect_bank(raw_text)
    df_txn, check_df = process_document(doc_dict, raw_text, active_bank, uploaded_file.name)
    owner, acc_no    = extract_account_info(raw_text)

    return {
        "raw_text":      raw_text,
        "doc_json_text": doc_json_text,
        "active_bank":   active_bank,
        "df_txn":        df_txn,
        "check_df":      check_df,
        "owner":         owner,
        "acc_no":        acc_no,
    }


# ============================================================
# 15) UI
# ============================================================
def render_section(step_no, title):
    st.markdown(
        f'<div class="section-title"><span class="step-number">{step_no}</span>'
        f'<span>{title}</span></div>',
        unsafe_allow_html=True
    )


def main_app():
    st.title("STM Document AI Parser v2")
    st.caption("JSON Structure Edition: อ่าน Table + Bounding Box โดยตรง — ไม่ต้อง keyword / BANK_CONFIGS")
    st.divider()

    render_section("1", "เลือกธนาคาร")
    bank_choice = st.radio(
        "ธนาคาร",
        options=["AUTO","KBANK","BBL","SCB","KRUNGSRI"],
        format_func=lambda x: BANK_LABELS.get(x, x),
        index=0, horizontal=True, label_visibility="collapsed"
    )

    render_section("2", "ตรวจสอบ / แก้ไขข้อมูลบัญชี")
    st.info("ระบบจะพยายามค้นหา ชื่อบัญชี และ เลขที่บัญชี ให้อัตโนมัติ")
    col1, col2 = st.columns(2)
    owner_input = col1.text_input("ชื่อเจ้าของบัญชี", placeholder="เว้นว่างให้ระบบหาอัตโนมัติ")
    acc_input   = col2.text_input("เลขที่บัญชี",       placeholder="เว้นว่างให้ระบบหาอัตโนมัติ")

    render_section("3", "อัปโหลด Statement")
    uploaded_files = st.file_uploader(
        "เลือกไฟล์ Statement JPG / PNG / PDF",
        type=["jpg","jpeg","png","pdf"],
        accept_multiple_files=True
    )

    render_section("4", "ประมวลผล")
    c1, c2 = st.columns(2)
    process_clicked = c1.button("เริ่มประมวลผล", type="primary",
                                 use_container_width=True, disabled=not uploaded_files)
    if c2.button("ออกจากระบบ", use_container_width=True):
        st.session_state.authenticated = False
        st.session_state.results = None
        st.rerun()

    if process_clicked:
        st.session_state.results = None
        progress = st.progress(0, text="เริ่มประมวลผล...")
        status   = st.empty()
        try:
            all_txn, all_check, all_raw, all_json = [], [], [], []
            banks, auto_owner, auto_acc = [], "", ""
            total = len(uploaded_files)

            for idx, uf in enumerate(uploaded_files, start=1):
                status.info(f"กำลังประมวลผลไฟล์ {idx}/{total}: {uf.name}")
                progress.progress(int((idx-1)/total*70))
                res = process_one_file(uf, bank_choice)

                if not res["df_txn"].empty:
                    all_txn.append(res["df_txn"])
                if not res["check_df"].empty:
                    tmp_c = res["check_df"].copy()
                    tmp_c["ไฟล์ต้นฉบับ"] = uf.name
                    all_check.append(tmp_c)
                all_raw.append(f"===== {uf.name} =====\n{res['raw_text']}")
                all_json.append(f"===== {uf.name} =====\n{res['doc_json_text']}")
                banks.append(res["active_bank"])
                if not auto_owner or auto_owner=="ไม่ระบุ":
                    if res["owner"] != "ไม่ระบุ": auto_owner = res["owner"]
                if not auto_acc or auto_acc=="ไม่ระบุ":
                    if res["acc_no"] != "ไม่ระบุ": auto_acc = res["acc_no"]

            progress.progress(80, text="กำลังสร้าง Excel...")
            EMPTY_COLS = ["ลำดับ","วันที่เดือนปี","เวลา","รายการ","เดบิต","เครดิต",
                          "ยอดคงเหลือ","รายละเอียด","ช่องทาง","source","ไฟล์ต้นฉบับ"]
            final_df = (pd.concat(all_txn, ignore_index=True)
                        if all_txn else pd.DataFrame(columns=EMPTY_COLS))
            final_df["ลำดับ"] = range(1, len(final_df)+1)
            check_df = (pd.concat(all_check, ignore_index=True)
                        if all_check else pd.DataFrame())

            owner   = owner_input.strip() or auto_owner or "ไม่ระบุ"
            acc_no  = acc_input.strip()   or auto_acc   or "ไม่ระบุ"
            active_bank = (banks[0] if bank_choice=="AUTO" and banks else bank_choice)
            bank_name   = BANK_LABELS.get(active_bank, active_bank)
            df_account  = pd.DataFrame([{
                "เจ้าของบัญชี": owner,
                "เลขที่บัญชีเงินฝาก": acc_no,
                "ธนาคาร": bank_name,
            }])

            excel_bytes = build_output_excel(
                df_account, final_df, owner, acc_no, bank_name,
                "\n\n".join(all_raw), "\n\n".join(all_json), check_df
            )
            excel_name = f"Statement_{owner}.xlsx".replace("/","_").replace("\\","_")

            st.session_state.results = {
                "df": final_df, "check": check_df,
                "excel": excel_bytes, "excel_name": excel_name,
                "owner": owner, "acc_no": acc_no, "bank_name": bank_name,
            }
            progress.progress(100, text="เสร็จสิ้น ✅")
            status.success(f"ประมวลผลสำเร็จ พบ {len(final_df):,} รายการ")

        except Exception as e:
            progress.empty(); status.error(f"ประมวลผลไม่สำเร็จ: {e}"); st.exception(e)

    if st.session_state.results:
        res = st.session_state.results
        df  = res["df"]; check_df = res["check"]
        st.divider()
        st.success(f"ประมวลผลสำเร็จ! พบ {len(df):,} รายการ")

        debit_sum  = float(pd.to_numeric(df.get("เดบิต",  pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
        credit_sum = float(pd.to_numeric(df.get("เครดิต", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
        last_balance = 0.0
        if not df.empty:
            bal = pd.to_numeric(df.get("ยอดคงเหลือ", pd.Series(dtype=float)), errors="coerce").dropna()
            if len(bal): last_balance = float(bal.iloc[-1])

        review_count = 0
        if check_df is not None and not check_df.empty and "balance_check" in check_df.columns:
            review_count = int(check_df["balance_check"].astype(str)
                               .str.contains("REVIEW|BROKEN", regex=True, na=False).sum())

        # ── แสดง source breakdown (table vs bbox)
        if "source" in df.columns:
            src = df["source"].value_counts().to_dict()
            table_n = src.get("table", 0); bbox_n = src.get("bbox", 0)
            st.caption(f"อ่านจาก table structure: {table_n:,} รายการ | bounding box fallback: {bbox_n:,} รายการ")

        m1,m2,m3,m4 = st.columns(4)
        m1.metric("เจ้าของบัญชี",  res["owner"])
        m2.metric("จำนวนรายการ",  f"{len(df):,}")
        m3.metric("รวมเดบิต",     f"{debit_sum:,.2f}")
        m4.metric("รวมเครดิต",    f"{credit_sum:,.2f}")
        m5,m6,m7 = st.columns(3)
        m5.metric("ยอดคงเหลือล่าสุด", f"{last_balance:,.2f}")
        m6.metric("ต้องตรวจ",         f"{review_count:,}")
        m7.metric("ธนาคาร",           res["bank_name"])

        st.download_button(
            "📥 ดาวน์โหลดไฟล์ Excel",
            data=res["excel"], file_name=res["excel_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary"
        )

        st.markdown("#### ตัวอย่างข้อมูลที่ประมวลผลได้")
        if df.empty:
            st.warning("ไม่พบรายการธุรกรรม")
        else:
            st.dataframe(
                df.head(30).style.format({"เดบิต":"{:,.2f}","เครดิต":"{:,.2f}","ยอดคงเหลือ":"{:,.2f}"}),
                use_container_width=True, height=400
            )

        tab1, tab2 = st.tabs(["รายละเอียด เดบิต-เครดิต","check"])
        with tab1: st.dataframe(df, use_container_width=True, height=520)
        with tab2: st.dataframe(check_df, use_container_width=True, height=520)

        if st.button("ล้างข้อมูลหลังใช้งาน", use_container_width=True):
            st.session_state.results = None; st.rerun()


# ============================================================
# ENTRY POINT
# ============================================================
if not st.session_state.authenticated:
    login_page()
else:
    main_app()

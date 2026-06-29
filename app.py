# ============================================================
# STM Document AI Bank Statement Parser - Streamlit Edition
# JPG / PNG / PDF -> Google Document AI -> Excel
# Output sheets:
# 1) เจ้าของบัญชี
# 2) รายละเอียด เดบิต-เครดิต
# 3) สรุปยอด
# 4) BANK STATEMENT 1
# 5) check
# 6) ocr_review
# 7) raw_text
# 8) document_json
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
st.set_page_config(page_title="STM Document AI Parser", layout="centered")

st.markdown("""
<style>
.block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 980px; }
h1 { font-size: 2.35rem !important; font-weight: 800 !important; }
.section-title { display:flex; align-items:center; gap:8px; font-weight:800; font-size:1.15rem; margin:18px 0 10px 0; }
.section-title .step-number { display:inline-flex; align-items:center; justify-content:center; min-width:30px; height:30px; border-radius:999px; background:#4c82fb; color:white; font-weight:800; }
div[data-testid="stButton"] button, div[data-testid="stDownloadButton"] button { border-radius:10px; font-weight:750; }
div[data-testid="stFileUploader"] section { border-radius:10px; }
</style>
""", unsafe_allow_html=True)


# ============================================================
# 1) CONFIG
# ============================================================
BALANCE_TOLERANCE = 0.05
FORCE_YEAR = datetime.now().year

BANK_LABELS = {
    "AUTO": "ตรวจจับอัตโนมัติ",
    "KBANK": "กสิกรไทย (KBANK)",
    "BBL": "กรุงเทพ (BBL)",
    "SCB": "ไทยพาณิชย์ (SCB)",
    "KRUNGSRI": "กรุงศรี (Krungsri)",
    "DEFAULT": "ไม่ระบุ / JSON Data",
}

BANK_CONFIGS = {
    "DEFAULT": {
        "bank_keywords": [],
        "opening_keywords": ["ยอด ยก มา", "ยอดยกมา", "ยอด ยก", "ยก มา", "BALANCE BROUGHT FORWARD", "BROUGHT FORWARD", "BALANCE B/F", "B/F"],
        "credit_keywords": ["รับ โอน", "รับโอน", "ฝาก", "เงิน เข้า", "เงินเข้า", "ดอกเบี้ย", "คืนเงิน", "CREDIT", "DEPOSIT", "TRANSFER IN", "TRF FROM", "TRF FR", "SALARY", "PAYROLL"],
        "debit_keywords": ["โอน เงิน", "โอนเงิน", "ถอน", "ถอนเงิน", "หัก", "ชำระ", "จ่าย", "ค่าธรรมเนียม", "DEBIT", "WITHDRAW", "PAYMENT", "FEE", "TRANSFER OUT", "TRF TO", "ATM", "PMT"],
    },
    "KBANK": {
        "bank_keywords": ["KASIKORN", "KBANK", "K PLUS", "กสิกร"],
        "opening_keywords": ["ยอด ยก มา", "ยอดยกมา", "BROUGHT FORWARD"],
        "credit_keywords": ["รับ โอน", "รับโอน", "ฝาก", "เงิน เข้า", "CREDIT", "DEPOSIT"],
        "debit_keywords": ["โอน เงิน", "โอนเงิน", "ถอน", "จ่าย", "PAYMENT", "WITHDRAW", "FEE"],
    },
    "BBL": {
        "bank_keywords": ["BANGKOK BANK", "BANGKOKBANK", "ธนาคารกรุงเทพ", "BBL", "BUALUANG"],
        "opening_keywords": ["B/F", "BALANCE B/F", "BROUGHT FORWARD", "ยอดยกมา"],
        "credit_keywords": ["TRF FR", "TRF FROM", "TRANSFER IN", "CREDIT", "DEPOSIT", "รับโอน", "ฝาก", "SMART"],
        "debit_keywords": ["TRF TO", "PMT", "CASH W/D", "WITHDRAW", "PAYMENT", "DEBIT", "FEE"],
    },
    "SCB": {
        "bank_keywords": ["SCB", "ไทยพาณิชย์", "SIAM COMMERCIAL BANK"],
        "opening_keywords": ["ยอดเงินคงเหลือยกมา", "BROUGHT FORWARD", "B/F"],
        "credit_keywords": ["รับ โอน", "รับโอน", "ฝาก", "เงิน เข้า", "CREDIT", "DEPOSIT", "TRANSFER IN"],
        "debit_keywords": ["โอน ไป", "โอนไป", "จ่าย", "ถอน", "PAYMENT", "WITHDRAW", "FEE", "TRANSFER OUT"],
    },
    "KRUNGSRI": {
        "bank_keywords": ["KRUNGSRI", "กรุงศรี", "AYUDHYA", "BANK OF AYUDHYA"],
        "opening_keywords": ["ยอดยกมา", "BROUGHT FORWARD", "B/F"],
        "credit_keywords": ["รับ โอน", "รับโอน", "ฝาก", "เงิน เข้า", "CREDIT", "DEPOSIT", "TRANSFER IN"],
        "debit_keywords": ["โอน เงิน", "ถอน", "จ่าย", "PAYMENT", "WITHDRAW", "FEE", "TRANSFER OUT"],
    },
}


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
# 3) DOCUMENT AI
# ============================================================
def get_documentai_config():
    try:
        return (
            str(st.secrets["DOCUMENTAI_PROJECT_ID"]),
            str(st.secrets["DOCUMENTAI_LOCATION"]),
            str(st.secrets["DOCUMENTAI_PROCESSOR_ID"]),
        )
    except Exception as e:
        st.error(f"ไม่พบค่า DOCUMENTAI_PROJECT_ID / LOCATION / PROCESSOR_ID ใน Secrets: {e}")
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
            info,
            scopes=["https://www.googleapis.com/auth/cloud-platform"],
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
    if name.endswith(".pdf"):
        return "application/pdf"
    if name.endswith(".jpg") or name.endswith(".jpeg"):
        return "image/jpeg"
    if name.endswith(".png"):
        return "image/png"
    raise ValueError("รองรับเฉพาะ PDF, JPG, JPEG, PNG")


def process_uploaded_file_with_document_ai(uploaded_file):
    client, processor_name = get_documentai_client()
    content = uploaded_file.read()
    uploaded_file.seek(0)
    raw_document = documentai.RawDocument(content=content, mime_type=get_mime_type(uploaded_file.name))
    request = documentai.ProcessRequest(name=processor_name, raw_document=raw_document)
    result = client.process_document(request=request)
    return result.document


def document_to_json_text(document):
    doc_dict = MessageToDict(document._pb, preserving_proto_field_name=True)
    return json.dumps(doc_dict, ensure_ascii=False, indent=2)


# ============================================================
# 4) TEXT / MONEY HELPERS
# ============================================================
def clean_text(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\x00", " ").strip())


def estimate_text_width(value):
    text = clean_text(value)
    w = 0.0
    for ch in text:
        if "\u0e00" <= ch <= "\u0e7f":
            w += 1.4
        elif ch.isupper():
            w += 1.1
        elif ch.isdigit() or ch in ".,:-/()":
            w += 0.85
        elif ch == " ":
            w += 0.4
        else:
            w += 0.95
    return w


def normalize_amount_text(s):
    s = str(s).strip()
    s = re.sub(r"\s+", "", s)
    ocr_map = {"O": "0", "o": "0", "I": "1", "l": "1", "|": "1", "S": "5", "s": "5", "B": "8", "g": "9", "q": "9", "Z": "2", "z": "2"}
    for c, d in ocr_map.items():
        s = s.replace(c, d)
    if re.match(r"^\d{1,3}\.\d{3}\.\d{2}$", s):
        parts = s.split(".")
        s = parts[0] + "," + parts[1] + "." + parts[2]
    if re.match(r"^\d{1,3}\.\d{3},\d{2}$", s):
        s = s.replace(".", "").replace(",", ".")
    if re.match(r"^\d+,\d{2}$", s):
        s = s.replace(",", ".")
    return s


def parse_money(s):
    s = normalize_amount_text(str(s))
    s = re.sub(r"[^0-9,.\-]", "", s)
    if not s:
        return None
    try:
        return float(s.replace(",", ""))
    except Exception:
        return None


def extract_money_values(text):
    text = str(text)
    ocr_map = {"O": "0", "o": "0", "I": "1", "l": "1", "|": "1", "S": "5", "s": "5", "B": "8", "g": "9", "q": "9", "Z": "2", "z": "2"}
    for c, d in ocr_map.items():
        text = text.replace(c, d)
    patterns = [
        r"(?<!\d)-?\d{1,3}(?:,\d{3})+\.\d{2}(?!\d)",
        r"(?<!\d)-?\d{1,3}(?:\.\d{3})+\.\d{2}(?!\d)",
        r"(?<!\d)-?\d{1,3}(?:\.\d{3})+,\d{2}(?!\d)",
        r"(?<!\d)-?\d+\.\d{2}(?!\d)",
        r"(?<!\d)-?\d+,\d{2}(?!\d)",
    ]
    vals = []
    for pat in patterns:
        for m in re.findall(pat, text):
            v = parse_money(m)
            if v is not None:
                vals.append(v)
    unique = []
    for v in vals:
        if v not in unique:
            unique.append(v)
    return unique


def fix_ocr_date(s):
    s = str(s).strip()
    s = s.replace("O", "0").replace("o", "0")
    s = s.replace("I", "1").replace("l", "1")
    s = s.replace(".", "-").replace("/", "-")
    return s


def normalize_date(s):
    s = fix_ocr_date(s)
    m1 = re.match(r"^(\d{1,2})[-](\d{1,2})[-](\d{2,4})$", s)
    m2 = re.match(r"^(\d{1,2})[-](\d{1,2})$", s)
    if m1:
        day = int(m1.group(1))
        month = int(m1.group(2))
        y = m1.group(3)
        year = 2000 + int(y) if len(y) == 2 else int(y)
    elif m2:
        day = int(m2.group(1))
        month = int(m2.group(2))
        year = FORCE_YEAR
    else:
        return None
    try:
        return datetime(year, month, day).strftime("%d/%m/%Y")
    except Exception:
        return None


def extract_date_from_line(line):
    for tok in str(line).split():
        d = normalize_date(tok)
        if d:
            return d
    return None


def extract_time_from_line(line):
    m = re.search(r"\b\d{1,2}:\d{2}(?::\d{2})?\b", str(line))
    return m.group(0) if m else ""


def get_config(bank):
    return BANK_CONFIGS.get(bank, BANK_CONFIGS["DEFAULT"])


def has_opening_text(text, bank="DEFAULT"):
    up = clean_text(text).upper()
    for kw in get_config(bank).get("opening_keywords", []):
        if kw.upper() in up:
            return True
    for kw in BANK_CONFIGS["DEFAULT"]["opening_keywords"]:
        if kw.upper() in up:
            return True
    return False


def classify_by_keyword(text, bank="DEFAULT"):
    up = clean_text(text).upper()
    for kw in get_config(bank).get("credit_keywords", []):
        if kw.upper() in up:
            return "credit"
    for kw in get_config(bank).get("debit_keywords", []):
        if kw.upper() in up:
            return "debit"
    for kw in BANK_CONFIGS["DEFAULT"]["credit_keywords"]:
        if kw.upper() in up:
            return "credit"
    for kw in BANK_CONFIGS["DEFAULT"]["debit_keywords"]:
        if kw.upper() in up:
            return "debit"
    return "unknown"


def detect_bank(full_text):
    text = clean_text(full_text).upper()
    best_bank = "DEFAULT"
    best_score = 0
    for bank, cfg in BANK_CONFIGS.items():
        if bank == "DEFAULT":
            continue
        score = 0
        for kw in cfg.get("bank_keywords", []):
            if kw.upper() in text:
                score += 10
        for kw in cfg.get("credit_keywords", []):
            if kw.upper() in text:
                score += 1
        for kw in cfg.get("debit_keywords", []):
            if kw.upper() in text:
                score += 1
        if score > best_score:
            best_score = score
            best_bank = bank
    return best_bank


def is_valid_number(x):
    if x is None:
        return False
    try:
        return not math.isnan(float(x))
    except Exception:
        return False


def remove_money_from_desc(text):
    desc = str(text)
    money_patterns = [
        r"(?<!\d)-?\d{1,3}(?:,\d{3})+\.\d{2}(?!\d)",
        r"(?<!\d)-?\d{1,3}(?:\.\d{3})+\.\d{2}(?!\d)",
        r"(?<!\d)-?\d{1,3}(?:\.\d{3})+,\d{2}(?!\d)",
        r"(?<!\d)-?\d+\.\d{2}(?!\d)",
        r"(?<!\d)-?\d+,\d{2}(?!\d)",
    ]
    for pat in money_patterns:
        desc = re.sub(pat, " ", desc)
    desc = re.sub(r"\b\d{1,2}:\d{2}(?::\d{2})?\b", " ", desc)
    desc = re.sub(r"\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b", " ", desc)
    return clean_text(desc)


# ============================================================
# 5) ACCOUNT INFO
# ============================================================
def extract_account_info_from_text(raw_text):
    owner_name = ""
    account_no = ""
    lines = [clean_text(x) for x in raw_text.splitlines() if clean_text(x)]
    for i, line in enumerate(lines[:50]):
        if not owner_name:
            if re.search(r"(ชื่อบัญชี|Account Name|ชื่อ\s*-\s*นามสกุล|ชื่อ\s*สกุล)", line, re.IGNORECASE):
                owner_name = re.sub(r"^(.*?)(ชื่อบัญชี|Account\s*Name|ชื่อ\s*-\s*นามสกุล|ชื่อ\s*สกุล)\s*[:：]?\s*", "", line, flags=re.IGNORECASE).strip()
                if not owner_name and i + 1 < len(lines):
                    owner_name = lines[i + 1].strip()
            elif re.match(r"^(นาย|นาง|น\.ส\.|นางสาว|MR\.|MRS\.|MS\.)\s+", line, re.IGNORECASE):
                owner_name = line
        if not account_no:
            m = re.search(r"(\d{3}[- ]?\d{1}[- ]?\d{5}[- ]?\d{1}|\d{10,12})", line)
            if m:
                account_no = m.group(1).replace(" ", "")
    return owner_name if owner_name else "ไม่ระบุ", account_no if account_no else "ไม่ระบุ"


# ============================================================
# 6) PARSE AND BALANCE CHAIN
# ============================================================
def parse_document_ai_text_to_blocks(raw_text, active_bank, filename):
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    rows = []
    current = None
    for line in lines:
        date = extract_date_from_line(line)
        if date:
            if current:
                rows.append(current)
            current = {"date": date, "time": extract_time_from_line(line), "raw_lines": [line]}
        else:
            if current:
                current["raw_lines"].append(line)
    if current:
        rows.append(current)

    parsed = []
    for idx, item in enumerate(rows, start=1):
        text_block = " ".join(item["raw_lines"])
        money_vals = extract_money_values(text_block)
        is_opening = has_opening_text(text_block, active_bank)
        amount = None
        balance = None
        if is_opening:
            if money_vals:
                balance = money_vals[-1]
        else:
            if len(money_vals) >= 2:
                amount = money_vals[0]
                balance = money_vals[-1]
            elif len(money_vals) == 1:
                amount = money_vals[0]
        parsed.append({
            "seq": idx,
            "วันที่เดือนปี": item.get("date"),
            "เวลา": item.get("time", ""),
            "amount_read": amount,
            "ยอดคงเหลือ": balance,
            "money_tokens": " | ".join(f"{v:.2f}" for v in money_vals),
            "raw_line_text": text_block,
            "is_opening_balance": is_opening,
            "ไฟล์ต้นฉบับ": filename,
        })
    return pd.DataFrame(parsed)


def amount_balance_match(prev_balance, amount, balance):
    if not all(is_valid_number(v) for v in [prev_balance, amount, balance]):
        return None
    prev_balance = float(prev_balance)
    amount = float(amount)
    balance = float(balance)
    debit_expected = round(prev_balance - amount, 2)
    credit_expected = round(prev_balance + amount, 2)
    if abs(balance - debit_expected) <= BALANCE_TOLERANCE:
        return {"type": "debit", "expected_balance": debit_expected, "diff": round(balance - debit_expected, 2), "balance_check": "OK_DEBIT"}
    if abs(balance - credit_expected) <= BALANCE_TOLERANCE:
        return {"type": "credit", "expected_balance": credit_expected, "diff": round(balance - credit_expected, 2), "balance_check": "OK_CREDIT"}
    return None


def suggest_from_balance(prev_balance, balance):
    if not all(is_valid_number(v) for v in [prev_balance, balance]):
        return None, None
    diff = round(float(balance) - float(prev_balance), 2)
    if diff > 0:
        return abs(diff), "credit"
    if diff < 0:
        return abs(diff), "debit"
    return 0.0, "zero"


def build_transaction_dfs(parsed_blocks, active_bank):
    check_rows = []
    tx_rows = []
    prev_balance = None
    seq = 1
    for _, row in parsed_blocks.iterrows():
        date = row.get("วันที่เดือนปี", "")
        time = row.get("เวลา", "")
        amount = row.get("amount_read")
        balance = row.get("ยอดคงเหลือ")
        raw = row.get("raw_line_text", "")
        money_tokens = row.get("money_tokens", "")
        is_opening = bool(row.get("is_opening_balance", False))
        filename = row.get("ไฟล์ต้นฉบับ", "")
        debit = 0.0
        credit = 0.0
        expected_balance = None
        diff_py = None
        suggested_amount = None
        suggested_type = None
        balance_check = ""
        if is_opening or has_opening_text(raw, active_bank):
            balance_check = "OPENING_BALANCE"
        else:
            match = amount_balance_match(prev_balance, amount, balance)
            if match:
                expected_balance = match["expected_balance"]
                diff_py = match["diff"]
                balance_check = match["balance_check"]
                if match["type"] == "debit":
                    debit = float(amount)
                else:
                    credit = float(amount)
            else:
                s_amount, s_type = suggest_from_balance(prev_balance, balance)
                if s_amount is not None and s_type in ["debit", "credit"]:
                    suggested_amount = s_amount
                    suggested_type = s_type
                    expected_balance = balance
                    diff_py = 0.0
                    kw_type = classify_by_keyword(raw, active_bank)
                    balance_check = "SUGGEST_BY_BALANCE_AND_KEYWORD" if kw_type == s_type else "SUGGEST_BY_BALANCE_REVIEW"
                    if s_type == "debit":
                        debit = float(s_amount)
                    else:
                        credit = float(s_amount)
                else:
                    balance_check = "CHAIN_BROKEN_REVIEW"
        desc = remove_money_from_desc(raw)
        check_rows.append({
            "seq": seq,
            "วันที่เดือนปี": date,
            "เวลา": time,
            "เดบิต": debit,
            "เครดิต": credit,
            "ยอดคงเหลือ": balance if is_valid_number(balance) else 0.0,
            "prev_balance": prev_balance,
            "amount_read": amount,
            "suggested_amount": suggested_amount,
            "suggested_type": suggested_type,
            "expected_balance_py": expected_balance,
            "diff_py": diff_py,
            "money_tokens": money_tokens,
            "raw_line_text": raw,
            "balance_check": balance_check,
        })
        if not (is_opening or balance_check == "OPENING_BALANCE") and (debit > 0 or credit > 0):
            tx_rows.append({
                "ลำดับ": len(tx_rows) + 1,
                "วันที่เดือนปี": date,
                "เวลา": time,
                "รายการ": desc[:50],
                "เดบิต": debit,
                "เครดิต": credit,
                "ยอดคงเหลือ": balance if is_valid_number(balance) else 0.0,
                "รายละเอียด": desc,
                "ช่องทาง": "Document AI",
                "หน้า": 1,
                "ไฟล์ต้นฉบับ": filename,
            })
        if is_valid_number(balance):
            prev_balance = float(balance)
        seq += 1
    return pd.DataFrame(tx_rows), pd.DataFrame(check_rows)


# ============================================================
# 7) SALARY / SUMMARY / BANK STATEMENT 1
# ============================================================
def detect_salary_from_df(df_txn):
    if df_txn.empty:
        return pd.DataFrame()
    salary_kws = ["SALARY", "PAYROLL", "SALA", "PAYR", "เงินเดือน", "เงินเดือนหลัก", "SAL"]
    bonus_kws = ["BONUS", "INCENTIVE", "COMMISSION", "โบนัส", "อินเซนทีฟ", "ค่าคอม"]
    rows = []
    for _, row in df_txn.iterrows():
        credit = float(row.get("เครดิต", 0) or 0)
        if credit > 0:
            desc = str(row.get("รายละเอียด", "")).upper()
            if any(kw.upper() in desc for kw in bonus_kws):
                rows.append({"กลุ่ม": "รายได้พิเศษ/โบนัส", "วันที่": row["วันที่เดือนปี"], "จำนวนเงิน": credit})
            elif any(kw.upper() in desc for kw in salary_kws):
                rows.append({"กลุ่ม": "เงินเดือน", "วันที่": row["วันที่เดือนปี"], "จำนวนเงิน": credit})
    return pd.DataFrame(rows)


def build_summary_df(df_txn, owner, acc_no, bank_name):
    if df_txn.empty:
        last_bal, s_date, e_date = 0.0, "", ""
    else:
        valid = pd.to_numeric(df_txn["ยอดคงเหลือ"], errors="coerce").dropna()
        last_bal = float(valid.iloc[-1]) if len(valid) else 0.0
        s_date = df_txn["วันที่เดือนปี"].iloc[0]
        e_date = df_txn["วันที่เดือนปี"].iloc[-1]
    rows = [
        {"รายการ": "เจ้าของบัญชี", "ข้อมูล": owner},
        {"รายการ": "เลขที่บัญชีเงินฝาก", "ข้อมูล": acc_no},
        {"รายการ": "ธนาคาร", "ข้อมูล": bank_name},
        {"รายการ": "จำนวนรายการ", "ข้อมูล": len(df_txn)},
        {"รายการ": "รวมเดบิต", "ข้อมูล": float(df_txn["เดบิต"].sum()) if not df_txn.empty else 0.0},
        {"รายการ": "รวมเครดิต", "ข้อมูล": float(df_txn["เครดิต"].sum()) if not df_txn.empty else 0.0},
        {"รายการ": "ยอดคงเหลือล่าสุด", "ข้อมูล": last_bal},
        {"รายการ": "วันที่เริ่มต้น", "ข้อมูล": s_date},
        {"รายการ": "วันที่สิ้นสุด", "ข้อมูล": e_date},
    ]
    df_income = detect_salary_from_df(df_txn)
    for group_name in ["เงินเดือน", "รายได้พิเศษ/โบนัส"]:
        rows.append({"รายการ": "หมวดหมู่รายได้", "ข้อมูล": group_name})
        if not df_income.empty and group_name in df_income["กลุ่ม"].values:
            g = df_income[df_income["กลุ่ม"] == group_name].copy()
            total = float(g["จำนวนเงิน"].sum())
            rows += [
                {"รายการ": f"{group_name} - สถานะ", "ข้อมูล": "พบ"},
                {"รายการ": f"{group_name} - จำนวนรายการ", "ข้อมูล": len(g)},
                {"รายการ": f"{group_name} - ยอดรวม", "ข้อมูล": total},
                {"รายการ": f"{group_name} - ค่าเฉลี่ยต่อเดือน", "ข้อมูล": total / len(g)},
            ]
            for idx, r in enumerate(g.itertuples(), start=1):
                rows.append({"รายการ": f"{group_name} - รายการที่ {idx}", "ข้อมูล": r.วันที่})
                rows.append({"รายการ": f"{group_name} - ยอด", "ข้อมูล": r.จำนวนเงิน})
        else:
            rows += [
                {"รายการ": f"{group_name} - สถานะ", "ข้อมูล": "ไม่พบ"},
                {"รายการ": f"{group_name} - จำนวนรายการ", "ข้อมูล": 0},
                {"รายการ": f"{group_name} - ยอดรวม", "ข้อมูล": 0.0},
            ]
        rows.append({"รายการ": "", "ข้อมูล": ""})
    rows.append({"รายการ": "สร้างไฟล์เมื่อ", "ข้อมูล": datetime.now().strftime("%d/%m/%Y %H:%M:%S")})
    return pd.DataFrame(rows)


def calculate_30day_statement_logic(df_txn, period_days=30, periods=3):
    empty = {"daily": pd.DataFrame(columns=["วันที่", "ยอดคงเหลือสิ้นวัน", "ช่วงที่"]), "blocks": []}
    if df_txn.empty:
        return empty
    work = df_txn.copy()
    work["_date"] = pd.to_datetime(work["วันที่เดือนปี"].astype(str).str.replace("-", "/"), dayfirst=True, errors="coerce")
    work = work.dropna(subset=["_date"])
    if work.empty:
        return empty
    work["ยอดคงเหลือ"] = pd.to_numeric(work["ยอดคงเหลือ"], errors="coerce").fillna(0.0)
    work["_row_order"] = range(len(work))
    work = work.sort_values(["_date", "_row_order"])
    work["_day"] = work["_date"].dt.normalize()
    daily_txn = work.groupby("_day").agg(**{"ยอดคงเหลือสิ้นวัน": ("ยอดคงเหลือ", "last")}).sort_index()
    full_days = pd.date_range(start=daily_txn.index.min(), end=daily_txn.index.max(), freq="D")
    daily = daily_txn.reindex(full_days)
    daily["ยอดคงเหลือสิ้นวัน"] = daily["ยอดคงเหลือสิ้นวัน"].ffill().fillna(0.0)
    daily = daily.tail(period_days * periods).reset_index().rename(columns={"index": "วันที่"})
    daily["ลำดับ"] = range(1, len(daily) + 1)
    daily["ช่วงที่"] = ((daily["ลำดับ"] - 1) // period_days) + 1
    blocks = []
    for block_no in range(1, periods + 1):
        block = daily[daily["ช่วงที่"] == block_no].copy()
        blocks.append({"ช่วงที่": block_no, "จำนวนวัน": len(block), "ผลรวม": float(block["ยอดคงเหลือสิ้นวัน"].sum()) if len(block) else 0.0, "ค่าเฉลี่ย": float(block["ยอดคงเหลือสิ้นวัน"].mean()) if len(block) else 0.0, "data": block})
    return {"daily": daily, "blocks": blocks}


def add_bank_statement_logic_sheet(path, df_txn, owner, acc_no, bank_name):
    result = calculate_30day_statement_logic(df_txn)
    blocks = result["blocks"]
    wb = load_workbook(path)
    ws = wb.create_sheet("BANK STATEMENT 1")
    title_blue = PatternFill("solid", fgColor="6FA1E8")
    header_blue = PatternFill("solid", fgColor="A9C4F5")
    light_blue = PatternFill("solid", fgColor="DDEBF7")
    yellow = PatternFill("solid", fgColor="FFFF00")
    gray = PatternFill("solid", fgColor="D9D9D9")
    heavy_border = Border(left=Side(style="medium", color="000000"), right=Side(style="medium", color="000000"), top=Side(style="medium", color="000000"), bottom=Side(style="medium", color="000000"))
    data_border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"), top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))
    ws.merge_cells("A1:I1")
    ws["A1"] = "BANK STATEMENT 1"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A1"].fill = title_blue
    for idx in range(3):
        r = 2 + idx
        ws.cell(row=r, column=1, value=1)
        ws.cell(row=r, column=2, value=f"เดือนที่ {idx + 1}")
        ws.cell(row=r, column=3, value=f"=ROUND({['C37','F37','I37'][idx]},2)")
    ws.cell(row=2, column=5, value="เฉลี่ยต่อเดือน")
    ws.cell(row=2, column=6, value="=IF((C39>0)+(F39>0)+(I39>0)=0,0,ROUND((IF(C39>0,C37,0)+IF(F39>0,F37,0)+IF(I39>0,I37,0))/((C39>0)+(F39>0)+(I39>0)),2))")
    ws.cell(row=3, column=5, value="จำนวนวันรวม")
    ws.cell(row=3, column=6, value="=SUM(C39,F39,I39)")
    ws.cell(row=4, column=5, value="ประเมินวงเงินได้")
    ws.cell(row=4, column=6, value="=ROUND(F2*G3*G4,2)")
    ws.merge_cells("H2:I2")
    ws["H2"] = "Account Detail"
    ws.cell(row=3, column=8, value="ธนาคาร")
    ws.cell(row=3, column=9, value="เลขที่บัญชี")
    ws.cell(row=4, column=8, value=bank_name)
    ws.cell(row=4, column=9, value=acc_no)
    for r in range(2, 5):
        for c in range(1, 10):
            cell = ws.cell(row=r, column=c)
            cell.fill = header_blue
            cell.border = heavy_border
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
    for addr in ["C2", "C3", "C4", "F2", "F3"]:
        ws[addr].fill = light_blue
    for idx, block in enumerate(blocks):
        start_col = [1, 4, 7][idx]
        n_col, date_col, amount_col = start_col, start_col + 1, start_col + 2
        for col, value in [(n_col, "N"), (date_col, "DATE"), (amount_col, "AMOUNT")]:
            cell = ws.cell(row=5, column=col, value=value)
            cell.fill = header_blue
            cell.font = Font(bold=True)
            cell.border = heavy_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
        data = block.get("data", pd.DataFrame())
        for i in range(1, 31):
            erow = 6 + i - 1
            ws.cell(row=erow, column=n_col, value=i).fill = light_blue
            ws.cell(row=erow, column=date_col).fill = light_blue
            if i <= len(data):
                row_data = data.iloc[i - 1]
                date_val = row_data.get("วันที่", "")
                amt_val = float(row_data.get("ยอดคงเหลือสิ้นวัน", 0))
                ws.cell(row=erow, column=date_col, value=date_val)
                ws.cell(row=erow, column=amount_col, value=amt_val)
                if i > 1:
                    prev_amt = float(data.iloc[i - 2].get("ยอดคงเหลือสิ้นวัน", 0))
                    if round(amt_val, 2) != round(prev_amt, 2):
                        ws.cell(row=erow, column=date_col).fill = yellow
                        ws.cell(row=erow, column=amount_col).fill = yellow
            for col in [n_col, date_col, amount_col]:
                cell = ws.cell(row=erow, column=col)
                cell.border = data_border
                cell.alignment = Alignment(horizontal="center" if col != amount_col else "right", vertical="center")
            ws.cell(row=erow, column=date_col).number_format = "d mmm yy"
            ws.cell(row=erow, column=amount_col).number_format = "#,##0.00"
        amount_letter = ws.cell(row=5, column=amount_col).column_letter
        ws.cell(row=36, column=n_col, value="รวม")
        ws.cell(row=36, column=amount_col, value=f"=SUM({amount_letter}6:{amount_letter}35)")
        ws.cell(row=37, column=amount_col, value=f"=IF({amount_letter}39=0,0,{amount_letter}36/{amount_letter}39)")
        ws.cell(row=38, column=amount_col, value="จำนวนวัน")
        ws.cell(row=39, column=amount_col, value=f"=COUNT({amount_letter}6:{amount_letter}35)")
        for r in range(36, 40):
            for c in [n_col, date_col, amount_col]:
                cell = ws.cell(row=r, column=c)
                cell.border = heavy_border
                cell.alignment = Alignment(horizontal="center" if c != amount_col else "right", vertical="center")
                if c == amount_col:
                    cell.number_format = "#,##0.00"
        ws.cell(row=38, column=amount_col).fill = gray
        ws.cell(row=38, column=amount_col).font = Font(bold=True)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"
    for letter, width in {"A": 5, "B": 14, "C": 16, "D": 5, "E": 14, "F": 16, "G": 5, "H": 14, "I": 16}.items():
        ws.column_dimensions[letter].width = width
    wb.save(path)


# ============================================================
# 8) EXCEL EXPORT
# ============================================================
def auto_fit_worksheet(ws):
    ws.freeze_panes = "A2"
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    headers = {c: clean_text(ws.cell(row=1, column=c).value) for c in range(1, ws.max_column + 1)}
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
        ws.column_dimensions[letter].width = max(8, min(width + 1.5, 40))


def format_summary_sheet(ws):
    ws.sheet_view.showGridLines = False
    header_fill = PatternFill("solid", fgColor="1F4E79")
    section_fill = PatternFill("solid", fgColor="DDEBF7")
    sub_fill = PatternFill("solid", fgColor="EAF4EA")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for r in range(2, ws.max_row + 1):
        label = clean_text(ws.cell(row=r, column=1).value)
        if label == "หมวดหมู่รายได้":
            for c in range(1, 3):
                ws.cell(row=r, column=c).fill = section_fill
                ws.cell(row=r, column=c).font = Font(bold=True)
        if any(k in label for k in ["สถานะ", "จำนวนรายการ", "ยอดรวม", "ค่าเฉลี่ยต่อเดือน"]):
            ws.cell(row=r, column=1).fill = sub_fill
            ws.cell(row=r, column=2).fill = sub_fill
            ws.cell(row=r, column=1).font = Font(bold=True)
        if isinstance(ws.cell(row=r, column=2).value, (int, float)):
            ws.cell(row=r, column=2).number_format = "#,##0.00"


def format_workbook(path):
    wb = load_workbook(path)
    for ws in wb.worksheets:
        if ws.title != "BANK STATEMENT 1":
            auto_fit_worksheet(ws)
            if ws.title == "สรุปยอด":
                format_summary_sheet(ws)
    wb.save(path)


def build_output_excel(df_account, df_txn, owner, acc_no, bank_name, raw_text, document_json_text, check_df):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp_path = tmp.name
    review_df = pd.DataFrame()
    if check_df is not None and not check_df.empty and "balance_check" in check_df.columns:
        review_df = check_df[check_df["balance_check"].astype(str).str.contains("REVIEW|BROKEN", regex=True, na=False)].copy()
    with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
        df_account.to_excel(writer, sheet_name="เจ้าของบัญชี", index=False)
        df_txn.to_excel(writer, sheet_name="รายละเอียด เดบิต-เครดิต", index=False)
        build_summary_df(df_txn, owner, acc_no, bank_name).to_excel(writer, sheet_name="สรุปยอด", index=False)
        check_df.to_excel(writer, sheet_name="check", index=False)
        review_df.to_excel(writer, sheet_name="ocr_review", index=False)
        pd.DataFrame({"raw_text": [raw_text]}).to_excel(writer, sheet_name="raw_text", index=False)
        pd.DataFrame({"document_json": [document_json_text]}).to_excel(writer, sheet_name="document_json", index=False)
    add_bank_statement_logic_sheet(tmp_path, df_txn, owner, acc_no, bank_name)
    format_workbook(tmp_path)
    with open(tmp_path, "rb") as f:
        data = f.read()
    os.unlink(tmp_path)
    return data


# ============================================================
# 9) PIPELINE + UI
# ============================================================
def process_one_file(uploaded_file, selected_bank):
    document = process_uploaded_file_with_document_ai(uploaded_file)
    raw_text = document.text or ""
    if not raw_text.strip():
        raise ValueError("Document AI อ่านข้อความไม่เจอ")
    document_json_text = document_to_json_text(document)
    active_bank = selected_bank if selected_bank != "AUTO" else detect_bank(raw_text)
    blocks = parse_document_ai_text_to_blocks(raw_text, active_bank, uploaded_file.name)
    df_txn, check_df = build_transaction_dfs(blocks, active_bank)
    owner, acc_no = extract_account_info_from_text(raw_text)
    return {"raw_text": raw_text, "document_json_text": document_json_text, "active_bank": active_bank, "df_txn": df_txn, "check_df": check_df, "owner": owner, "acc_no": acc_no}


def render_section(step_no, title):
    st.markdown(f'<div class="section-title"><span class="step-number">{step_no}</span><span>{title}</span></div>', unsafe_allow_html=True)


def main_app():
    st.title("STM Document AI Parser")
    st.caption("JPG / PNG / PDF → Document AI → Excel 4 ชีทหลัก + สรุปเงินเดือน + BANK STATEMENT 1")
    st.divider()
    render_section("1", "เลือกธนาคาร")
    bank_choice = st.radio("ธนาคาร", options=["AUTO", "KBANK", "BBL", "SCB", "KRUNGSRI"], format_func=lambda x: BANK_LABELS.get(x, x), index=0, horizontal=True, label_visibility="collapsed")
    render_section("2", "ตรวจสอบ / แก้ไขข้อมูลบัญชี")
    st.info("ระบบจะพยายามค้นหา ชื่อบัญชี และ เลขที่บัญชี ให้อัตโนมัติ แต่คุณสามารถกรอกเองได้")
    col1, col2 = st.columns(2)
    owner_input = col1.text_input("ชื่อเจ้าของบัญชี", placeholder="เว้นว่างไว้เพื่อให้ระบบหาอัตโนมัติ")
    acc_input = col2.text_input("เลขที่บัญชี", placeholder="เว้นว่างไว้เพื่อให้ระบบหาอัตโนมัติ")
    render_section("3", "อัปโหลด Statement")
    uploaded_files = st.file_uploader("เลือกไฟล์ Statement JPG / PNG / PDF", type=["jpg", "jpeg", "png", "pdf"], accept_multiple_files=True)
    render_section("4", "ประมวลผล")
    c1, c2 = st.columns(2)
    process_clicked = c1.button("เริ่มประมวลผล", type="primary", use_container_width=True, disabled=not uploaded_files)
    if c2.button("ออกจากระบบ", use_container_width=True):
        st.session_state.authenticated = False
        st.session_state.results = None
        st.rerun()
    if process_clicked:
        st.session_state.results = None
        progress = st.progress(0, text="เริ่มประมวลผล...")
        status = st.empty()
        try:
            all_txn, all_check, all_raw, all_json, banks = [], [], [], [], []
            auto_owner, auto_acc = "", ""
            total = len(uploaded_files)
            for idx, uf in enumerate(uploaded_files, start=1):
                status.info(f"กำลังประมวลผลไฟล์ {idx}/{total}: {uf.name}")
                progress.progress(int((idx - 1) / total * 70))
                res = process_one_file(uf, bank_choice)
                if not res["df_txn"].empty:
                    all_txn.append(res["df_txn"])
                if not res["check_df"].empty:
                    temp_check = res["check_df"].copy()
                    temp_check["ไฟล์ต้นฉบับ"] = uf.name
                    all_check.append(temp_check)
                all_raw.append(f"===== {uf.name} =====\n{res['raw_text']}")
                all_json.append(f"===== {uf.name} =====\n{res['document_json_text']}")
                banks.append(res["active_bank"])
                if (not auto_owner or auto_owner == "ไม่ระบุ") and res["owner"] != "ไม่ระบุ":
                    auto_owner = res["owner"]
                if (not auto_acc or auto_acc == "ไม่ระบุ") and res["acc_no"] != "ไม่ระบุ":
                    auto_acc = res["acc_no"]
            progress.progress(80, text="กำลังสร้าง Excel...")
            final_df = pd.concat(all_txn, ignore_index=True) if all_txn else pd.DataFrame(columns=["ลำดับ", "วันที่เดือนปี", "เวลา", "รายการ", "เดบิต", "เครดิต", "ยอดคงเหลือ", "รายละเอียด", "ช่องทาง", "หน้า", "ไฟล์ต้นฉบับ"])
            final_df["ลำดับ"] = range(1, len(final_df) + 1)
            check_df = pd.concat(all_check, ignore_index=True) if all_check else pd.DataFrame()
            owner = owner_input.strip() if owner_input.strip() else (auto_owner if auto_owner else "ไม่ระบุ")
            acc_no = acc_input.strip() if acc_input.strip() else (auto_acc if auto_acc else "ไม่ระบุ")
            active_bank = banks[0] if bank_choice == "AUTO" and banks else bank_choice
            bank_name = BANK_LABELS.get(active_bank, active_bank)
            df_account = pd.DataFrame([{"เจ้าของบัญชี": owner, "เลขที่บัญชีเงินฝาก": acc_no, "ธนาคาร": bank_name}])
            excel_bytes = build_output_excel(df_account, final_df, owner, acc_no, bank_name, "\n\n".join(all_raw), "\n\n".join(all_json), check_df)
            excel_name = f"Statement_{owner}.xlsx".replace("/", "_").replace("\\", "_")
            st.session_state.results = {"df": final_df, "check": check_df, "excel": excel_bytes, "excel_name": excel_name, "owner": owner, "acc_no": acc_no, "bank_name": bank_name}
            progress.progress(100, text="เสร็จสิ้น ✅")
            status.success(f"ประมวลผลสำเร็จ พบ {len(final_df):,} รายการ")
        except Exception as e:
            progress.empty()
            status.error(f"ประมวลผลไม่สำเร็จ: {e}")
            st.exception(e)
    if st.session_state.results:
        res = st.session_state.results
        df = res["df"]
        check_df = res["check"]
        st.divider()
        st.success(f"ประมวลผลสำเร็จ! พบ {len(df):,} รายการ")
        debit_sum = float(pd.to_numeric(df.get("เดบิต", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
        credit_sum = float(pd.to_numeric(df.get("เครดิต", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
        last_balance = 0.0
        if not df.empty:
            bal = pd.to_numeric(df.get("ยอดคงเหลือ", pd.Series(dtype=float)), errors="coerce").dropna()
            if len(bal):
                last_balance = float(bal.iloc[-1])
        review_count = int(check_df["balance_check"].astype(str).str.contains("REVIEW|BROKEN", regex=True, na=False).sum()) if check_df is not None and not check_df.empty and "balance_check" in check_df.columns else 0
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("เจ้าของบัญชี", res["owner"])
        m2.metric("จำนวนรายการ", f"{len(df):,}")
        m3.metric("รวมเดบิต", f"{debit_sum:,.2f}")
        m4.metric("รวมเครดิต", f"{credit_sum:,.2f}")
        m5, m6, m7 = st.columns(3)
        m5.metric("ยอดคงเหลือล่าสุด", f"{last_balance:,.2f}")
        m6.metric("ต้องตรวจ", f"{review_count:,}")
        m7.metric("ธนาคาร", res["bank_name"])
        st.download_button("📥 ดาวน์โหลดไฟล์ Excel", data=res["excel"], file_name=res["excel_name"], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")
        st.markdown("#### ตัวอย่างข้อมูลที่ประมวลผลได้")
        if df.empty:
            st.warning("ไม่พบรายการธุรกรรม")
        else:
            st.dataframe(df.head(30).style.format({"เดบิต": "{:,.2f}", "เครดิต": "{:,.2f}", "ยอดคงเหลือ": "{:,.2f}"}), use_container_width=True, height=400)
        tab1, tab2 = st.tabs(["รายละเอียด เดบิต-เครดิต", "check"])
        with tab1:
            st.dataframe(df, use_container_width=True, height=520)
        with tab2:
            st.dataframe(check_df, use_container_width=True, height=520)
        if st.button("ล้างข้อมูลหลังใช้งาน", use_container_width=True):
            st.session_state.results = None
            st.rerun()


if not st.session_state.authenticated:
    login_page()
else:
    main_app()

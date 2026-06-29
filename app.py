# ============================================================
# STM Document AI to Excel Parser — Streamlit Edition
# JPG / PNG / PDF -> Google Document AI -> Excel
# ============================================================

import io
import os
import re
import json
import math
from datetime import datetime

import streamlit as st
import pandas as pd

from google.api_core.client_options import ClientOptions
from google.cloud import documentai_v1 as documentai
from google.oauth2 import service_account
from google.protobuf.json_format import MessageToDict

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter


# ============================================================
# 0) Page Config
# ============================================================
st.set_page_config(
    page_title="STM Document AI Parser",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
.block-container {
    padding-top: 2rem;
    padding-bottom: 3rem;
    max-width: 980px;
}
h1 {
    font-size: 2.2rem !important;
    font-weight: 800 !important;
    letter-spacing: -0.5px;
}
.section-title {
    font-weight: 800;
    font-size: 1.1rem;
    margin: 18px 0 8px 0;
}
div[data-testid="stButton"] button,
div[data-testid="stDownloadButton"] button {
    border-radius: 10px;
    font-weight: 750;
}
div[data-testid="stAlert"] {
    border-radius: 10px;
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# 1) Constants
# ============================================================
BALANCE_TOLERANCE = 0.05
LOW_CONF_THRESHOLD = 0.75

BANK_LABELS = {
    "AUTO": "ตรวจจับอัตโนมัติ",
    "KBANK": "กสิกรไทย",
    "BBL": "กรุงเทพ",
    "SCB": "ไทยพาณิชย์",
    "KRUNGSRI": "กรุงศรี",
}

BANK_CONFIGS = {
    "DEFAULT": {
        "bank_keywords": [],
        "opening_keywords": [
            "ยอด ยก มา", "ยอดยกมา", "ยอด ยก", "ยก มา",
            "BALANCE BROUGHT FORWARD", "BROUGHT FORWARD",
            "BALANCE B/F", "B/F"
        ],
        "credit_keywords": [
            "รับ โอน", "รับโอน", "ฝาก", "ฝาก เงินสด", "เงิน เข้า", "เงินเข้า",
            "ดอกเบี้ย", "คืนเงิน", "CREDIT", "DEPOSIT", "TRANSFER IN",
            "Internet/Mobile", "รับจาก"
        ],
        "debit_keywords": [
            "โอน เงิน", "โอนเงิน", "ถอน", "ถอนเงิน", "หัก", "ชำระ",
            "จ่าย", "ค่าธรรมเนียม", "DEBIT", "WITHDRAW", "PAYMENT",
            "FEE", "TRANSFER OUT", "ATM"
        ],
    },
    "KBANK": {
        "bank_keywords": ["KASIKORN", "KBANK", "K PLUS", "กสิกร"],
        "opening_keywords": ["ยอด ยก มา", "ยอดยกมา", "BROUGHT FORWARD"],
        "credit_keywords": ["รับ โอน", "รับโอน", "ฝาก", "เงิน เข้า", "CREDIT", "DEPOSIT"],
        "debit_keywords": ["โอน เงิน", "โอนเงิน", "ถอน", "จ่าย", "PAYMENT", "WITHDRAW", "FEE"],
    },
    "BBL": {
        "bank_keywords": ["BANGKOK BANK", "BANGKOKBANK", "ธนาคารกรุงเทพ", "BBL"],
        "opening_keywords": ["B/F", "BALANCE B/F", "BROUGHT FORWARD", "ยอดยกมา"],
        "credit_keywords": ["TRF FR", "TRANSFER IN", "CREDIT", "DEPOSIT", "รับโอน", "ฝาก"],
        "debit_keywords": ["TRF TO", "PMT", "CASH W/D", "WITHDRAW", "PAYMENT", "DEBIT", "FEE"],
    },
    "SCB": {
        "bank_keywords": ["SCB", "ไทยพาณิชย์", "SIAM COMMERCIAL BANK"],
        "opening_keywords": ["ยอดเงินคงเหลือยกมา", "BROUGHT FORWARD", "B/F"],
        "credit_keywords": ["รับ โอน", "รับโอน", "ฝาก", "เงิน เข้า", "CREDIT", "DEPOSIT"],
        "debit_keywords": ["โอน ไป", "โอนไป", "จ่าย", "ถอน", "PAYMENT", "WITHDRAW", "FEE"],
    },
    "KRUNGSRI": {
        "bank_keywords": ["KRUNGSRI", "กรุงศรี", "AYUDHYA"],
        "opening_keywords": ["ยอดยกมา", "BROUGHT FORWARD", "B/F"],
        "credit_keywords": ["รับ โอน", "รับโอน", "ฝาก", "เงิน เข้า", "CREDIT", "DEPOSIT"],
        "debit_keywords": ["โอน เงิน", "ถอน", "จ่าย", "PAYMENT", "WITHDRAW", "FEE"],
    },
}


# ============================================================
# 2) Login
# ============================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if "results" not in st.session_state:
    st.session_state.results = None


def login_page():
    st.title("Login")
    st.caption("กรุณาเข้าสู่ระบบก่อนใช้งาน")

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
# 3) Google Document AI Client
# ============================================================
def get_documentai_config():
    try:
        project_id = st.secrets["DOCUMENTAI_PROJECT_ID"]
        location = st.secrets["DOCUMENTAI_LOCATION"]
        processor_id = st.secrets["DOCUMENTAI_PROCESSOR_ID"]
        return project_id, location, processor_id
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

    opts = ClientOptions(
        api_endpoint=f"{location}-documentai.googleapis.com"
    )

    client = documentai.DocumentProcessorServiceClient(
        credentials=creds,
        client_options=opts,
    )

    processor_name = client.processor_path(
        project_id,
        location,
        processor_id,
    )

    return client, processor_name


def get_mime_type(filename: str) -> str:
    filename = filename.lower()

    if filename.endswith(".pdf"):
        return "application/pdf"
    if filename.endswith(".jpg") or filename.endswith(".jpeg"):
        return "image/jpeg"
    if filename.endswith(".png"):
        return "image/png"

    raise ValueError("รองรับเฉพาะ PDF, JPG, JPEG, PNG")


def process_uploaded_file_with_document_ai(uploaded_file):
    client, processor_name = get_documentai_client()

    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)

    mime_type = get_mime_type(uploaded_file.name)

    raw_document = documentai.RawDocument(
        content=file_bytes,
        mime_type=mime_type,
    )

    request = documentai.ProcessRequest(
        name=processor_name,
        raw_document=raw_document,
    )

    result = client.process_document(request=request)

    return result.document


def document_to_dict(document):
    return MessageToDict(
        document._pb,
        preserving_proto_field_name=True,
    )


def document_to_json_bytes(document):
    document_dict = document_to_dict(document)

    json_text = json.dumps(
        document_dict,
        ensure_ascii=False,
        indent=2,
    )

    return json_text.encode("utf-8")


# ============================================================
# 4) Text Helpers
# ============================================================
def clean_text(s: str) -> str:
    s = str(s).replace("\n", " ")
    return re.sub(r"\s+", " ", s).strip()


def normalize_amount_text(s: str) -> str:
    s = str(s).strip()
    s = re.sub(r"\s+", "", s)

    ocr_map = {
        "O": "0", "o": "0",
        "I": "1", "l": "1", "|": "1",
        "S": "5", "s": "5",
        "B": "8",
        "g": "9", "q": "9",
        "Z": "2", "z": "2",
    }

    for c, d in ocr_map.items():
        s = s.replace(c, d)

    # OCR อ่าน 25,000.00 เป็น 25.000.00
    if re.match(r"^\d{1,3}\.\d{3}\.\d{2}$", s):
        parts = s.split(".")
        s = parts[0] + "," + parts[1] + "." + parts[2]

    # OCR อ่าน 25.000,00 แบบยุโรป
    if re.match(r"^\d{1,3}\.\d{3},\d{2}$", s):
        s = s.replace(".", "").replace(",", ".")

    # OCR อ่าน 100,00
    if re.match(r"^\d+,\d{2}$", s):
        s = s.replace(",", ".")

    return s


def parse_money(s):
    if s is None:
        return None

    s = normalize_amount_text(str(s))
    s = re.sub(r"[^0-9,.\-]", "", s)

    if not s:
        return None

    try:
        return float(s.replace(",", ""))
    except Exception:
        return None


def extract_money_values(text: str) -> list[float]:
    """
    หาเลขเงินจากทั้งบรรทัด ไม่จำเป็นต้องเป็นตัวเลขล้วน
    เช่น 4,151.01 ATM a -> ได้ 4151.01
    """

    text = str(text)

    ocr_map = {
        "O": "0", "o": "0",
        "I": "1", "l": "1", "|": "1",
        "S": "5", "s": "5",
        "B": "8",
        "g": "9", "q": "9",
        "Z": "2", "z": "2",
    }

    for c, d in ocr_map.items():
        text = text.replace(c, d)

    patterns = [
        r"(?<!\d)-?\d{1,3}(?:,\d{3})+\.\d{2}(?!\d)",
        r"(?<!\d)-?\d{1,3}(?:\.\d{3})+\.\d{2}(?!\d)",
        r"(?<!\d)-?\d{1,3}(?:\.\d{3})+,\d{2}(?!\d)",
        r"(?<!\d)-?\d+\.\d{2}(?!\d)",
        r"(?<!\d)-?\d+,\d{2}(?!\d)",
    ]

    results = []

    for pat in patterns:
        matches = re.findall(pat, text)
        for m in matches:
            value = parse_money(m)
            if value is not None:
                results.append(value)

    # ลบซ้ำแต่รักษาลำดับ
    unique = []
    for v in results:
        if v not in unique:
            unique.append(v)

    return unique


def fix_ocr_date(s: str) -> str:
    s = str(s).strip()
    s = s.replace("O", "0").replace("o", "0")
    s = s.replace("I", "1").replace("l", "1")
    s = s.replace(".", "-").replace("/", "-")
    return s


def normalize_date(s: str):
    s = fix_ocr_date(s)

    m1 = re.match(r"^(\d{1,2})[-](\d{1,2})[-](\d{2,4})$", s)
    m2 = re.match(r"^(\d{1,2})[-](\d{1,2})$", s)

    if m1:
        day = int(m1.group(1))
        month = int(m1.group(2))
        year_raw = m1.group(3)

        if len(year_raw) == 2:
            year = 2000 + int(year_raw)
        else:
            year = int(year_raw)

    elif m2:
        day = int(m2.group(1))
        month = int(m2.group(2))
        year = datetime.now().year
    else:
        return None

    try:
        dt = datetime(year, month, day)
        return dt.strftime("%d-%m-%Y")
    except Exception:
        return None


def extract_date_from_line(line: str):
    tokens = str(line).split()

    for tok in tokens:
        d = normalize_date(tok)
        if d:
            return d

    return None


def extract_time_from_line(line: str):
    m = re.search(r"\b\d{1,2}:\d{2}(?::\d{2})?\b", str(line))
    if m:
        return m.group(0)
    return None


def get_config(bank: str):
    return BANK_CONFIGS.get(bank, BANK_CONFIGS["DEFAULT"])


def has_opening_text(text: str, bank: str = "DEFAULT") -> bool:
    up = clean_text(text).upper()

    for kw in get_config(bank).get("opening_keywords", []):
        if kw.upper() in up:
            return True

    for kw in BANK_CONFIGS["DEFAULT"]["opening_keywords"]:
        if kw.upper() in up:
            return True

    return False


def classify_by_keyword(text: str, bank: str = "DEFAULT") -> str:
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


def detect_bank(full_text: str) -> str:
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


# ============================================================
# 5) Parse Document AI Text
# ============================================================
def parse_document_ai_text(raw_text: str, active_bank: str) -> pd.DataFrame:
    """
    แยกธุรกรรมจาก document.text
    หลัก:
    - เจอวันที่ = เริ่มรายการใหม่
    - เก็บบรรทัดถัดไปจนกว่าจะเจอวันที่ใหม่
    - หาเงินจากทั้งกลุ่มข้อความ
    - เงินตัวแรก = amount
    - เงินตัวท้าย = balance
    """

    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]

    rows = []
    current = None

    for line in lines:
        date = extract_date_from_line(line)

        if date:
            if current:
                rows.append(current)

            current = {
                "date": date,
                "time": extract_time_from_line(line),
                "raw_lines": [line],
            }
        else:
            if current:
                current["raw_lines"].append(line)

    if current:
        rows.append(current)

    parsed_rows = []

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

        parsed_rows.append({
            "seq": idx,
            "date": item.get("date"),
            "time": item.get("time"),
            "amount": amount,
            "balance": balance,
            "money_tokens": " | ".join(f"{v:.2f}" for v in money_vals),
            "raw_line_text": text_block,
            "is_opening_balance": is_opening,
        })

    return pd.DataFrame(parsed_rows)


# ============================================================
# 6) Balance Chain / Debit Credit
# ============================================================
def is_valid_number(x) -> bool:
    if x is None:
        return False

    try:
        x = float(x)
        return not math.isnan(x)
    except Exception:
        return False


def amount_balance_match(prev_balance, amount, balance, tolerance=BALANCE_TOLERANCE):
    if not all(is_valid_number(v) for v in [prev_balance, amount, balance]):
        return None

    prev_balance = float(prev_balance)
    amount = float(amount)
    balance = float(balance)

    debit_expected = round(prev_balance - amount, 2)
    credit_expected = round(prev_balance + amount, 2)

    if abs(balance - debit_expected) <= tolerance:
        return {
            "type": "debit",
            "expected_balance": debit_expected,
            "diff": round(balance - debit_expected, 2),
            "balance_check": "OK_DEBIT",
        }

    if abs(balance - credit_expected) <= tolerance:
        return {
            "type": "credit",
            "expected_balance": credit_expected,
            "diff": round(balance - credit_expected, 2),
            "balance_check": "OK_CREDIT",
        }

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


def add_debit_credit_check(parsed_lines: pd.DataFrame, active_bank: str):
    check_rows = []
    prev_balance = None

    for _, row in parsed_lines.iterrows():
        seq = row.get("seq")
        date = row.get("date")
        time = row.get("time")
        amount = row.get("amount")
        balance = row.get("balance")
        raw = row.get("raw_line_text", "")
        money_tokens = row.get("money_tokens", "")
        is_opening = row.get("is_opening_balance", False)

        debit = None
        credit = None
        expected_balance = None
        diff = None
        suggested_amount = None
        suggested_type = None
        balance_check = ""

        if is_opening or has_opening_text(raw, active_bank):
            balance_check = "OPENING_BALANCE"

        else:
            match = amount_balance_match(prev_balance, amount, balance)

            if match:
                expected_balance = match["expected_balance"]
                diff = match["diff"]
                balance_check = match["balance_check"]

                if match["type"] == "debit":
                    debit = amount
                else:
                    credit = amount

            else:
                # ถ้า OCR อ่าน amount ไม่ดี แต่ balance chain บอกได้
                s_amount, s_type = suggest_from_balance(prev_balance, balance)

                if s_amount is not None and s_type in ["debit", "credit"]:
                    suggested_amount = s_amount
                    suggested_type = s_type
                    expected_balance = balance
                    diff = 0.0

                    kw_type = classify_by_keyword(raw, active_bank)

                    if kw_type == s_type:
                        balance_check = "SUGGEST_BY_BALANCE_AND_KEYWORD"
                    else:
                        balance_check = "SUGGEST_BY_BALANCE_REVIEW"

                    if s_type == "debit":
                        debit = s_amount
                    else:
                        credit = s_amount
                else:
                    balance_check = "CHAIN_BROKEN_REVIEW"

        check_rows.append({
            "seq": seq,
            "date": date,
            "time": time,
            "debit": debit,
            "credit": credit,
            "balance": balance,
            "prev_balance": prev_balance,
            "amount_read": amount,
            "suggested_amount": suggested_amount,
            "suggested_type": suggested_type,
            "expected_balance_py": expected_balance,
            "diff_py": diff,
            "money_tokens": money_tokens,
            "raw_line_text": raw,
            "balance_check": balance_check,
        })

        if is_valid_number(balance):
            prev_balance = float(balance)

    check_df = pd.DataFrame(check_rows)

    parsed_df = check_df.copy()

    parsed_df = parsed_df[
        ~parsed_df["balance_check"].astype(str).str.contains("OPENING_BALANCE", na=False)
    ].copy()

    has_amount = (
        parsed_df["debit"].apply(is_valid_number)
        | parsed_df["credit"].apply(is_valid_number)
    )

    parsed_df = parsed_df[has_amount].copy()
    parsed_df = parsed_df[["date", "debit", "credit", "balance"]].reset_index(drop=True)

    return parsed_df, check_df


# ============================================================
# 7) Excel Export
# ============================================================
def create_summary_df(parsed_df: pd.DataFrame, check_df: pd.DataFrame, active_bank: str):
    debit_total = pd.to_numeric(parsed_df.get("debit", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    credit_total = pd.to_numeric(parsed_df.get("credit", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()

    balance_series = pd.to_numeric(parsed_df.get("balance", pd.Series(dtype=float)), errors="coerce").dropna()
    last_balance = float(balance_series.iloc[-1]) if len(balance_series) else 0.0

    ok_count = 0
    review_count = 0

    if not check_df.empty and "balance_check" in check_df.columns:
        ok_count = int(check_df["balance_check"].astype(str).str.contains("OK|SUGGEST", regex=True).sum())
        review_count = int(check_df["balance_check"].astype(str).str.contains("REVIEW|BROKEN", regex=True).sum())

    return pd.DataFrame([
        ["detected_bank", active_bank],
        ["transaction_rows", len(parsed_df)],
        ["debit_total", float(debit_total)],
        ["credit_total", float(credit_total)],
        ["net_change", float(credit_total - debit_total)],
        ["last_balance", last_balance],
        ["ok_or_suggest_checks", ok_count],
        ["review_rows", review_count],
    ], columns=["metric", "value"])


def create_excel(parsed_df, check_df, raw_text, document_json_bytes, active_bank) -> bytes:
    buf = io.BytesIO()

    summary_df = create_summary_df(parsed_df, check_df, active_bank)
    raw_df = pd.DataFrame({"raw_text": [raw_text]})

    review_df = check_df[
        check_df["balance_check"].astype(str).str.contains("REVIEW|BROKEN", regex=True, na=False)
    ].copy()

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        parsed_df.to_excel(writer, sheet_name="parsed_stm", index=False)
        check_df.to_excel(writer, sheet_name="check", index=False)
        summary_df.to_excel(writer, sheet_name="summary", index=False)
        review_df.to_excel(writer, sheet_name="ocr_review", index=False)
        raw_df.to_excel(writer, sheet_name="raw_text", index=False)

    buf.seek(0)
    wb = load_workbook(buf)

    header_fill = PatternFill("solid", fgColor="D9EAF7")
    header_font = Font(color="000000", bold=True)
    thin = Side(style="thin", color="D9E2F3")

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(bottom=thin)

        ws.freeze_panes = "A2"

        headers = {str(cell.value): cell.column for cell in ws[1] if cell.value}

        for col_name in [
            "debit", "credit", "balance", "prev_balance", "amount_read",
            "suggested_amount", "expected_balance_py", "diff_py"
        ]:
            if col_name in headers:
                col_idx = headers[col_name]
                for row in range(2, ws.max_row + 1):
                    ws.cell(row=row, column=col_idx).number_format = "#,##0.00"

        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 0

            for cell in ws[col_letter]:
                value = str(cell.value) if cell.value is not None else ""
                max_len = max(max_len, len(value))

            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 55)

    # Conditional formatting เฉพาะ sheet check
    if "check" in wb.sheetnames:
        ws = wb["check"]
        headers = {str(cell.value): cell.column for cell in ws[1] if cell.value}

        if "balance_check" in headers and ws.max_row >= 2:
            reason_col = get_column_letter(headers["balance_check"])
            red_fill = PatternFill("solid", fgColor="FCE4E4")
            red_font = Font(color="9C0006")

            ws.conditional_formatting.add(
                f"A2:{get_column_letter(ws.max_column)}{ws.max_row}",
                FormulaRule(
                    formula=[f'=ISNUMBER(SEARCH("REVIEW",${reason_col}2))'],
                    fill=red_fill,
                    font=red_font,
                ),
            )

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ============================================================
# 8) Main App
# ============================================================
def main_app():
    st.title("STM Document AI Parser")
    st.caption("JPG / PNG / PDF → Google Document AI → Excel")
    st.divider()

    st.markdown('<div class="section-title">1. เลือกธนาคาร</div>', unsafe_allow_html=True)

    bank_choice = st.radio(
        "ธนาคาร",
        options=list(BANK_LABELS.keys()),
        format_func=lambda x: BANK_LABELS[x],
        index=0,
        horizontal=True,
        label_visibility="collapsed",
    )

    st.markdown('<div class="section-title">2. อัปโหลด Statement</div>', unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "เลือกไฟล์ JPG / PNG / PDF",
        type=["jpg", "jpeg", "png", "pdf"],
        accept_multiple_files=False,
    )

    col_a, col_b = st.columns(2)

    with col_a:
        process_clicked = st.button(
            "เริ่มอ่านด้วย Document AI",
            type="primary",
            use_container_width=True,
            disabled=uploaded_file is None,
        )

    with col_b:
        if st.button("ออกจากระบบ", use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.results = None
            st.rerun()

    if process_clicked and uploaded_file:
        st.session_state.results = None

        progress = st.progress(0, text="เริ่มประมวลผล...")
        status = st.empty()

        try:
            status.info("กำลังส่งไฟล์เข้า Google Document AI...")
            progress.progress(20, text="20%")

            document = process_uploaded_file_with_document_ai(uploaded_file)
            raw_text = document.text or ""

            if not raw_text.strip():
                st.error("Document AI อ่านข้อความไม่เจอ")
                st.stop()

            document_json_bytes = document_to_json_bytes(document)

            progress.progress(45, text="45%")
            status.info("กำลังตรวจจับธนาคาร...")

            active_bank = bank_choice
            if bank_choice == "AUTO":
                active_bank = detect_bank(raw_text)

            progress.progress(60, text="60%")
            status.info("กำลังแยกรายการธุรกรรม...")

            parsed_lines = parse_document_ai_text(raw_text, active_bank)

            progress.progress(75, text="75%")
            status.info("กำลังคำนวณ debit / credit จาก balance chain...")

            parsed_df, check_df = add_debit_credit_check(parsed_lines, active_bank)

            progress.progress(90, text="90%")
            status.info("กำลังสร้าง Excel...")

            excel_bytes = create_excel(
                parsed_df=parsed_df,
                check_df=check_df,
                raw_text=raw_text,
                document_json_bytes=document_json_bytes,
                active_bank=active_bank,
            )

            excel_name = f"stm_documentai_{active_bank.lower()}.xlsx"

            st.session_state.results = {
                "active_bank": active_bank,
                "parsed_df": parsed_df,
                "check_df": check_df,
                "raw_text": raw_text,
                "document_json_bytes": document_json_bytes,
                "excel_bytes": excel_bytes,
                "excel_name": excel_name,
            }

            progress.progress(100, text="เสร็จสิ้น ✅")
            status.success(f"สำเร็จ พบ {len(parsed_df):,} รายการ · ธนาคาร: {active_bank}")

        except Exception as e:
            progress.empty()
            status.error(f"ประมวลผลไม่สำเร็จ: {e}")
            st.exception(e)

    if st.session_state.results:
        res = st.session_state.results

        st.divider()
        st.success(f"ประมวลผลสำเร็จ พบ {len(res['parsed_df']):,} รายการ")
        st.caption(f"ธนาคาร: {res['active_bank']}")

        parsed_df = res["parsed_df"]
        check_df = res["check_df"]

        debit_total = pd.to_numeric(parsed_df.get("debit", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
        credit_total = pd.to_numeric(parsed_df.get("credit", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()

        balance_series = pd.to_numeric(parsed_df.get("balance", pd.Series(dtype=float)), errors="coerce").dropna()
        last_balance = float(balance_series.iloc[-1]) if len(balance_series) else 0.0

        ok_count = int(check_df["balance_check"].astype(str).str.contains("OK|SUGGEST", regex=True).sum()) if not check_df.empty else 0
        review_count = int(check_df["balance_check"].astype(str).str.contains("REVIEW|BROKEN", regex=True).sum()) if not check_df.empty else 0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("รายการ", f"{len(parsed_df):,}")
        c2.metric("ยอดล่าสุด", f"{last_balance:,.2f}")
        c3.metric("ผ่าน/เดาได้", f"{ok_count:,}")
        c4.metric("ต้องตรวจ", f"{review_count:,}")

        c5, c6, c7 = st.columns(3)
        c5.metric("รวมเดบิต", f"{debit_total:,.2f}")
        c6.metric("รวมเครดิต", f"{credit_total:,.2f}")
        c7.metric("สุทธิ", f"{credit_total - debit_total:,.2f}")

        st.download_button(
            "ดาวน์โหลด Excel",
            data=res["excel_bytes"],
            file_name=res["excel_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

        st.download_button(
            "ดาวน์โหลด document.json",
            data=res["document_json_bytes"],
            file_name="document.json",
            mime="application/json",
            use_container_width=True,
        )

        st.divider()

        tab1, tab2, tab3 = st.tabs(["parsed_stm", "check", "raw_text"])

        with tab1:
            if parsed_df.empty:
                st.warning("ไม่พบรายการธุรกรรม")
            else:
                show_df = parsed_df.copy()
                for col in ["debit", "credit", "balance"]:
                    show_df[col] = pd.to_numeric(show_df[col], errors="coerce")

                st.dataframe(
                    show_df.style.format({
                        "debit": lambda v: f"{v:,.2f}" if pd.notna(v) else "",
                        "credit": lambda v: f"{v:,.2f}" if pd.notna(v) else "",
                        "balance": lambda v: f"{v:,.2f}" if pd.notna(v) else "",
                    }),
                    use_container_width=True,
                    height=420,
                )

        with tab2:
            st.dataframe(check_df, use_container_width=True, height=520)

        with tab3:
            st.text_area("ข้อความ OCR จาก Document AI", res["raw_text"], height=500)

        if st.button("ล้างผลลัพธ์", use_container_width=True):
            st.session_state.results = None
            st.rerun()


# ============================================================
# 9) Entry Point
# ============================================================
if not st.session_state.authenticated:
    login_page()
else:
    main_app()

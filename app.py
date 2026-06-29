# ============================================================
# JSON Bank Statement Parser - Streamlit Edition (Full Version)
# รองรับการอัปโหลดไฟล์ document.json จาก Document AI / OCR
# ระบบอัตโนมัติ: ดึงชื่อ/เลขบัญชี, แยกเดบิต/เครดิต, หาเงินเดือน
# ============================================================

import streamlit as st
import pandas as pd
import json
import re
import tempfile
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

# ============================================================
# 1. CORE HELPERS & EXCEL FORMATTING
# ============================================================
def clean_text(value):
    if value is None: return ""
    return re.sub(r"\s+", " ", str(value).replace("\x00", " ").strip())

def estimate_text_width(value):
    text = clean_text(value)
    if not text: return 0
    w = 0.0
    for ch in text:
        if "\u0e00" <= ch <= "\u0e7f": w += 1.4
        elif ch.isupper(): w += 1.1
        elif ch.isdigit() or ch in ".,:-/()": w += 0.85
        elif ch == " ": w += 0.4
        else: w += 0.95
    return w

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
        measured_w = estimate_text_width(headers.get(col_idx, ""))
        for cell in col_cells:
            if cell.value is not None:
                parts = str(cell.value).splitlines() or [str(cell.value)]
                measured_w = max(measured_w, max(estimate_text_width(p) for p in parts))
        ws.column_dimensions[letter].width = max(8, min(measured_w + 1.5, 40))

    for row_idx in range(1, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 18 if row_idx > 1 else 22

def format_summary_sheet(ws):
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 24

    header_fill = PatternFill("solid", fgColor="1F4E79")
    section_fill = PatternFill("solid", fgColor="DDEBF7")
    sub_fill = PatternFill("solid", fgColor="EAF4EA")
    blank_fill = PatternFill("solid", fgColor="FFFFFF")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for cell in ws[1]:
        cell.fill, cell.font, cell.border = header_fill, Font(bold=True, color="FFFFFF"), border
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    for row_idx in range(2, ws.max_row + 1):
        label = clean_text(ws.cell(row=row_idx, column=1).value)
        value_cell = ws.cell(row=row_idx, column=2)
        for col in range(1, 3):
            cell = ws.cell(row=row_idx, column=col)
            cell.border, cell.alignment = border, Alignment(vertical="center", wrap_text=False)

        if not label:
            ws.row_dimensions[row_idx].height = 8
            for col in range(1, 3): ws.cell(row=row_idx, column=col).fill = blank_fill
            continue

        if label == "หมวดหมู่รายได้":
            for col in range(1, 3):
                cell = ws.cell(row=row_idx, column=col)
                cell.fill, cell.font, cell.alignment = section_fill, Font(bold=True), Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[row_idx].height = 22
            continue

        if any(k in label for k in ["สถานะ", "จำนวนรายการ", "ยอดรวม", "ค่าเฉลี่ยต่อเดือน", "เดือนที่ใช้คำนวณ"]):
            ws.cell(row=row_idx, column=1).fill = ws.cell(row=row_idx, column=2).fill = sub_fill
            ws.cell(row=row_idx, column=1).font = Font(bold=True)

        if " - รายการที่ " in label:
            ws.cell(row=row_idx, column=1).font = Font(bold=False)
            value_cell.number_format = "dd/mm/yyyy"
        elif label.endswith(" - ยอด") or isinstance(value_cell.value, (int, float)):
            value_cell.number_format = "#,##0.00"
            value_cell.alignment = Alignment(horizontal="right", vertical="center")

def format_workbook_no_color(path):
    wb = load_workbook(path)
    for ws in wb.worksheets:
        if ws.title != "BANK STATEMENT 1":
            auto_fit_worksheet(ws)
            if ws.title == "สรุปยอด": format_summary_sheet(ws)
    wb.save(path)

# ============================================================
# 2. SALARY DETECTION & SUMMARY LOGIC
# ============================================================
def detect_salary_from_df(df_txn):
    if df_txn.empty: return pd.DataFrame()
    
    salary_kws = ["SALARY", "PAYROLL", "SALA", "PAYR", "เงินเดือน", "เงินเดือนหลัก", "SAL"]
    bonus_kws = ["BONUS", "INCENTIVE", "COMMISSION", "โบนัส", "อินเซนทีฟ", "ค่าคอม"]
    income_records = []
    
    for _, row in df_txn.iterrows():
        credit = float(row.get("เครดิต", 0))
        if credit > 0:
            desc = str(row.get("รายละเอียด", "")).upper()
            is_bonus = any(kw in desc for kw in bonus_kws)
            is_salary = any(kw in desc for kw in salary_kws)
            
            if is_bonus:
                income_records.append({"กลุ่ม": "รายได้พิเศษ/โบนัส", "วันที่": row["วันที่เดือนปี"], "จำนวนเงิน": credit})
            elif is_salary:
                income_records.append({"กลุ่ม": "เงินเดือน", "วันที่": row["วันที่เดือนปี"], "จำนวนเงิน": credit})
                
    return pd.DataFrame(income_records)

def build_summary_df(df_txn, owner, acc_no, bank_name="JSON Data"):
    if df_txn.empty:
        last_bal, s_date, e_date = 0.0, "", ""
    else:
        valid = df_txn["ยอดคงเหลือ"].dropna()
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
            group_data = df_income[df_income["กลุ่ม"] == group_name].copy()
            total_amt = float(group_data["จำนวนเงิน"].sum())
            rows.extend([
                {"รายการ": f"{group_name} - สถานะ", "ข้อมูล": "พบ"},
                {"รายการ": f"{group_name} - จำนวนรายการ", "ข้อมูล": len(group_data)},
                {"รายการ": f"{group_name} - ยอดรวม", "ข้อมูล": total_amt},
                {"รายการ": f"{group_name} - ค่าเฉลี่ยต่อเดือน", "ข้อมูล": total_amt / len(group_data)},
            ])
            for idx, r in enumerate(group_data.itertuples(), start=1):
                rows.append({"รายการ": f"{group_name} - รายการที่ {idx}", "ข้อมูล": r.วันที่})
                rows.append({"รายการ": f"{group_name} - ยอด", "ข้อมูล": r.จำนวนเงิน})
        else:
            rows.extend([
                {"รายการ": f"{group_name} - สถานะ", "ข้อมูล": "ไม่พบ"},
                {"รายการ": f"{group_name} - จำนวนรายการ", "ข้อมูล": 0},
                {"รายการ": f"{group_name} - ยอดรวม", "ข้อมูล": 0.0},
            ])
        rows.append({"รายการ": "", "ข้อมูล": ""})

    rows.append({"รายการ": "สร้างไฟล์เมื่อ", "ข้อมูล": datetime.now().strftime("%d/%m/%Y %H:%M:%S")})
    return pd.DataFrame(rows)

# ============================================================
# 3. BANK STATEMENT 1 LOGIC
# ============================================================
def calculate_30day_statement_logic(df_txn, period_days=30, periods=3):
    empty_result = {"daily": pd.DataFrame(columns=["วันที่", "ยอดคงเหลือสิ้นวัน", "ช่วงที่"]), "blocks": []}
    if df_txn.empty: return empty_result

    work = df_txn.copy()
    work["_date"] = pd.to_datetime(work["วันที่เดือนปี"].apply(lambda x: str(x).replace("-","/")), dayfirst=True, errors="coerce")
    work = work.dropna(subset=["_date"])
    if work.empty: return empty_result

    work["ยอดคงเหลือ"] = pd.to_numeric(work["ยอดคงเหลือ"], errors="coerce").fillna(0.0)
    work["_row_order"] = range(len(work))
    work = work.sort_values(["_date", "_row_order"]).reset_index(drop=True)
    work["_day"] = work["_date"].dt.normalize()

    daily_txn = work.groupby("_day", as_index=True).agg(**{"ยอดคงเหลือสิ้นวัน": ("ยอดคงเหลือ", "last")}).sort_index()
    full_days = pd.date_range(start=daily_txn.index.min(), end=daily_txn.index.max(), freq="D")
    
    daily = daily_txn.reindex(full_days)
    daily["ยอดคงเหลือสิ้นวัน"] = daily["ยอดคงเหลือสิ้นวัน"].ffill().fillna(0.0)
    daily = daily.tail(period_days * periods).copy().reset_index().rename(columns={"index": "วันที่"})
    
    daily["ลำดับ"] = range(1, len(daily) + 1)
    daily["ช่วงที่"] = ((daily["ลำดับ"] - 1) // period_days) + 1

    blocks = []
    for block_no in range(1, periods + 1):
        block = daily[daily["ช่วงที่"] == block_no].copy()
        blocks.append({
            "ช่วงที่": block_no, "จำนวนวัน": len(block),
            "ผลรวม": float(block["ยอดคงเหลือสิ้นวัน"].sum()) if len(block) else 0.0,
            "ค่าเฉลี่ย": float(block["ยอดคงเหลือสิ้นวัน"].mean()) if len(block) else 0.0,
            "data": block,
        })

    daily_display = daily.copy()
    daily_display["วันที่"] = daily_display["วันที่"].dt.strftime("%d/%m/%Y")
    return {"daily": daily_display[["วันที่", "ยอดคงเหลือสิ้นวัน", "ช่วงที่"]], "blocks": blocks}

def add_bank_statement_logic_sheet(path, df_txn, owner, acc_no, bank_name="JSON Data"):
    result = calculate_30day_statement_logic(df_txn)
    blocks = result["blocks"]
    wb = load_workbook(path)
    ws = wb.create_sheet("BANK STATEMENT 1")

    title_blue = PatternFill("solid", fgColor="6FA1E8")
    header_blue = PatternFill("solid", fgColor="A9C4F5")
    light_blue = PatternFill("solid", fgColor="DDEBF7")
    yellow = PatternFill("solid", fgColor="FFFF00")
    gray = PatternFill("solid", fgColor="D9D9D9")
    
    normal_border = Border(left=Side(style="thin", color="000000"), right=Side(style="thin", color="000000"), top=Side(style="thin", color="000000"), bottom=Side(style="thin", color="000000"))
    heavy_border = Border(left=Side(style="medium", color="000000"), right=Side(style="medium", color="000000"), top=Side(style="medium", color="000000"), bottom=Side(style="medium", color="000000"))
    data_border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"), top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))

    ws.merge_cells("A1:I1")
    ws["A1"] = "BANK STATEMENT 1"
    ws["A1"].font, ws["A1"].alignment, ws["A1"].fill = Font(bold=True, size=13), Alignment(horizontal="center", vertical="center"), title_blue

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

    for row in range(2, 5):
        for col in range(1, 10):
            cell = ws.cell(row=row, column=col)
            cell.fill, cell.border, cell.font = header_blue, heavy_border, Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
    
    for addr in ["C2", "C3", "C4", "F2", "F3"]: ws[addr].fill = light_blue

    block_start_cols = [1, 4, 7]
    for idx, block in enumerate(blocks):
        start_col = block_start_cols[idx]
        n_col, date_col, amount_col = start_col, start_col + 1, start_col + 2

        for col, value in [(n_col, "N"), (date_col, "DATE"), (amount_col, "AMOUNT")]:
            cell = ws.cell(row=5, column=col, value=value)
            cell.fill, cell.font, cell.border, cell.alignment = header_blue, Font(bold=True), heavy_border, Alignment(horizontal="center", vertical="center")

        data = block.get("data", pd.DataFrame())
        for i in range(1, 31):
            excel_row = 6 + i - 1
            ws.cell(row=excel_row, column=n_col, value=i).fill = light_blue
            ws.cell(row=excel_row, column=date_col).fill = light_blue

            if i <= len(data):
                row = data.iloc[i - 1]
                date_val = row.get("วันที่", "")
                if hasattr(date_val, "to_pydatetime"): date_val = date_val.to_pydatetime()
                amt_val = float(row.get("ยอดคงเหลือสิ้นวัน", 0))
                ws.cell(row=excel_row, column=date_col, value=date_val)
                ws.cell(row=excel_row, column=amount_col, value=amt_val)
                if i > 1 and round(amt_val, 2) != round(float(data.iloc[i - 2].get("ยอดคงเหลือสิ้นวัน", 0)), 2):
                    ws.cell(row=excel_row, column=date_col).fill = yellow
                    ws.cell(row=excel_row, column=amount_col).fill = yellow
            for col in [n_col, date_col, amount_col]:
                cell = ws.cell(row=excel_row, column=col)
                cell.border, cell.alignment = data_border, Alignment(horizontal="center" if col != amount_col else "right", vertical="center")
            ws.cell(row=excel_row, column=date_col).number_format = "d mmm yy"
            ws.cell(row=excel_row, column=amount_col).number_format = "#,##0.00"

        amt_let = ws.cell(row=5, column=amount_col).column_letter
        ws.cell(row=36, column=n_col, value="รวม")
        ws.cell(row=36, column=amount_col, value=f"=SUM({amt_let}6:{amt_let}35)")
        ws.cell(row=37, column=amount_col, value=f"=IF({amt_let}39=0,0,{amt_let}36/{amt_let}39)")
        ws.cell(row=38, column=amount_col, value="จำนวนวัน")
        ws.cell(row=39, column=amount_col, value=f"=COUNT({amt_let}6:{amt_let}35)")
        for row in range(36, 40):
            for col in [n_col, date_col, amount_col]:
                cell = ws.cell(row=row, column=col)
                cell.border, cell.alignment = heavy_border, Alignment(horizontal="center" if col != amount_col else "right", vertical="center")
                if col == amount_col: cell.number_format = "#,##0.00"
        ws.cell(row=38, column=amount_col).fill = gray
        ws.cell(row=38, column=amount_col).font = Font(bold=True)

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"
    for letter, width in {"A": 5, "B": 14, "C": 16, "D": 5, "E": 14, "F": 16, "G": 5, "H": 14, "I": 16}.items():
        ws.column_dimensions[letter].width = width
    for addr in ["C2", "C3", "C4", "F2", "F3", "F4"]: ws[addr].number_format = "#,##0.00"
    wb.save(path)

# ============================================================
# 4. JSON PARSER LOGIC
# ============================================================
def extract_account_info_from_json(lines):
    owner_name, account_no = "", ""
    for i, line in enumerate(lines[:30]):
        line = line.strip()
        if not line: continue
        
        # 1. ค้นหา "ชื่อบัญชี"
        if not owner_name:
            if re.search(r'(ชื่อบัญชี|Account Name|ชื่อ\s*-\s*นามสกุล)', line, re.IGNORECASE):
                owner_name = re.sub(r'^(.*?)(ชื่อบัญชี|Account\s*Name|ชื่อ\s*-\s*นามสกุล)\s*[:：]?\s*', '', line, flags=re.IGNORECASE).strip()
                if not owner_name and i + 1 < len(lines):
                    owner_name = lines[i+1].strip()
            elif re.match(r'^(นาย|นาง|น\.ส\.|นางสาว|MR\.|MRS\.|MS\.)\s+', line, re.IGNORECASE):
                owner_name = line
                
        # 2. ค้นหา "เลขที่บัญชี"
        if not account_no:
            m = re.search(r'(\d{3}[- ]?\d{1}[- ]?\d{5}[- ]?\d{1}|\d{10,12})', line)
            if m: account_no = m.group(1).replace(" ", "")

    return owner_name if owner_name else "ไม่ระบุ", account_no if account_no else "ไม่ระบุ"

def parse_json_to_df(json_bytes, filename):
    data = json.loads(json_bytes.decode('utf-8'))
    text = data.get('text', '')
    lines = text.split('\n')

    date_pattern = re.compile(r'^\d{1,2}-\d{1,2}-\d{2,4}$')
    money_pattern = re.compile(r'\d{1,3}(?:,\d{3})*\.\d{2}')

    transactions = []
    current_tx = None

    for line in lines:
        line = line.strip()
        if not line: continue
        
        if date_pattern.match(line):
            if current_tx: transactions.append(current_tx)
            current_tx = {'date': line, 'money_tokens': [], 'raw_text': []}
        elif current_tx is not None:
            current_tx['raw_text'].append(line)
            moneys = money_pattern.findall(line)
            for m in moneys:
                current_tx['money_tokens'].append(float(m.replace(',', '')))

    if current_tx: transactions.append(current_tx)

    results = []
    prev_balance = None
    seq = 1

    for tx in transactions:
        date = tx['date']
        tokens = tx['money_tokens']
        raw_lines = " ".join(tx['raw_text'])
        debit, credit, balance = 0.0, 0.0, 0.0
        
        if len(tokens) == 1:
            balance = tokens[0]
            prev_balance = balance
        elif len(tokens) >= 2:
            amount = tokens[0]
            balance = tokens[-1]
            if prev_balance is not None:
                if abs(prev_balance - amount - balance) < 0.01: debit = amount
                elif abs(prev_balance + amount - balance) < 0.01: credit = amount
                else:
                    if balance < prev_balance: debit = amount
                    else: credit = amount
            prev_balance = balance

        desc = raw_lines
        for m in money_pattern.findall(raw_lines):
            desc = desc.replace(m, "")
        desc = re.sub(r'\s+', ' ', desc).strip()
        date_str = date.replace("-", "/")

        results.append({
            "ลำดับ": seq, "วันที่เดือนปี": date_str, "เวลา": "", "รายการ": desc[:50],
            "เดบิต": debit if debit else 0.0, "เครดิต": credit if credit else 0.0,
            "ยอดคงเหลือ": balance if balance else 0.0, "รายละเอียด": desc,
            "ช่องทาง": "JSON Upload", "หน้า": 1, "ไฟล์ต้นฉบับ": filename
        })
        seq += 1

    return pd.DataFrame(results)

# ============================================================
# 5. MASTER EXCEL BUILDER
# ============================================================
def build_output_excel(df_account, df_txn, owner, acc_no):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp_path = tmp.name

    with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
        df_account.to_excel(writer, sheet_name="เจ้าของบัญชี", index=False)
        df_txn.to_excel(writer, sheet_name="รายละเอียด เดบิต-เครดิต", index=False)
        build_summary_df(df_txn, owner, acc_no).to_excel(writer, sheet_name="สรุปยอด", index=False)

    add_bank_statement_logic_sheet(tmp_path, df_txn, owner, acc_no)
    format_workbook_no_color(tmp_path)

    with open(tmp_path, "rb") as f:
        data = f.read()
    os.unlink(tmp_path)
    return data

# ============================================================
# 6. STREAMLIT UI
# ============================================================
st.set_page_config(page_title="JSON Bank Statement", layout="centered")

st.markdown("""
<style>
    .block-container { padding-top: 2rem; max-width: 980px; }
    h1 { font-size: 2.4rem !important; font-weight: 800 !important; }
    .section-title { display: flex; align-items: center; gap: 8px; font-weight: 800; font-size: 1.15rem; margin: 18px 0 10px 0; }
    .section-title .step-number { display: inline-flex; align-items: center; justify-content: center; min-width: 30px; height: 30px; border-radius: 999px; background: #4c82fb; color: #ffffff; }
    div[data-testid="stFileUploader"] section { border-radius: 10px; }
    div[data-testid="stButton"] button { border-radius: 10px; font-weight: 700; }
    div[data-testid="stDownloadButton"] button { border-radius: 10px; font-weight: 800; }
</style>
""", unsafe_allow_html=True)

st.title("JSON Bank Statement Parser")
st.caption("ดึงข้อมูลจากไฟล์ JSON อัตโนมัติ (ชื่อบัญชี, เงินเดือน, ธุรกรรม) พร้อมสร้าง Excel 4 ชีท")
st.divider()

st.markdown("""<div class="section-title"><span class="step-number">1</span><span class="step-text">ตรวจสอบ / แก้ไขข้อมูลบัญชี (ถ้าต้องการ)</span></div>""", unsafe_allow_html=True)
st.info("💡 ระบบจะพยายามค้นหา ชื่อ และ เลขบัญชี ให้อัตโนมัติ แต่คุณสามารถกรอกดักไว้ก่อนได้")
col1, col2 = st.columns(2)
owner_name = col1.text_input("ชื่อเจ้าของบัญชี", placeholder="เว้นว่างไว้เพื่อให้ระบบหาอัตโนมัติ", value="")
account_no = col2.text_input("เลขที่บัญชี", placeholder="เว้นว่างไว้เพื่อให้ระบบหาอัตโนมัติ", value="")

st.markdown("""<div class="section-title"><span class="step-number">2</span><span class="step-text">อัปโหลดไฟล์ JSON</span></div>""", unsafe_allow_html=True)
uploaded_files = st.file_uploader("เลือกไฟล์ .json (อัปโหลดหลายไฟล์ได้)", type=["json"], accept_multiple_files=True)

st.markdown("""<div class="section-title"><span class="step-number">3</span><span class="step-text">ประมวลผล</span></div>""", unsafe_allow_html=True)
process_btn = st.button("เริ่มประมวลผล", type="primary", use_container_width=True)

if process_btn:
    if not uploaded_files:
        st.warning("กรุณาอัปโหลดไฟล์ JSON ก่อนประมวลผล")
        st.stop()
        
    all_df = []
    auto_owner = ""
    auto_acc_no = ""
    
    for f in uploaded_files:
        try:
            json_bytes = f.read()
            # 1. ให้ระบบลองค้นหาชื่อบัญชีและเลขบัญชีก่อน
            if not auto_owner or auto_owner == "ไม่ระบุ":
                data = json.loads(json_bytes.decode('utf-8'))
                lines = data.get('text', '').split('\n')
                ext_owner, ext_acc = extract_account_info_from_json(lines)
                if ext_owner != "ไม่ระบุ": auto_owner = ext_owner
                if ext_acc != "ไม่ระบุ": auto_acc_no = ext_acc
            
            # 2. ประมวลผลธุรกรรม
            df_part = parse_json_to_df(json_bytes, f.name)
            all_df.append(df_part)
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์ {f.name}: {e}")
            
    if not all_df:
        st.stop()
        
    final_df = pd.concat(all_df, ignore_index=True)
    final_df["ลำดับ"] = range(1, len(final_df) + 1)
    
    # เลือกชื่อที่จะนำไปใช้ (ลำดับความสำคัญ: ผู้ใช้พิมพ์เอง > โปรแกรมหาเจอ > "ไม่ระบุ")
    final_owner = owner_name if (owner_name and owner_name.strip() != "") else (auto_owner if auto_owner else "ไม่ระบุ")
    final_acc_no = account_no if (account_no and account_no.strip() != "") else (auto_acc_no if auto_acc_no else "ไม่ระบุ")
    
    df_account = pd.DataFrame([{"เจ้าของบัญชี": final_owner, "เลขที่บัญชีเงินฝาก": final_acc_no}])
    excel_bytes = build_output_excel(df_account, final_df, final_owner, final_acc_no)
    
    st.session_state["result_df"] = final_df
    st.session_state["excel_bytes"] = excel_bytes
    st.session_state["owner_name"] = final_owner
    st.session_state["account_no"] = final_acc_no

if "result_df" in st.session_state:
    df = st.session_state["result_df"]
    st.divider()
    st.success(f"ประมวลผลสำเร็จ! พบ {len(df):,} รายการ")
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("เจ้าของบัญชี", st.session_state["owner_name"])
    c2.metric("จำนวนรายการ", f"{len(df):,}")
    c3.metric("รวมเดบิต", f"{df['เดบิต'].sum():,.2f}")
    c4.metric("รวมเครดิต", f"{df['เครดิต'].sum():,.2f}")

    st.markdown("#### ตัวอย่างข้อมูลที่ประมวลผลได้ (30 รายการแรก)")
    st.dataframe(
        df.head(30).style.format({"เดบิต": "{:,.2f}", "เครดิต": "{:,.2f}", "ยอดคงเหลือ": "{:,.2f}"}),
        use_container_width=True, height=400
    )

    st.divider()
    st.download_button(
        label="📥 ดาวน์โหลดไฟล์ Excel (ครบ 4 ชีท + สรุปเงินเดือน)",
        data=st.session_state["excel_bytes"],
        file_name=f"Statement_{st.session_state['owner_name']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )
    
    if st.button("ล้างข้อมูลหลังใช้งาน", use_container_width=True):
        for key in ["result_df", "excel_bytes", "owner_name", "account_no"]:
            if key in st.session_state: del st.session_state[key]
        st.rerun()

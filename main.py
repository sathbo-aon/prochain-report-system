"""
ProChain Daily Report System
=============================
อ่านอีเมล Outlook → ดาวน์โหลด Excel → บันทึก OneDrive → เขียน Google Sheet

Environment variables ที่ต้องตั้งใน GitHub Secrets:
  AZURE_CLIENT_ID       — Application (client) ID
  AZURE_CLIENT_SECRET   — Client Secret Value
  AZURE_TENANT_ID       — Directory (tenant) ID
  AZURE_USER_EMAIL      — อีเมลที่ใช้รับ daily report
  AZURE_USER_PASSWORD   — password ของอีเมลนั้น
  GOOGLE_SERVICE_ACCOUNT_JSON — JSON ทั้งก้อนของ service account
  SPREADSHEET_ID        — Google Sheet ID
  SYSTEM_ACTIVE         — TRUE หรือ FALSE (dead key)
"""

import os
import io
import json
import base64
import logging
from datetime import datetime, timezone, timedelta

import msal
import requests
from openpyxl import load_workbook
import gspread
from google.oauth2.service_account import Credentials

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

# ─── Config ────────────────────────────────────────────────────────
CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
TENANT_ID     = os.environ["AZURE_TENANT_ID"]
USER_EMAIL    = os.environ["AZURE_USER_EMAIL"]
USER_PASSWORD = os.environ["AZURE_USER_PASSWORD"]
SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
SYSTEM_ACTIVE  = os.environ.get("SYSTEM_ACTIVE", "TRUE").upper()

ONEDRIVE_FOLDER = "/ProChain Reports"
SUBJECT_KEYWORD = "daily report"

SCOPES = [
    "https://graph.microsoft.com/Mail.Read",
    "https://graph.microsoft.com/Files.ReadWrite",
    "https://graph.microsoft.com/User.Read",
]

GRAPH = "https://graph.microsoft.com/v1.0"

# ─── Dead Key ──────────────────────────────────────────────────────
def check_dead_key():
    if SYSTEM_ACTIVE != "TRUE":
        log.warning("🔴 SYSTEM_ACTIVE = FALSE — ระบบถูกปิดการทำงาน ไม่บันทึกข้อมูล")
        return False
    return True

# ─── Microsoft Graph Auth (Delegated) ──────────────────────────────
def get_access_token() -> str:
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    )
    result = app.acquire_token_by_username_password(
        username=USER_EMAIL,
        password=USER_PASSWORD,
        scopes=SCOPES,
    )
    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description')}")
    log.info("✅ ได้ access token แล้ว")
    return result["access_token"]

# ─── Outlook: ดึงอีเมลวันนี้ ──────────────────────────────────────
def fetch_today_emails(token: str) -> list[dict]:
    headers = {"Authorization": f"Bearer {token}"}
    tz_thai = timezone(timedelta(hours=7))
    today = datetime.now(tz_thai).strftime("%Y-%m-%dT00:00:00Z")

    filter_q = (
        f"receivedDateTime ge {today}"
        f" and contains(subject, '{SUBJECT_KEYWORD}')"
        f" and hasAttachments eq true"
    )
    url = f"{GRAPH}/me/messages?$filter={filter_q}&$select=id,subject,receivedDateTime"
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    messages = resp.json().get("value", [])
    log.info(f"📬 พบ {len(messages)} อีเมลวันนี้")
    return messages

# ─── Outlook: ดึง attachments ─────────────────────────────────────
def fetch_attachments(token: str, message_id: str) -> list[dict]:
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH}/me/messages/{message_id}/attachments"
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.json().get("value", [])

# ─── OneDrive: บันทึกไฟล์ ─────────────────────────────────────────
def save_to_onedrive(token: str, filename: str, content_b64: str, project_name: str, report_date: str) -> str:
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/octet-stream",
    }
    content = base64.b64decode(content_b64)

    # จัดโฟลเดอร์ตาม Site และเดือน
    try:
        date_obj = datetime.strptime(report_date.strip(), "%d/%m/%Y")
        month_folder = date_obj.strftime("%Y-%m")
    except Exception:
        month_folder = datetime.now().strftime("%Y-%m")

    folder_path = f"{ONEDRIVE_FOLDER}/{project_name}/{month_folder}"
    safe_filename = filename.replace("'", "")
    upload_url = f"{GRAPH}/me/drive/root:{folder_path}/{safe_filename}:/content"

    resp = requests.put(upload_url, headers=headers, data=content)
    resp.raise_for_status()
    log.info(f"💾 บันทึกไฟล์ OneDrive: {folder_path}/{safe_filename}")
    return safe_filename

# ─── Excel: อ่านข้อมูล ────────────────────────────────────────────
def parse_excel(content_b64: str, filename: str) -> dict:
    content = base64.b64decode(content_b64)
    wb = load_workbook(io.BytesIO(content), data_only=True)
    ws = wb.active

    def cell(r, c):
        v = ws.cell(row=r, column=c).value
        if v is None:
            return ""
        if isinstance(v, datetime):
            return v.strftime("%d/%m/%Y")
        return str(v).strip()

    # ─── Header fields ──────────────────────────────────────────
    project_name = cell(7, 3)
    report_date  = cell(8, 3)
    week_no      = cell(8, 6)
    report_no    = cell(8, 9)
    site_manager = cell(9, 3)
    contract_no  = cell(9, 6)
    work_hours   = cell(10, 3)
    weather_am   = cell(10, 7)
    weather_pm   = cell(10, 9)
    workforce    = cell(13, 8)
    accident     = cell(39, 2)
    near_miss    = cell(39, 3)
    first_aid    = cell(39, 4)
    remarks      = cell(44, 1)
    imported_at  = datetime.now().strftime("%d/%m/%Y %H:%M")

    # ─── Activity rows (dynamic) ────────────────────────────────
    activities = []
    r = 16  # เริ่มจาก row แรกของข้อมูล (ต่อจาก header row 15)
    while True:
        act_code = cell(r, 1)
        act_desc = cell(r, 3)
        plan_pct = cell(r, 6)
        actual_pct = cell(r, 7)
        remark   = cell(r, 8)

        # หยุดเมื่อไม่มีข้อมูลทั้ง ActivityCode และ Description
        if not act_code and not act_desc:
            if r > 19:  # อ่านอย่างน้อย 4 แถว
                break
            r += 1
            continue

        activities.append({
            "วันที่": report_date,
            "ชื่อโครงการ": project_name,
            "Report No.": report_no,
            "ActivityCode": act_code,
            "Description": act_desc,
            "PlanPct": plan_pct,
            "ActualPct": actual_pct,
            "Variance": str(
                round(float(actual_pct) - float(plan_pct), 2)
            ) if plan_pct and actual_pct else "",
            "Status": "Delayed" if plan_pct and actual_pct and
                      float(actual_pct) < float(plan_pct) else "On Track",
            "Remark": remark,
            "ไฟล์ต้นฉบับ": filename,
            "นำเข้าเมื่อ": imported_at,
        })
        r += 1
        if r > 100:  # safety limit
            break

    # ─── Material rows (dynamic) ────────────────────────────────
    materials = []
    r = 22  # เริ่มจาก row แรกของข้อมูล (ต่อจาก header row 21)
    while True:
        mat_code = cell(r, 2)
        mat_desc = cell(r, 3)
        mat_unit = cell(r, 5)
        mat_qty  = cell(r, 6)
        mat_inout = cell(r, 7)
        mat_supplier = cell(r, 8)

        if not mat_desc and not mat_qty:
            if r > 29:
                break
            r += 1
            continue

        materials.append({
            "วันที่": report_date,
            "ชื่อโครงการ": project_name,
            "Report No.": report_no,
            "MatCode": mat_code,
            "Description": mat_desc,
            "Unit": mat_unit,
            "Qty": mat_qty,
            "InOut": mat_inout,
            "Supplier": mat_supplier,
            "ไฟล์ต้นฉบับ": filename,
            "นำเข้าเมื่อ": imported_at,
        })
        r += 1
        if r > 200:
            break

    # ─── Daily Log ──────────────────────────────────────────────
    daily_log = {
        "วันที่": report_date,
        "ชื่อโครงการ": project_name,
        "Report No.": report_no,
        "สัปดาห์ที่": week_no,
        "Site Manager": site_manager,
        "เลขที่สัญญา": contract_no,
        "ชั่วโมงทำงาน": work_hours,
        "อากาศเช้า": weather_am,
        "อากาศบ่าย": weather_pm,
        "คนงานรวม": workforce,
        "อุบัติเหตุ": accident,
        "Near Miss": near_miss,
        "First Aid": first_aid,
        "หมายเหตุ": remarks,
        "ไฟล์ต้นฉบับ": filename,
        "นำเข้าเมื่อ": imported_at,
    }

    log.info(f"📊 อ่านข้อมูล: {len(activities)} กิจกรรม, {len(materials)} รายการวัสดุ")
    return {
        "project_name": project_name,
        "report_date": report_date,
        "daily_log": daily_log,
        "activities": activities,
        "materials": materials,
    }

# ─── Google Sheets ────────────────────────────────────────────────
def get_gsheet_client():
    sa_json = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    creds_dict = json.loads(sa_json)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)

def check_system_active_in_sheet(ss) -> bool:
    """เช็ค dead key จาก Config sheet"""
    try:
        config = ss.worksheet("Config")
        val = config.acell("B1").value
        if str(val).upper() != "TRUE":
            log.warning("🔴 Config sheet: SYSTEM_ACTIVE = FALSE — หยุดการบันทึก")
            return False
        return True
    except Exception as e:
        log.warning(f"⚠️ อ่าน Config sheet ไม่ได้: {e} — ดำเนินการต่อ")
        return True

def append_to_sheet(ss, sheet_name: str, rows: list[list]):
    ws = ss.worksheet(sheet_name)
    if rows:
        ws.append_rows(rows, value_input_option="USER_ENTERED")
        log.info(f"✅ บันทึก {len(rows)} แถวลง '{sheet_name}'")

def write_to_gsheet(data: dict):
    client = get_gsheet_client()
    ss = client.open_by_key(SPREADSHEET_ID)

    if not check_system_active_in_sheet(ss):
        return

    dl = data["daily_log"]
    # Daily Log
    daily_row = [[
        dl["วันที่"], dl["ชื่อโครงการ"], dl["Report No."], dl["สัปดาห์ที่"],
        dl["Site Manager"], dl["เลขที่สัญญา"], dl["ชั่วโมงทำงาน"],
        dl["อากาศเช้า"], dl["อากาศบ่าย"], dl["คนงานรวม"],
        dl["อุบัติเหตุ"], dl["Near Miss"], dl["First Aid"],
        dl["หมายเหตุ"], dl["ไฟล์ต้นฉบับ"], dl["นำเข้าเมื่อ"],
    ]]
    append_to_sheet(ss, "Daily Log", daily_row)

    # Work Progress
    act_rows = [[
        a["วันที่"], a["ชื่อโครงการ"], a["Report No."],
        a["ActivityCode"], a["Description"],
        a["PlanPct"], a["ActualPct"], a["Variance"],
        a["Status"], a["Remark"],
        a["ไฟล์ต้นฉบับ"], a["นำเข้าเมื่อ"],
    ] for a in data["activities"]]
    append_to_sheet(ss, "Work Progress", act_rows)

    # Material Ledger
    mat_rows = [[
        m["วันที่"], m["ชื่อโครงการ"], m["Report No."],
        m["MatCode"], m["Description"], m["Unit"],
        m["Qty"], m["InOut"], m["Supplier"],
        m["ไฟล์ต้นฉบับ"], m["นำเข้าเมื่อ"],
    ] for m in data["materials"]]
    append_to_sheet(ss, "Material Ledger", mat_rows)

    # Stock Summary — เพิ่ม/อัปเดตรายการ
    update_stock_summary(ss, data["materials"])

def update_stock_summary(ss, materials: list[dict]):
    """อัปเดต Stock Summary โดยหา row ที่ตรงกับ ชื่อโครงการ + MatCode แล้ว recalculate"""
    if not materials:
        return
    ws = ss.worksheet("Stock Summary")
    existing = ws.get_all_values()
    # header = existing[0] ถ้ามี

    # หา pairs ที่ยังไม่มีใน summary แล้วเพิ่มแถวใหม่
    existing_keys = set()
    for row in existing[1:]:  # skip header
        if len(row) >= 2:
            existing_keys.add((row[0], row[1]))  # ชื่อโครงการ, MatCode

    new_rows = []
    for m in materials:
        key = (m["ชื่อโครงการ"], m["MatCode"])
        if key not in existing_keys:
            # เพิ่มแถวใหม่ สูตร SUMIF จะคำนวณเองใน Sheet
            project = m["ชื่อโครงการ"]
            mat_code = m["MatCode"]
            desc = m["Description"]
            unit = m["Unit"]
            new_rows.append([
                project, mat_code, desc, unit,
                f'=SUMIFS(\'Material Ledger\'!G:G,\'Material Ledger\'!B:B,A{len(existing)+len(new_rows)+1},\'Material Ledger\'!D:D,B{len(existing)+len(new_rows)+1},\'Material Ledger\'!H:H,"IN")',
                f'=SUMIFS(\'Material Ledger\'!G:G,\'Material Ledger\'!B:B,A{len(existing)+len(new_rows)+1},\'Material Ledger\'!D:D,B{len(existing)+len(new_rows)+1},\'Material Ledger\'!H:H,"OUT")',
                f'=E{len(existing)+len(new_rows)+1}-F{len(existing)+len(new_rows)+1}',
                "",  # Threshold (กรอกเองใน Sheet)
                f'=MAXIFS(\'Material Ledger\'!A:A,\'Material Ledger\'!B:B,A{len(existing)+len(new_rows)+1},\'Material Ledger\'!D:D,B{len(existing)+len(new_rows)+1})',
            ])
            existing_keys.add(key)

    if new_rows:
        ws.append_rows(new_rows, value_input_option="USER_ENTERED")
        log.info(f"✅ เพิ่ม {len(new_rows)} รายการใหม่ใน Stock Summary")

# ─── Main ─────────────────────────────────────────────────────────
def main():
    log.info("🚀 เริ่มระบบ ProChain Daily Report")

    # เช็ค dead key จาก env variable ก่อน
    if not check_dead_key():
        return

    token = get_access_token()
    messages = fetch_today_emails(token)

    if not messages:
        log.info("📭 ไม่มีอีเมลใหม่วันนี้")
        return

    for msg in messages:
        subject = msg.get("subject", "")
        log.info(f"📧 ประมวลผล: {subject}")

        # แยก project name และ date จาก subject
        # format: "daily report-ชื่อโครงการ-วันที่"
        parts = subject.split("-")
        project_from_subject = parts[1].strip() if len(parts) > 1 else ""
        date_from_subject    = parts[2].strip() if len(parts) > 2 else ""

        attachments = fetch_attachments(token, msg["id"])
        for att in attachments:
            name = att.get("name", "")
            if not name.lower().endswith((".xlsx", ".xls")):
                continue

            log.info(f"📎 ประมวลผลไฟล์: {name}")
            content_b64 = att.get("contentBytes", "")

            # อ่าน Excel
            data = parse_excel(content_b64, name)

            # ใช้ข้อมูลจาก Excel ถ้า parse ได้ ไม่งั้นใช้จาก subject
            if not data["project_name"]:
                data["daily_log"]["ชื่อโครงการ"] = project_from_subject
                data["project_name"] = project_from_subject
            if not data["report_date"]:
                data["daily_log"]["วันที่"] = date_from_subject
                data["report_date"] = date_from_subject

            # บันทึกลง OneDrive
            try:
                save_to_onedrive(
                    token, name, content_b64,
                    data["project_name"], data["report_date"]
                )
            except Exception as e:
                log.warning(f"⚠️ บันทึก OneDrive ไม่ได้: {e} — ดำเนินการต่อ")

            # บันทึกลง Google Sheet
            write_to_gsheet(data)

    log.info("✅ เสร็จสิ้น")

if __name__ == "__main__":
    main()

"""
setup_sheets.py
================
ตั้งค่า header row ใน Google Sheet ทุก sheet
รันครั้งเดียวก่อนใช้งานระบบ

วิธีใช้:
  GOOGLE_SERVICE_ACCOUNT_JSON='...' SPREADSHEET_ID='...' python setup_sheets.py
"""

import os
import json
import gspread
from google.oauth2.service_account import Credentials

SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]

HEADERS = {
    "Daily Log": [
        "วันที่", "ชื่อโครงการ", "Report No.", "สัปดาห์ที่",
        "Site Manager", "เลขที่สัญญา", "ชั่วโมงทำงาน",
        "อากาศเช้า", "อากาศบ่าย", "คนงานรวม",
        "อุบัติเหตุ", "Near Miss", "First Aid",
        "หมายเหตุ", "ไฟล์ต้นฉบับ", "นำเข้าเมื่อ",
    ],
    "Work Progress": [
        "วันที่", "ชื่อโครงการ", "Report No.",
        "ActivityCode", "Description",
        "PlanPct", "ActualPct", "Variance",
        "Status", "Remark",
        "ไฟล์ต้นฉบับ", "นำเข้าเมื่อ",
    ],
    "Material Ledger": [
        "วันที่", "ชื่อโครงการ", "Report No.",
        "MatCode", "Description", "Unit",
        "Qty", "InOut", "Supplier",
        "ไฟล์ต้นฉบับ", "นำเข้าเมื่อ",
    ],
    "Stock Summary": [
        "ชื่อโครงการ", "MatCode", "Description", "Unit",
        "รวม IN", "รวม OUT", "คงเหลือ",
        "Threshold แจ้งเตือน", "อัปเดตล่าสุด",
    ],
}

def main():
    sa_json = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    creds_dict = json.loads(sa_json)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    ss = client.open_by_key(SPREADSHEET_ID)

    for sheet_name, headers in HEADERS.items():
        try:
            ws = ss.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            ws = ss.add_worksheet(title=sheet_name, rows=1000, cols=len(headers))
            print(f"✅ สร้าง sheet ใหม่: {sheet_name}")

        # เช็คว่ามี header แล้วหรือยัง
        existing = ws.row_values(1)
        if not existing:
            ws.insert_row(headers, index=1)
            print(f"✅ ใส่ header: {sheet_name}")
        else:
            print(f"⏭️  มี header แล้ว: {sheet_name}")

    # Config sheet
    try:
        config = ss.worksheet("Config")
    except gspread.exceptions.WorksheetNotFound:
        config = ss.add_worksheet(title="Config", rows=10, cols=2)

    config.update("A1:B3", [
        ["SYSTEM_ACTIVE", "TRUE"],
        ["DEACTIVATED_BY", ""],
        ["DEACTIVATED_DATE", ""],
    ])
    print("✅ Config sheet พร้อมแล้ว")
    print("\n🎉 ตั้งค่า Google Sheet เสร็จสิ้น")

if __name__ == "__main__":
    main()

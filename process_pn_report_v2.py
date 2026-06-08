"""
Park+ MoEngage Daily PN Report — Full Automation Script
=========================================================
This script does everything automatically:
1. Calls MoEngage API to download yesterday's campaign report
2. Processes it (same as your Excel macro)
3. Saves a clean Excel file
4. Appends data to Google Sheet (for live dashboard)
5. Sends you a daily email summary
"""

import os
import re
import json
import hashlib
import zipfile
import io
import smtplib
import pandas as pd
import requests
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


# ═══════════════════════════════════════════════════════════
#  SETTINGS — all loaded from Railway Environment Variables
# ═══════════════════════════════════════════════════════════

WORKSPACE_ID     = os.environ.get("WORKSPACE_ID")
DATA_API_KEY     = os.environ.get("DATA_API_KEY")
CAMPAIGN_API_KEY = os.environ.get("CAMPAIGN_API_KEY")
EMAIL_SENDER     = os.environ.get("EMAIL_SENDER")
EMAIL_PASSWORD   = os.environ.get("EMAIL_PASSWORD")
EMAIL_RECIPIENT  = os.environ.get("EMAIL_RECIPIENT")
GOOGLE_SHEET_ID  = os.environ.get("GOOGLE_SHEET_ID")
GOOGLE_CREDS_JSON = os.environ.get("GOOGLE_CREDS_JSON")

OUTPUT_FOLDER = "./output"

MOENGAGE_API_BASE = "https://api-01.moengage.com"


# ─────────────────────────────────────────────
# STEP 1: DOWNLOAD REPORT FROM MOENGAGE API
# ─────────────────────────────────────────────

def generate_signature(workspace_id, filename, secret_key):
    signature_string = f"{workspace_id}|{filename}|{secret_key}"
    return hashlib.sha256(signature_string.encode('utf-8')).hexdigest()


def download_moengage_report(date=None):
    if date is None:
        report_date = datetime.now() - timedelta(days=1)
    else:
        report_date = date

    filename = report_date.strftime("%Y%m%d") + ".zip"
    date_str = report_date.strftime("%Y-%m-%d")

    print(f"[1/5] Downloading MoEngage report for {date_str}...")

    signature = generate_signature(WORKSPACE_ID, filename, CAMPAIGN_API_KEY)
    url = f"{MOENGAGE_API_BASE}/dailyCampaignReportDump/{WORKSPACE_ID}/{filename}?Signature={signature}"

    response = requests.get(url, timeout=60)

    print(f"      → Downloaded {len(response.content) / 1024:.1f} KB")
    print(f"      → Status: {response.status_code}")
    print(f"      → Response preview: {response.text[:300]}")

    if response.status_code != 200:
        raise Exception(f"MoEngage API failed: {response.text}")

    with zipfile.ZipFile(io.BytesIO(response.content)) as z:
        file_list = z.namelist()
        print(f"      → Files in ZIP: {file_list}")

        target_file = None
        for f in file_list:
            if 'PUSH' in f.upper():
                target_file = f
                break
        if not target_file:
            target_file = file_list[0]

        print(f"      → Reading: {target_file}")
        with z.open(target_file) as excel_file:
            try:
                df = pd.read_excel(excel_file, engine="openpyxl", sheet_name="sheet1")
            except Exception:
                df = pd.read_excel(excel_file, engine="openpyxl")

    print(f"      → {len(df)} rows loaded from MoEngage")
    return df


# ─────────────────────────────────────────────
# STEP 2: PROCESS (replaces your Excel macro)
# ─────────────────────────────────────────────

def parse_campaign_name(raw_name):
    name_clean = re.sub(r'\s*@.*$', '', str(raw_name)).strip()
    parts = name_clean.split('_')

    def get(i):
        return parts[i] if len(parts) > i else ''

    return {
        'Campaign Name':     get(0),
        'Category':          get(1),
        'Sub-Category':      get(2) if get(2) not in ('NA', '') else None,
        'Sticky/Non-sticky': get(3),
        'Frequency':         get(4),
        'Raw Date':          get(5),
        'Time':              get(6),
        'Audience':          get(7),
        'Tone':              get(8),
        'Emotion':           '_'.join(parts[9:]) if len(parts) > 9 else '',
    }


def convert_date(raw_date_str, year=None):
    if not raw_date_str:
        return None
    try:
        y = year or datetime.now().year
        return datetime.strptime(f"{raw_date_str}{y}", "%d%b%Y").date()
    except Exception:
        return raw_date_str


def process_report(raw_df):
    print(f"[2/5] Processing report...")

    df = raw_df[raw_df['Campaign Status'] == 'Sent'].copy()
    print(f"      → {len(df)} rows after filtering to Campaign Status = 'Sent'")

    parsed = df['Campaign Name'].apply(parse_campaign_name)
    parsed_df = pd.DataFrame(parsed.tolist())
    parsed_df['Date'] = parsed_df['Raw Date'].apply(convert_date)

    output = pd.DataFrame({
        'Campaign Status':        df['Campaign Status'].values,
        'Campaign Name':          parsed_df['Campaign Name'].values,
        'Category':               parsed_df['Category'].values,
        'Sub-Category':           parsed_df['Sub-Category'].values,
        'Sticky/Non-sticky':      parsed_df['Sticky/Non-sticky'].values,
        'Frequency':              parsed_df['Frequency'].values,
        'Date':                   parsed_df['Date'].values,
        'Time':                   parsed_df['Time'].values,
        'Audience':               parsed_df['Audience'].values,
        'Tone':                   parsed_df['Tone'].values,
        'Emotion':                parsed_df['Emotion'].values,
        'Campaign Delivery Type': df['Campaign Delivery Type'].values,
        'Variation':              df['Variation'].values,
        'Android Message Title (Android, Web), Title (iOS)':  df['Android Message Title (Android, Web), Title (iOS)'].values,
        'Android Message (Android, Web), Subtitle (iOS)':     df['Android Message (Android, Web), Subtitle (iOS)'].values,
        'All Platform Impressions': df['All Platform Impressions'].values,
        'All Platform Clicks':      df['All Platform Clicks'].values,
        'All Platform CTR':         df['All Platform CTR'].values,
        'Android Impressions':      df['Android Impressions'].values,
        'Android Clicks':           df['Android Clicks'].values,
        'Android CTR':              df['Android CTR'].values,
        'Ios Impressions':          df['Ios Impressions'].values,
        'Ios Clicks':               df['Ios Clicks'].values,
        'Ios CTR':                  df['Ios CTR'].values,
    })

    print(f"      → Done. {len(output)} rows, {len(output.columns)} columns")
    return output


# ─────────────────────────────────────────────
# STEP 3: SAVE TO EXCEL
# ─────────────────────────────────────────────

def save_to_excel(df):
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    today = datetime.now().strftime("%Y-%m-%d")
    path = os.path.join(OUTPUT_FOLDER, f"PN_Report_{today}.xlsx")

    print(f"[3/5] Saving Excel → {path}")
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Sheet1", index=False)
        ws = writer.sheets["Sheet1"]
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)

    print(f"      → Saved.")
    return path


# ─────────────────────────────────────────────
# STEP 4: APPEND TO GOOGLE SHEET
# ─────────────────────────────────────────────

def append_to_google_sheet(df):
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        print(f"[4/5] Appending to Google Sheet...")

        creds_dict = json.loads(GOOGLE_CREDS_JSON)
        creds = Credentials.from_service_account_info(
            creds_dict,
            scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )

        gc = gspread.authorize(creds)
        ws = gc.open_by_key(GOOGLE_SHEET_ID).sheet1

        if not ws.get_all_values():
            ws.append_row(list(df.columns))
            print("      → Headers written (first time setup)")

        df_copy = df.copy()
        df_copy['Date'] = df_copy['Date'].astype(str)
        rows = df_copy.fillna('').values.tolist()
        ws.append_rows(rows, value_input_option='USER_ENTERED')
        print(f"      → {len(rows)} rows appended.")

    except ImportError:
        print("      [SKIP] Run: pip install gspread google-auth")
    except Exception as e:
        print(f"      [ERROR] Google Sheet failed: {e}")


# ─────────────────────────────────────────────
# STEP 5: SEND EMAIL
# ─────────────────────────────────────────────

def send_email(df, excel_path):
    try:
        import base64

        print(f"[5/5] Sending email to {EMAIL_RECIPIENT}...")
        today = datetime.now().strftime("%d %b %Y")

        total_campaigns   = df['Campaign Name'].nunique()
        total_variations  = len(df)
        total_impressions = df['All Platform Impressions'].sum()
        total_clicks      = df['All Platform Clicks'].sum()
        avg_ctr           = df['All Platform CTR'].mean()

        top5 = df.nlargest(5, 'All Platform CTR')[
            ['Category', 'Date', 'Time', 'Audience',
             'All Platform Impressions', 'All Platform Clicks', 'All Platform CTR']
        ].to_html(index=False, border=1)

        html = f"""
        <html><body style="font-family:Arial,sans-serif;">
        <h2>📱 Daily Push Notification Report — {today}</h2>
        <table cellpadding="8" style="border-collapse:collapse;">
            <tr style="background:#f0f0f0"><td><b>Total Campaigns</b></td><td>{total_campaigns}</td></tr>
            <tr><td><b>Total Variations</b></td><td>{total_variations}</td></tr>
            <tr style="background:#f0f0f0"><td><b>Total Impressions</b></td><td>{total_impressions:,.0f}</td></tr>
            <tr><td><b>Total Clicks</b></td><td>{total_clicks:,.0f}</td></tr>
            <tr style="background:#f0f0f0"><td><b>Average CTR</b></td><td>{avg_ctr:.2f}%</td></tr>
        </table>
        <br>
        <h3>🏆 Top 5 Campaigns by CTR</h3>
        {top5}
        <br>
        <p>Full report is attached as Excel.</p>
        </body></html>
        """

        import sendgrid
        from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition
        import base64

        sg_api_key = os.environ.get("SENDGRID_API_KEY")

        with open(excel_path, 'rb') as f:
            encoded = base64.b64encode(f.read()).decode()

        message = Mail(
            from_email=EMAIL_SENDER,
            to_emails=EMAIL_RECIPIENT,
            subject=f"PN Report {today} — {total_campaigns} campaigns | Avg CTR {avg_ctr:.2f}%",
            html_content=html
        )

        attachment = Attachment(
            FileContent(encoded),
            FileName(os.path.basename(excel_path)),
            FileType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            Disposition("attachment")
        )
        message.attachment = attachment

        sg = sendgrid.SendGridAPIClient(api_key=sg_api_key)
        response = sg.send(message)
        print(f"      → Email sent! Status: {response.status_code}")
    except Exception as e:
        print(f"      [ERROR] Email failed: {e}")


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

if __name__ == "__main__":
    print("=" * 50)
    print("Park+ PN Report Automation")
    print(f"Running at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 50)

    raw_df     = download_moengage_report()
    clean_df   = process_report(raw_df)
    excel_path = save_to_excel(clean_df)
    append_to_google_sheet(clean_df)
    send_email(clean_df, excel_path)

    print("\n✅ All done!")

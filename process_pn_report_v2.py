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
    """
    Downloads daily campaign report using Campaign Report API.
    Uses Basic Auth + Signature header.
    """
    if date is None:
        report_date = datetime.now()   # today
    else:
        report_date = date

    import base64

    # Build filename: ReportName_YYYYMMDD.zip
    report_name = "Daily-PN-Report"
    date_str    = report_date.strftime("%Y%m%d")
    filename    = f"{report_name}_{date_str}.zip"

    print(f"[1/5] Downloading MoEngage report: {filename}")

    # Basic Auth: username=WORKSPACE_ID, password=CAMPAIGN_API_KEY
    credentials = base64.b64encode(
        f"{WORKSPACE_ID}:{CAMPAIGN_API_KEY}".encode()
    ).decode()

    # Signature: SHA256(WORKSPACE_ID + "|" + FILENAME + "|" + CAMPAIGN_API_KEY)
    signature = generate_signature(WORKSPACE_ID, filename, CAMPAIGN_API_KEY)

    url = f"{MOENGAGE_API_BASE}/campaign_reports/rest_api/{WORKSPACE_ID}/{filename}"

    response = requests.get(
        url,
        headers={
            "Authorization": f"Basic {credentials}",
            "Signature": signature
        },
        timeout=60
    )

    print(f"      → Status: {response.status_code}")
    print(f"      → Downloaded {len(response.content) / 1024:.1f} KB")

    if response.status_code != 200:
        print(f"      → Response: {response.text[:300]}")
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
        with z.open(target_file) as data_file:
            if target_file.endswith('.csv'):
                df = pd.read_csv(data_file)
            else:
                try:
                    df = pd.read_excel(data_file, engine="openpyxl", sheet_name="sheet1")
                except Exception:
                    df = pd.read_excel(data_file, engine="openpyxl")

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

    # Filter: only Sent campaigns
    df = raw_df[raw_df['Campaign Status'] == 'Sent'].copy()
    print(f"      → {len(df)} rows after filtering to Campaign Status = 'Sent'")

    # Parse campaign name into structured columns
    parsed = df['Campaign Name'].apply(parse_campaign_name)
    parsed_df = pd.DataFrame(parsed.tolist())
    parsed_df['Date'] = parsed_df['Raw Date'].apply(convert_date)

    # Safely get column — returns None if column doesn't exist
    def safe_col(col_name):
        if col_name in df.columns:
            return df[col_name].values
        return [None] * len(df)

    # Build output matching Sheet1 format
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
        'Campaign Delivery Type': safe_col('Campaign Delivery Type'),
        'Variation':              safe_col('Campaign Version Name'),
        'Android Message Title':  safe_col('Android Message Title (Android, Web), Title (iOS)'),
        'Android Message':        safe_col('Android Message (Android, Web), Subtitle (iOS)'),
        'All Platform Impressions': safe_col('All Platform Impressions'),
        'All Platform Clicks':      safe_col('All Platform Clicks'),
        'All Platform CTR':         safe_col('All Platform CTR'),
        'Android Impressions':      safe_col('Android Impressions'),
        'Android Clicks':           safe_col('Android Clicks'),
        'Android CTR':              safe_col('Android CTR'),
        'Ios Impressions':          safe_col('Ios Impressions'),
        'Ios Clicks':               safe_col('Ios Clicks'),
        'Ios CTR':                  safe_col('Ios CTR'),
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
        import sendgrid
        from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition
        import base64
        import gspread
        from google.oauth2.service_account import Credentials
        import json

        print(f"[5/5] Sending email to {EMAIL_RECIPIENT}...")
        today = datetime.now().strftime("%d %b %Y")

        # ── Summary metrics ──────────────────────────────
        total_campaigns   = df['Campaign Name'].nunique()
        total_impressions = int(df['All Platform Impressions'].sum())
        total_clicks      = int(df['All Platform Clicks'].sum())
        avg_ctr           = df['All Platform CTR'].mean()

        # ── Campaign breakdown by Type → Category ────────
        # Aggregate to campaign level (combine variations)
        camp_df = df.groupby(['Campaign Name', 'Category', 'Date', 'Time', 'Audience']).agg(
            Type=('Campaign Status', 'first'),
            Impressions=('All Platform Impressions', 'sum'),
            Clicks=('All Platform Clicks', 'sum'),
        ).reset_index()
        camp_df['CTR'] = (camp_df['Clicks'] / camp_df['Impressions'] * 100).round(2)

        # Build Type → Category table
        type_col = 'Frequency'  # use Frequency as Type (One-time, Periodic etc)
        # Actually use Campaign Name part[0] as type — already in 'Campaign Name' parsed col
        # Group by Category for the table
        cat_rows = ""
        for _, row in camp_df.sort_values('CTR', ascending=False).iterrows():
            ctr_color = "#27ae60" if row['CTR'] > 1 else "#e67e22" if row['CTR'] > 0.3 else "#e74c3c"
            cat_rows += f"""
            <tr>
                <td style="padding:8px;border-bottom:1px solid #f0f0f0">{row['Date']}</td>
                <td style="padding:8px;border-bottom:1px solid #f0f0f0"><b>{row['Campaign Name']}</b></td>
                <td style="padding:8px;border-bottom:1px solid #f0f0f0">{row['Category']}</td>
                <td style="padding:8px;border-bottom:1px solid #f0f0f0">{row['Time']}</td>
                <td style="padding:8px;border-bottom:1px solid #f0f0f0">{row['Audience']}</td>
                <td style="padding:8px;border-bottom:1px solid #f0f0f0">{row['Impressions']:,}</td>
                <td style="padding:8px;border-bottom:1px solid #f0f0f0">{row['Clicks']:,}</td>
                <td style="padding:8px;border-bottom:1px solid #f0f0f0;color:{ctr_color}"><b>{row['CTR']:.2f}%</b></td>
            </tr>"""

        # ── 7-day CTR trend from Google Sheet ────────────
        trend_html = ""
        try:
            creds_dict = json.loads(GOOGLE_CREDS_JSON)
            creds = Credentials.from_service_account_info(
                creds_dict,
                scopes=["https://www.googleapis.com/auth/spreadsheets"]
            )
            gc = gspread.authorize(creds)
            ws = gc.open_by_key(GOOGLE_SHEET_ID).sheet1
            all_data = ws.get_all_records()

            if all_data:
                hist_df = pd.DataFrame(all_data)
                hist_df['All Platform CTR'] = pd.to_numeric(hist_df['All Platform CTR'], errors='coerce')
                hist_df['Date'] = hist_df['Date'].astype(str)

                # Get last 7 unique dates
                dates = sorted(hist_df['Date'].unique())[-7:]

                # CTR trend by Category per date
                trend = hist_df[hist_df['Date'].isin(dates)].groupby(['Date', 'Category'])['All Platform CTR'].mean().reset_index()

                # Build trend table
                categories = trend['Category'].unique()
                trend_rows = ""
                for cat in categories:
                    cat_data = trend[trend['Category'] == cat].set_index('Date')['All Platform CTR']
                    cells = ""
                    for d in dates:
                        val = cat_data.get(d, None)
                        if val is not None:
                            color = "#27ae60" if val > 1 else "#e67e22" if val > 0.3 else "#e74c3c"
                            cells += f'<td style="padding:6px;text-align:center;color:{color}"><b>{val:.2f}%</b></td>'
                        else:
                            cells += '<td style="padding:6px;text-align:center;color:#ccc">—</td>'
                    trend_rows += f'<tr><td style="padding:6px;font-weight:bold">{cat}</td>{cells}</tr>'

                date_headers = "".join([f'<th style="padding:6px;background:#1a1a2e;color:white">{d}</th>' for d in dates])
                trend_html = f"""
                <h3>📈 7-Day CTR Trend by Category</h3>
                <table style="border-collapse:collapse;font-size:13px;width:100%">
                    <tr>
                        <th style="padding:6px;background:#1a1a2e;color:white;text-align:left">Category</th>
                        {date_headers}
                    </tr>
                    {trend_rows}
                </table>"""
        except Exception as e:
            trend_html = f"<p style='color:#999;font-size:12px'>Trend data unavailable: {e}</p>"

        # ── Build full HTML ───────────────────────────────
        html = f"""
        <html><body style="font-family:Arial,sans-serif;max-width:900px">
        <h2>📱 Daily Push Notification Report — {today}</h2>

        <table cellpadding="8" style="border-collapse:collapse;margin-bottom:20px">
            <tr style="background:#f0f0f0"><td><b>Total Campaigns</b></td><td>{total_campaigns}</td></tr>
            <tr><td><b>Total Impressions</b></td><td>{total_impressions:,}</td></tr>
            <tr style="background:#f0f0f0"><td><b>Total Clicks</b></td><td>{total_clicks:,}</td></tr>
            <tr><td><b>Average CTR</b></td><td>{avg_ctr:.2f}%</td></tr>
        </table>

        <h3>📊 Campaign Breakdown</h3>
        <table style="border-collapse:collapse;font-size:13px;width:100%">
            <tr style="background:#1a1a2e;color:white">
                <th style="padding:8px;text-align:left">Date</th>
                <th style="padding:8px;text-align:left">Type</th>
                <th style="padding:8px;text-align:left">Category</th>
                <th style="padding:8px;text-align:left">Time</th>
                <th style="padding:8px;text-align:left">Audience</th>
                <th style="padding:8px;text-align:right">Impressions</th>
                <th style="padding:8px;text-align:right">Clicks</th>
                <th style="padding:8px;text-align:right">CTR</th>
            </tr>
            {cat_rows}
        </table>

        {trend_html}

        <br><p style="color:#999;font-size:12px">Full report attached as Excel.</p>
        </body></html>
        """

        sg_api_key = os.environ.get("SENDGRID_API_KEY")
        with open(excel_path, 'rb') as f:
            encoded = base64.b64encode(f.read()).decode()

        # Support multiple recipients separated by comma
        from sendgrid.helpers.mail import To
        recipients = [To(r.strip()) for r in EMAIL_RECIPIENT.split(',')]

        message = Mail(
            from_email=EMAIL_SENDER,
            to_emails=recipients,
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

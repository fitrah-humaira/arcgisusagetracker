import pandas as pd
import matplotlib
matplotlib.use('Agg') # Critical for Task Scheduler (No-GUI mode)
import matplotlib.pyplot as plt
from datetime import datetime
from arcgis.gis import GIS
import os
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from fpdf import FPDF
from dotenv import load_dotenv

# ==================================
# 1. CONFIGURATION
# ==================================
load_dotenv()

# Configuration (load from environment for safety)
PORTAL_URL = os.getenv('PORTAL_URL', '')
USERNAME = os.getenv('PORTAL_USERNAME', '')
PASSWORD = os.getenv('PORTAL_PASSWORD', '')
OUTPUT_DIR = os.getenv('OUTPUT_DIR', r"C:\GIS")

# Monthly files - reset each month
CURRENT_MONTH = datetime.now().strftime("%Y_%m")
OUTPUT_EXCEL = os.path.join(OUTPUT_DIR, f"AU_Portal_Login_Report_{CURRENT_MONTH}.xlsx")
OUTPUT_PDF = os.path.join(OUTPUT_DIR, f"Monthly_Login_Summary_{CURRENT_MONTH}.pdf")
CHART_IMG = os.path.join(OUTPUT_DIR, "temp_chart.png")

# SMTP Configuration (from environment)
SMTP_HOST = os.getenv('SMTP_HOST', '')
SMTP_PORT = int(os.getenv('SMTP_PORT', '587'))
SENDER_EMAIL = os.getenv('SENDER_EMAIL', '')
SENDER_PWD = os.getenv('SENDER_PWD', '')
RECIPIENT_EMAIL = os.getenv('RECIPIENT_EMAIL', '')

def run_audit():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        
    print(f"--- Starting Audit: {datetime.now()} ---")
    
    # --- STEP 1: FETCH DATA ---
    records = []
    now = datetime.now()
    try:
        # Create GIS connection
        print(f"Connecting to {PORTAL_URL}...")
        gis = GIS(PORTAL_URL, USERNAME, PASSWORD, timeout=60)
        print(f"Connected successfully as {gis.users.me.username}")
        
        # Get all users and check their lastLogin timestamps
        print("Fetching all users...")
        users = gis.users.search(query="*", max_users=10000)
        print(f"Found {len(users)} total users. Checking today's logins...")
        
        user_groups_map = {}
        for user in users:
            try:
                # Capture user groups for later processing
                groups = getattr(user, 'groups', [])
                group_titles = []
                for g in groups:
                    if hasattr(g, 'title'):
                        group_titles.append(g.title)
                user_groups_map[user.username] = group_titles

                last_login = getattr(user, 'lastLogin', None)
                if last_login:
                    dt = datetime.fromtimestamp(last_login / 1000)
                    # Check if login was today
                    if dt.date() == now.date():
                        print(f"  - {user.username} logged in at {dt.strftime('%H:%M:%S')}")
                        records.append({
                            "User": user.username,
                            "Login Time": dt.strftime("%I:%M %p"),
                            "Date": dt.strftime("%Y-%m-%d"),
                            "Month": dt.strftime("%B %Y")
                        })
            except Exception as u_err:
                continue
                
        print(f"Successfully parsed {len(records)} login entries today.")

    except Exception as e:
        print(f"GIS Query Failed: {e}")
        return

    if not records:
        print("No login activity recorded for today. Ending process.")
        return

    # --- STEP 2: DATA MERGING ---
    new_df = pd.DataFrame(records)
    if os.path.exists(OUTPUT_EXCEL):
        try:
            # Safely migrate from older versions
            try:
                existing_df = pd.read_excel(OUTPUT_EXCEL, sheet_name='All Logins')
            except ValueError:
                existing_df = pd.read_excel(OUTPUT_EXCEL, sheet_name='Login History')
                
            # Only combine if it's the same month (data already filtered below)
            # Use 'User' and 'Date', keeping only the absolute newest entry if times differ
            combined_df = pd.concat([existing_df, new_df]).drop_duplicates(subset=['User', 'Date'], keep='last')
        except Exception as e:
            print(f"Error merging data: {e}")
            combined_df = new_df
    else:
        combined_df = new_df

    # Filter to only current month's data for pivot table
    current_month_str = now.strftime("%B %Y")
    combined_df_filtered = combined_df[combined_df['Month'] == current_month_str]
    
    pivot_df = combined_df_filtered.pivot_table(index='User', values='Date', aggfunc='count', fill_value=0)
    pivot_df.columns = ['Logins']
    pivot_df = pivot_df.sort_values(by='Logins', ascending=False)

    # --- STEP 3: EXCEL GENERATION ---
    try:
        with pd.ExcelWriter(OUTPUT_EXCEL, engine='xlsxwriter') as writer:
            new_df.to_excel(writer, index=False, sheet_name="Today's Logins")
            combined_df.to_excel(writer, index=False, sheet_name='All Logins')
            pivot_df.to_excel(writer, sheet_name='Monthly Summary')
            
            wb = writer.book
            header_fmt = wb.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
            
            # Process group-specific tabs
            all_groups = set()
            for groups in user_groups_map.values():
                all_groups.update(groups)
                
            group_sheet_names = []
            for group in sorted(all_groups):
                group_users = [u for u, g_list in user_groups_map.items() if group in g_list]
                if group_users:
                    group_df = combined_df[combined_df['User'].isin(group_users)]
                    if not group_df.empty:
                        # Clean sheet name to avoid Excel errors (max 31 chars, no invalid chars)
                        safe_name = "".join(c for c in group if c not in r'[]:*?/\'').strip()
                        safe_name = safe_name[:31]
                        
                        # Handle duplicate sheet names after truncation
                        original_safe_name = safe_name
                        counter = 1
                        while safe_name in group_sheet_names or safe_name in ["Today's Logins", 'All Logins', 'Monthly Summary']:
                            suffix = f"_{counter}"
                            safe_name = original_safe_name[:31 - len(suffix)] + suffix
                            counter += 1
                            
                        group_sheet_names.append(safe_name)
                        group_df.to_excel(writer, index=False, sheet_name=safe_name)

            sheets_to_format = ["Today's Logins", 'All Logins'] + group_sheet_names
            for sn in sheets_to_format:
                if sn in writer.sheets:
                    ws = writer.sheets[sn]
                    ws.set_column('A:D', 25)
                    for i, col in enumerate(['User', 'Login Time', 'Date', 'Month']):
                        ws.write(0, i, col, header_fmt)
    except PermissionError:
        print("ERROR: Please close the Excel file.")
        return

    # --- STEP 4: PDF & CHART ---
    try:
        plt.figure(figsize=(10, 6))
        pivot_df['Logins'].head(10).plot(kind='barh', color='#1F4E78').invert_yaxis()
        plt.title(f"Top 10 Active Users - {current_month_str}")
        plt.xlabel("Login Count")
        plt.tight_layout()
        plt.savefig(CHART_IMG)
        plt.close()

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(190, 10, f"AU ArcGIS Portal Monthly Login Report - {current_month_str}", ln=True, align='C')
        pdf.image(CHART_IMG, x=15, w=170)
        pdf.ln(10)
        
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(110, 10, "User", border=1, align='C')
        pdf.cell(40, 10, "Logins", border=1, ln=True, align='C')
        pdf.set_font("Arial", size=10)
        for user, logins in pivot_df['Logins'].head(10).items():
            pdf.cell(110, 8, str(user), border=1)
            pdf.cell(40, 8, str(int(logins)), border=1, ln=True, align='C')
        pdf.output(OUTPUT_PDF)
    except Exception as e:
        print(f"Reporting Error: {e}")

    # --- STEP 5: EMAIL DELIVERY ---
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = RECIPIENT_EMAIL
        msg['Subject'] = f"Monthly GIS Login Report - {current_month_str}"
        msg.attach(MIMEText(f"Portal login activity for {current_month_str}. See attached reports.", 'plain'))

        for path in [OUTPUT_EXCEL, OUTPUT_PDF]:
            if os.path.exists(path):
                with open(path, "rb") as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f"attachment; filename={os.path.basename(path)}")
                    msg.attach(part)

        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PWD)
            server.send_message(msg)
        
        if os.path.exists(CHART_IMG): os.remove(CHART_IMG)
        print(f"Monthly report complete. Email sent for {current_month_str}.")
    except Exception as e:
        print(f"Email Delivery Failed: {e}")

if __name__ == "__main__":
    run_audit()

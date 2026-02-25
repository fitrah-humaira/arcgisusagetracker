# GIS Portal Daily Login Audit - Setup Instructions

## Overview
This script automatically queries your ArcGIS Portal and generates daily login reports, creates PDF summaries, and emails them to stakeholders. It can be scheduled to run automatically using Windows Task Scheduler.

## Required Python Libraries

Install these packages in your production Python environment:

```bash
pip install arcgis pandas matplotlib fpdf xlsxwriter
```

### Full Library List (with versions used in development)
- **arcgis** (2.4.2) - ArcGIS API for Python, connects to Portal
- **pandas** (2.3.3) - Data manipulation and Excel export
- **matplotlib** (3.10.8) - Chart generation for reports
- **fpdf** (1.7.2) - PDF report generation
- **xlsxwriter** (3.2.9) - Advanced Excel formatting

Optional (for production monitoring):
```bash
pip install python-dotenv  # For secure credential management
```

## Installation Steps

### 1. Install Python
- Download Python 3.10+ from [python.org](https://www.python.org/)
- During installation, **check "Add Python to PATH"**
- Verify installation:
  ```bash
  python --version
  ```

### 2. Install Required Libraries
```bash
python -m pip install --upgrade pip
pip install arcgis pandas matplotlib fpdf xlsxwriter
```

### 3. Create Output Directory
```bash
mkdir C:\GIS
```

## Configuration

Configuration is read from environment variables for security. Create a `.env` file (do NOT commit to version control) or set system environment variables with the following keys (replace placeholders):

```
PORTAL_URL=https://your-portal-url/portal
PORTAL_USERNAME=your_portal_username
PORTAL_PASSWORD=your_portal_password
OUTPUT_DIR=C:\GIS

SMTP_HOST=smtp.example.com
SMTP_PORT=587
SENDER_EMAIL=reports@example.com
SENDER_PWD=your_smtp_password
RECIPIENT_EMAIL=recipient@example.com
```

The script uses `python-dotenv` (optional) to load `.env` at runtime. Do not store secrets in the repository; prefer OS-level environment variables, a secrets manager, or an encrypted store for production deployments.

## Running in Windows Task Scheduler

### Step 1: Open Task Scheduler
1. Press `Windows Key + R`
2. Type `taskschd.msc` and press Enter
3. Or search for "Task Scheduler" in Start Menu

### Step 2: Create New Task
1. Click **"Create Task"** in the right panel
2. **General Tab:**
   - **Name:** `GIS Portal Daily Login Audit`
   - **Description:** Generates daily login reports and sends via email
   - Check: **"Run with highest privileges"**
   - Check: **"Run whether user is logged in or not"** (important for background execution)

### Step 3: Configure Triggers
1. Click **"Triggers"** tab
2. Click **"New..."**
3. **Begin the task:** Select **"On a schedule"**
4. **Settings:**
   - **Recurrence:** Daily
   - **Start time:** `10:00:00` (adjust to your preferred time)
   - **Every:** 1 day
   - Check: **"Enabled"**

### Step 4: Configure Actions
1. Click **"Actions"** tab
2. Click **"New..."**
3. **Action:** Select **"Start a program"**
4. **Program/script:**
   ```
   C:\Python313\python.exe
   ```
5. **Add arguments:**
   ```
   C:\Users\shaffie\OneDrive - Gamuda Berhad\Desktop\GISAUusage.py
   ```
6. **Start in (optional):**
   ```
   C:\Users\shaffie\OneDrive - Gamuda Berhad\Desktop
   ```

### Step 5: Configure Conditions (Optional)
1. Click **"Conditions"** tab
2. **Power:**
   - Uncheck **"Wake the computer to run this task"** (unless you want it to wake PC)
3. **Network:**
   - Check **"Start the task only if the following network connection is available: Any connection"**

### Step 6: Configure Settings
1. Click **"Settings"** tab
2. Important settings:
   - Check: **"Allow task to be run on demand"**
   - Check: **"If the task fails, restart every: 10 minutes"** (set your preference)
   - Select: **"Stop the task if it runs longer than: 30 minutes"**
   - Check: **"If the running task does not end when requested, force it to stop"**

### Step 7: Set User Account
1. Click **"Change User or Group..."**
2. Enter: `SYSTEM` (to run without user login)
3. Or select your Windows user account for testing

### Step 8: Finish
1. Click **"OK"**
2. You may be prompted for admin password
3. Task is now scheduled!

## Testing the Scheduled Task

### Manual Test Run
1. Open Task Scheduler
2. Find your task in the list
3. Right-click and select **"Run"**
4. Check the output files (monthly files named by year_month):
   - `C:\GIS\AU_Portal_Login_Report_<YYYY_MM>.xlsx` (Excel with login history)
   - `C:\GIS\Monthly_Login_Summary_<YYYY_MM>.pdf` (PDF report with chart)
5. Check your email for the sent report

### View Task Logs
1. In Task Scheduler, click your task
2. In the bottom panel, check **"History"** tab
3. Look for **"Operational"** events showing success/failure

## Troubleshooting

### Issue: "Python not found"
- Verify Python path: Run `where python` in Command Prompt
- If not found, reinstall Python with "Add to PATH" option
- Update the script path in Task Scheduler

### Issue: Script runs but no files created
- Check C:\GIS directory exists
- Verify Portal connectivity and credentials
- Check Task Scheduler history for errors
- Run manually first: `python C:\path\to\GISAUusage.py`

### Issue: Email not sending
- Verify SMTP credentials are correct
- Check firewall/antivirus blocking port 587
- Test manually: `python GISAUusage.py`
- Check email settings in script configuration

### Issue: "Permission denied" errors
- Run Task Scheduler as Administrator
- Ensure SYSTEM account has write access to C:\GIS
- Or use a user account with proper permissions

### Issue: Task doesn't run automatically
- Verify task is **"Enabled"**
- Check trigger time is set correctly
- Verify system date/time is correct
- Restart Windows or Task Scheduler service

## Output Files

After successful execution, you'll have:

1. **AU_Portal_Login_Report.xlsx**
   - Sheet 1: `Login History` - All login records with timestamps
   - Sheet 2: `Monthly Pivot` - User login count by month
   - Formatted header row with blue background

2. **Daily_Login_Summary.pdf**
   - Title page with chart
   - Top 10 most active users table
   - Professional formatting suitable for email/sharing

3. **Email**
   - Sent to configured RECIPIENT_EMAIL
   - Subject: "Daily GIS Login Audit - [DATE]"
   - Both Excel and PDF attached

## Monitoring

### View Task Run History
```
Event Viewer → Windows Logs → System
Filter by Task Scheduler
```

### Verify Email Delivery
Check your email account for reports arriving daily at scheduled time.

### Manual Verification Script
Create a small test script to verify configuration:
```python
from arcgis.gis import GIS
import smtplib
import os

# Test Portal connection (use environment variables)
try:
   gis = GIS(os.getenv('PORTAL_URL'), os.getenv('PORTAL_USERNAME'), os.getenv('PORTAL_PASSWORD'))
   print(f"Portal connection successful: {gis.users.me.username}")
except Exception as e:
   print(f"Portal connection failed: {e}")

# Test email
try:
   server = smtplib.SMTP(os.getenv('SMTP_HOST'), int(os.getenv('SMTP_PORT', '587')))
   server.starttls()
   server.login(os.getenv('SENDER_EMAIL'), os.getenv('SENDER_PWD'))
   print("Email configuration successful")
   server.quit()
except Exception as e:
   print(f"Email configuration failed: {e}")
```

## Production Checklist

- [ ] Python 3.10+ installed with PATH configured
- [ ] All required libraries installed: `pip list`
- [ ] C:\GIS directory created
- [ ] Portal URL, username, password verified
- [ ] SMTP server, sender, and recipient email configured
- [ ] Script tested manually: `python GISAUusage.py`
- [ ] Output files created successfully
- [ ] Email received with attachments
- [ ] Task Scheduler task created
- [ ] Task scheduled with correct time/trigger
- [ ] Task run successfully at least once
- [ ] Task history shows no errors
- [ ] Backup of script saved
- [ ] Documentation shared with team

## Maintenance

### Regular Tasks
- Check task execution weekly in Task Scheduler history
- Review generated reports for accuracy
- Monitor email delivery success
- Update credentials if Portal/Email passwords change

### Updates
- Keep arcgis library updated: `pip install --upgrade arcgis`
- Monitor for new versions of other libraries
- Test updates before deploying to production

## Support

For issues or questions:
1. Check script output manually
2. Review Task Scheduler history
3. Verify configuration in script
4. Check Portal and email credentials
5. Review firewall/network connectivity

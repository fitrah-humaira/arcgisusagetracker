# ArcGIS Portal Monthly Login Audit

ArcGIS Portal Monthly Login Audit — generates monthly login reports, PDF summaries, and email delivery.

## Overview

This repository contains a Python script (`GISAUusage.py`) that automates monthly auditing of ArcGIS Portal logins. It collects user `lastLogin` timestamps, compiles a monthly Excel report, creates a PDF summary chart of the top active users, and emails the reports to stakeholders. Designed to run on Windows Task Scheduler.

## Files

- `GISAUusage.py` — main script that collects logins, writes Excel and PDF reports, and sends email.
- `SETUP_INSTRUCTIONS.md` — detailed installation, Task Scheduler, and troubleshooting instructions.
- `requirements.txt` — minimal production dependencies.

## Quick start

1. Create and activate a Python 3.10+ environment.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

2. Configure credentials in environment variables (recommended) or edit the top of `GISAUusage.py`:

```powershell
# Using a .env file (requires python-dotenv)
setx PORTAL_USERNAME "your_username"
setx PORTAL_PASSWORD "your_portal_password"
setx SMTP_PASSWORD "your_smtp_password"
```

3. Run the script manually to verify:

```powershell
python GISAUusage.py
```

## Scheduling (Task Scheduler)

See `SETUP_INSTRUCTIONS.md` for a step-by-step guide to schedule `GISAUusage.py` using Windows Task Scheduler.

## Security

- Do not commit secrets to source control. Use environment variables or a secured secrets store.
- If you must use a `.env` file, add it to `.gitignore`.

## Pushing to GitHub

If you want to push this repository to GitHub:

```bash
git init
git add .
git commit -m "Initial commit: ArcGIS monthly login audit"
git remote add origin https://github.com/OWNER/REPO.git
git branch -M main
git push -u origin main
```

Replace the remote URL with your repository.

## Support

Refer to `SETUP_INSTRUCTIONS.md` for troubleshooting and production checklist. Open an issue in the repository or contact the script owner for assistance.

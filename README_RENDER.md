# ChildPsych Report Merge App - Render trial deployment

This folder is Render-ready for a test deployment.

Render settings:
- Service type: Web Service
- Runtime: Python 3
- Build command: pip install -r requirements.txt
- Start command: python browser_report_app.py

Important: This is for dummy/test data only. Do not enter real patient/client data on free hosting.


## Added modules

### Administrator settings
The default administrator login is:

- Email: `drdavid@childpsych.ie`
- Password: `childpsych`

From **Settings**, an administrator can add/remove users, reset passwords, activate/deactivate users, and grant access to Report Generation and Invoice Entry.

### Invoice Entry
The Invoice Entry module is a browser-based prototype. It allows a user to upload a PDF invoice, enter key invoice variables, and download:

- a renamed PDF invoice; and
- an `Invoice_Register.xlsx` workbook containing the invoice register row(s).

For now, OneDrive upload remains manual. The VA should upload the renamed PDF and Excel register to OneDrive through the browser. Microsoft Graph/OneDrive automation can be added later.

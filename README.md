# Child Psych Home Page - Report Merge Browser App

This is a local browser-based report merge app.

It opens to a **home page** with two report options:

1. **Medicolegal Report Generation**
   - Uses `Report_Temp.docx`
   - This is the existing report generator already built.

2. **Supplementary Medicolegal Report**
   - Uses `Supplementary_Report_Temp.docx`
   - This is currently a duplicate of the main report template.
   - When you have the supplementary Word template, replace `Supplementary_Report_Temp.docx` with the new file, keeping the same filename.

The app detects merge fields in the form `«Field_name»`, displays a VA-friendly browser form, and generates a completed Word `.docx` file.

## Key features

- Runs locally in your browser at `http://127.0.0.1:8765`
- Starts with a home page showing the two report generators
- Uses the included Word templates
- Auto-calculates pronouns from Sex
- Auto-calculates age and elapsed-time fields
- Uses dropdowns for Sex, country, county, guardian relationships, translator status, and language
- Generates a completed Word report
- Opens Gmail, Outlook Web, or the default email app with recipient, subject and message pre-filled
- Includes a **New Report** button to clear the previously completed form and start a fresh report

## Important email limitation

A local browser app cannot silently attach a generated Word file to Gmail or Outlook Web. This is a browser/webmail security restriction.

Workflow:

1. Open the app home page.
2. Choose **Medicolegal Report Generation** or **Supplementary Medicolegal Report**.
3. Click **New Report** if starting a fresh case.
4. Generate the report.
5. Click Gmail / Outlook Web / default email app.
6. The email draft opens with recipient, subject and body filled in.
7. Download the generated Word report.
8. Manually attach the Word report before sending.

The fixed recipient options are:

- `drdavid@childpsych.ie`
- `info@childpsych.ie`
- `admin@childpsych.ie`

## How to run on Mac

1. Unzip the folder.
2. Open Terminal.
3. Type `cd ` and drag this folder into Terminal, then press Enter.
4. Run:

```bash
python3 browser_report_app.py
```

Then open:

```text
http://127.0.0.1:8765
```

Leave Terminal open while using the app.


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

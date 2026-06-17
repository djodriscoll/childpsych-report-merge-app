#!/usr/bin/env python3
"""
Browser Report Merge App
- Runs locally in a browser at http://127.0.0.1:8765
- Reads Word placeholders in the form «Field_name»
- Creates a web form using clear, VA-friendly labels
- Auto-calculates pronouns and age/time fields
- Generates a populated DOCX
- Generates a completed DOCX and can open Gmail or Outlook Web with a pre-addressed draft.
- Webmail cannot receive an automatic attachment from a local browser page; the generated DOCX must be attached manually.

No third-party Python packages are required.
"""

import html
import json
import os
import re
import sys
import threading
import time
import urllib.parse
import webbrowser
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from pathlib import Path

APP_TITLE = "Child Psych Home Page"
HOSTED_MODE = bool(os.environ.get("RENDER") or os.environ.get("PORT"))
HOST = "0.0.0.0" if HOSTED_MODE else "127.0.0.1"
PORT = int(os.environ.get("PORT", "8765"))
DISPLAY_LOCATION = "Hosted Render trial; do not enter real patient/client data" if HOSTED_MODE else f"Runs locally at http://{HOST}:{PORT}; not uploaded to the internet"
PLACEHOLDER_RE = re.compile(r"«([^»]+)»")
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = BASE_DIR / "Report_Temp.docx"
SUPPLEMENTARY_TEMPLATE = BASE_DIR / "Supplementary_Report_Temp.docx"
OUTPUT_DIR = BASE_DIR / "outputs"
SETTINGS_FILE = Path.home() / ".browser_report_merge_app_settings.json"
DEFAULT_RECIPIENTS = [
    "drdavid@childpsych.ie",
    "info@childpsych.ie",
    "admin@childpsych.ie",
]

FIELD_ORDER = [
    "Ass_date", "Agency", "Forename", "Surname", "DOB", "Claim_no", "Solicitor_name",
    "Date_accident", "Sex", "Pronoun1", "Pronoun2", "Pronoun3", "Age_today", "Age_accident",
    "Time_from_accident_today", "Guardian_relationship", "Guardian_name", "Guardian_attend1_rel",
    "Guardian_attend2_rel", "Add1", "Add2", "Add_country", "Add_county", "Add_eircode",
    "Report1_author", "Reprot1_date", "Report2_author", "Report2_date", "Report3_author",
    "Report3_date", "Report4_author", "Report4_date", "Report5_author", "Report5_date",
    "Report6_author", "Report6_date", "Translator_name", "Language",
]

FIELD_LABELS = {
    "Ass_date": "Assessment date",
    "Agency": "Agency / instructing party",
    "Forename": "Claimant forename",
    "Surname": "Claimant surname",
    "DOB": "Claimant date of birth",
    "Claim_no": "Claim number",
    "Solicitor_name": "Solicitor name",
    "Date_accident": "Date of incident / accident",
    "Sex": "Claimant sex",
    "Pronoun1": "Subject pronoun - auto-filled",
    "Pronoun2": "Object pronoun - auto-filled",
    "Pronoun3": "Possessive pronoun - auto-filled",
    "Age_today": "Age at assessment - auto-calculated",
    "Age_accident": "Age at accident - auto-calculated",
    "Time_from_accident_today": "Time from accident to assessment - auto-calculated",
    "Guardian_relationship": "Guardian relationship to claimant",
    "Guardian_name": "Guardian full name",
    "Guardian_attend1_rel": "Guardian attending 1 relationship",
    "Guardian_attend2_rel": "Guardian attending 2 relationship",
    "Add1": "Address line 1",
    "Add2": "Address line 2 / town",
    "Add_county": "County",
    "Add_country": "Country",
    "Add_eircode": "Eircode / postcode",
    "Report1_author": "Report 1 author",
    "Reprot1_date": "Report 1 date",
    "Report2_author": "Report 2 author",
    "Report2_date": "Report 2 date",
    "Report3_author": "Report 3 author",
    "Report3_date": "Report 3 date",
    "Report4_author": "Report 4 author",
    "Report4_date": "Report 4 date",
    "Report5_author": "Report 5 author",
    "Report5_date": "Report 5 date",
    "Report6_author": "Report 6 author",
    "Report6_date": "Report 6 date",
    "Translator_name": "Translator name",
    "Language": "Translator language",
}

SECTIONS = [
    ("Claimant and assessment details", ["Ass_date", "Agency", "Forename", "Surname", "DOB", "Sex", "Age_today"]),
    ("Claim and solicitor details", ["Claim_no", "Solicitor_name", "Date_accident", "Age_accident", "Time_from_accident_today"]),
    ("Guardian / accompanying adult details", ["Guardian_relationship", "Guardian_name", "Guardian_attend1_rel", "Guardian_attend2_rel"]),
    ("Address", ["Add1", "Add2", "Add_country", "Add_county", "Add_eircode"]),
    ("Pronouns - automatically based on claimant sex", ["Pronoun1", "Pronoun2", "Pronoun3"]),
    ("Reports reviewed", ["Report1_author", "Reprot1_date", "Report2_author", "Report2_date", "Report3_author", "Report3_date", "Report4_author", "Report4_date", "Report5_author", "Report5_date", "Report6_author", "Report6_date"]),
    ("Translator", ["Translator_name", "Language"]),
]

DATE_FIELDS = {"Ass_date", "DOB", "Date_accident", "Reprot1_date", "Report2_date", "Report3_date", "Report4_date", "Report5_date", "Report6_date"}
AUTO_FIELDS = {"Pronoun1", "Pronoun2", "Pronoun3", "Age_today", "Age_accident", "Time_from_accident_today"}

RELATIONSHIPS = [
    "", "mother", "father", "parent", "stepmother", "stepfather", "grandmother", "grandfather",
    "aunt", "uncle", "older sister", "older brother", "foster mother", "foster father", "legal guardian",
    "guardian ad litem", "social worker", "other"
]

LANGUAGES = [
    "", "Arabic", "Chinese - Cantonese", "Chinese - Mandarin", "Czech", "Dutch", "English", "French",
    "German", "Hindi", "Irish", "Italian", "Latvian", "Lithuanian", "Polish", "Portuguese",
    "Romanian", "Russian", "Slovak", "Spanish", "Ukrainian", "Urdu", "Other"
]

IRELAND_COUNTIES = [
    "", "Co.Antrim", "Co.Armagh", "Co.Carlow", "Co.Cavan", "Co.Clare", "Co.Cork", "Co.Derry",
    "Co.Donegal", "Co.Down", "Co.Dublin", "Co.Fermanagh", "Co.Galway", "Co.Kerry", "Co.Kildare",
    "Co.Kilkenny", "Co.Laois", "Co.Leitrim", "Co.Limerick", "Co.Longford", "Co.Louth", "Co.Mayo",
    "Co.Meath", "Co.Monaghan", "Co.Offaly", "Co.Roscommon", "Co.Sligo", "Co.Tipperary", "Co.Tyrone",
    "Co.Waterford", "Co.Westmeath", "Co.Wexford", "Co.Wicklow"
]

UK_COUNTIES = [
    "", "Bedfordshire", "Berkshire", "Bristol", "Buckinghamshire", "Cambridgeshire", "Cheshire",
    "City of London", "Cornwall", "County Antrim", "County Armagh", "County Down", "County Fermanagh",
    "County Londonderry", "County Tyrone", "Cumbria", "Derbyshire", "Devon", "Dorset", "Durham",
    "East Riding of Yorkshire", "East Sussex", "Essex", "Gloucestershire", "Greater London", "Greater Manchester",
    "Hampshire", "Herefordshire", "Hertfordshire", "Isle of Wight", "Kent", "Lancashire", "Leicestershire",
    "Lincolnshire", "Merseyside", "Norfolk", "North Yorkshire", "Northamptonshire", "Northumberland",
    "Nottinghamshire", "Oxfordshire", "Rutland", "Shropshire", "Somerset", "South Yorkshire", "Staffordshire",
    "Suffolk", "Surrey", "Tyne and Wear", "Warwickshire", "West Midlands", "West Sussex", "West Yorkshire",
    "Wiltshire", "Worcestershire", "Aberdeenshire", "Angus", "Argyll and Bute", "Ayrshire", "Clackmannanshire",
    "Dumfries and Galloway", "Dunbartonshire", "Edinburgh", "Fife", "Glasgow", "Highland", "Lanarkshire",
    "Midlothian", "Moray", "Orkney", "Perth and Kinross", "Renfrewshire", "Scottish Borders", "Shetland",
    "Stirling", "West Lothian", "Anglesey", "Blaenau Gwent", "Bridgend", "Caerphilly", "Cardiff", "Carmarthenshire",
    "Ceredigion", "Conwy", "Denbighshire", "Flintshire", "Gwynedd", "Merthyr Tydfil", "Monmouthshire",
    "Neath Port Talbot", "Newport", "Pembrokeshire", "Powys", "Rhondda Cynon Taf", "Swansea", "Torfaen",
    "Vale of Glamorgan", "Wrexham"
]


def load_settings():
    if HOSTED_MODE:
        return {}
    try:
        if SETTINGS_FILE.exists():
            return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def save_settings(settings):
    if HOSTED_MODE:
        return
    clean = {k: v for k, v in settings.items() if k != "smtp_password"}
    SETTINGS_FILE.write_text(json.dumps(clean, indent=2), encoding="utf-8")


def docx_xml_members(zf):
    for name in zf.namelist():
        if name.startswith("word/") and name.endswith(".xml"):
            yield name


def extract_placeholders(docx_path):
    fields = set()
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with zipfile.ZipFile(docx_path, "r") as zf:
        for name in docx_xml_members(zf):
            try:
                root = ET.fromstring(zf.read(name))
            except Exception:
                continue
            for para in root.findall(".//w:p", ns):
                texts = [node.text or "" for node in para.findall(".//w:t", ns)]
                if not texts:
                    continue
                paragraph_text = "".join(texts)
                for match in PLACEHOLDER_RE.findall(paragraph_text):
                    clean = match.strip()
                    if clean and "<" not in clean and ">" not in clean:
                        fields.add(clean)
    ordered = [f for f in FIELD_ORDER if f in fields]
    ordered += sorted(fields.difference(ordered), key=lambda x: x.lower())
    return ordered


def safe_filename_part(value):
    value = (value or "").strip()
    value = re.sub(r"[^A-Za-z0-9 _.-]+", "", value)
    value = re.sub(r"\s+", "_", value)
    return value[:60] or "Report"


def parse_date(value):
    value = (value or "").strip()
    if not value:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return None


def years_months_between(start, end):
    if not start or not end or end < start:
        return ""
    months = (end.year - start.year) * 12 + (end.month - start.month)
    if end.day < start.day:
        months -= 1
    years = months // 12
    rem_months = months % 12
    y_label = "year" if years == 1 else "years"
    m_label = "month" if rem_months == 1 else "months"
    return f"{years} {y_label}, {rem_months} {m_label}"


def computed_values(values):
    values = dict(values)
    sex = values.get("Sex", "")
    if sex == "Male":
        values["Pronoun1"] = "he"
        values["Pronoun2"] = "him"
        values["Pronoun3"] = "his"
    elif sex == "Female":
        values["Pronoun1"] = "she"
        values["Pronoun2"] = "her"
        values["Pronoun3"] = "her"
    else:
        values["Pronoun1"] = ""
        values["Pronoun2"] = ""
        values["Pronoun3"] = ""

    dob = parse_date(values.get("DOB"))
    accident = parse_date(values.get("Date_accident"))
    assessment = parse_date(values.get("Ass_date"))
    values["Age_today"] = years_months_between(dob, date.today())
    values["Age_accident"] = years_months_between(dob, accident)
    # The Word template labels this as time from accident to examination date, so it is calculated as assessment date minus accident date.
    values["Time_from_accident_today"] = years_months_between(accident, assessment)
    return values


def populate_docx(template_path, output_path, values):
    replacements = {f"«{k}»": str(v) for k, v in values.items()}
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    ET.register_namespace("w", ns["w"])

    with zipfile.ZipFile(template_path, "r") as zin:
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                    try:
                        root = ET.fromstring(data)
                        changed = False
                        for para in root.findall(".//w:p", ns):
                            nodes = para.findall(".//w:t", ns)
                            if not nodes:
                                continue
                            combined = "".join(node.text or "" for node in nodes)
                            new_text = combined
                            for placeholder, replacement in replacements.items():
                                new_text = new_text.replace(placeholder, replacement)
                            if new_text != combined:
                                nodes[0].text = new_text
                                nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                                for node in nodes[1:]:
                                    node.text = ""
                                changed = True
                        if changed:
                            data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                    except Exception:
                        text = data.decode("utf-8", errors="ignore")
                        for placeholder, replacement in replacements.items():
                            text = text.replace(placeholder, replacement)
                        data = text.encode("utf-8")
                zout.writestr(item, data)


def webmail_compose_url(provider, recipient, subject, body):
    subject = subject or "Medicolegal report"
    body = body or "Please find attached the completed medicolegal report."
    if provider == "gmail":
        qs = urllib.parse.urlencode({"view": "cm", "fs": "1", "to": recipient, "su": subject, "body": body})
        return "https://mail.google.com/mail/?" + qs
    if provider == "outlook":
        qs = urllib.parse.urlencode({"to": recipient, "subject": subject, "body": body})
        return "https://outlook.office.com/mail/deeplink/compose?" + qs
    qs = urllib.parse.urlencode({"subject": subject, "body": body})
    return f"mailto:{urllib.parse.quote(recipient)}?" + qs


REPORT_CONFIG = {
    "medicolegal": {
        "title": "Medicolegal Report Generation",
        "subtitle": "Generate the standard medicolegal report from the main Word template.",
        "template": DEFAULT_TEMPLATE,
        "filename_prefix": "Medicolegal_Report",
    },
    "supplementary": {
        "title": "Supplementary Medicolegal Report",
        "subtitle": "Generate a supplementary medicolegal report. Replace Supplementary_Report_Temp.docx with your supplementary template when ready.",
        "template": SUPPLEMENTARY_TEMPLATE,
        "filename_prefix": "Supplementary_Medicolegal_Report",
    },
}
FIELD_CACHE = {}
SETTINGS = load_settings()


def normalise_report_type(report_type):
    return report_type if report_type in REPORT_CONFIG else "medicolegal"


def template_for(report_type):
    report_type = normalise_report_type(report_type)
    template = REPORT_CONFIG[report_type]["template"]
    # While the supplementary template is pending, fall back to a duplicate of the main report template.
    if not template.exists() and report_type == "supplementary":
        return DEFAULT_TEMPLATE
    return template


def fields_for(report_type):
    report_type = normalise_report_type(report_type)
    template = template_for(report_type)
    cache_key = f"{report_type}:{template}:{template.stat().st_mtime if template.exists() else 0}"
    if cache_key not in FIELD_CACHE:
        FIELD_CACHE.clear()
        FIELD_CACHE[cache_key] = extract_placeholders(template) if template.exists() else []
    return FIELD_CACHE[cache_key]


def settings_key(report_type, name):
    return f"{normalise_report_type(report_type)}_{name}"


def esc(value):
    return html.escape(str(value or ""), quote=True)


def option_html(options, selected):
    return "".join(f'<option value="{esc(o)}" {"selected" if o == selected else ""}>{esc(o)}</option>' for o in options)


def render_field(f, saved_values):
    value = saved_values.get(f, "")
    label = FIELD_LABELS.get(f, f.replace("_", " "))
    data_field = esc(f)

    if f in AUTO_FIELDS:
        # Hidden value is submitted; visible value is calculated and non-editable.
        return f'''
        <label class="field readonly-field" data-field-wrap="{data_field}"><span>{esc(label)}</span>
          <div class="readonly-display" id="display_{data_field}">{esc(value)}</div>
          <input type="hidden" name="field_{data_field}" id="field_{data_field}" value="{esc(value)}">
        </label>'''

    if f == "Sex":
        return f'''
        <label class="field"><span>{esc(label)}</span>
          <select name="field_{data_field}" id="field_{data_field}">{option_html(["", "Male", "Female"], value)}</select>
        </label>'''

    if f == "Add_country":
        return f'''
        <label class="field"><span>{esc(label)}</span>
          <select name="field_{data_field}" id="field_{data_field}">{option_html(["", "Ireland", "United Kingdom"], value)}</select>
        </label>'''

    if f == "Add_county":
        all_options = sorted(set(IRELAND_COUNTIES + UK_COUNTIES), key=lambda x: x.lower())
        return f'''
        <label class="field"><span>{esc(label)}</span>
          <select name="field_{data_field}" id="field_{data_field}" data-saved="{esc(value)}">{option_html(all_options, value)}</select>
        </label>'''

    if f in {"Guardian_relationship", "Guardian_attend1_rel", "Guardian_attend2_rel"}:
        return f'''
        <label class="field"><span>{esc(label)}</span>
          <select name="field_{data_field}" id="field_{data_field}">{option_html(RELATIONSHIPS, value)}</select>
        </label>'''

    if f == "Language":
        return f'''
        <label class="field translator-detail"><span>{esc(label)}</span>
          <select name="field_{data_field}" id="field_{data_field}">{option_html(LANGUAGES, value)}</select>
        </label>'''

    css_class = "translator-detail" if f == "Translator_name" else ""
    input_type = "date" if f in DATE_FIELDS else "text"
    return f'''
        <label class="field {css_class}"><span>{esc(label)}</span>
          <input type="{input_type}" name="field_{data_field}" id="field_{data_field}" value="{esc(value)}">
        </label>'''


SECTION_HINTS = {
    "Claimant and assessment details": "Core identifying information and assessment date.",
    "Claim and solicitor details": "Claim number, solicitor and accident timing.",
    "Guardian / accompanying adult details": "Who attended and their relationship to the claimant.",
    "Address": "Postal address details used on the report front sheet.",
    "Pronouns - automatically based on claimant sex": "Locked fields; these are auto-filled from the sex dropdown.",
    "Reports reviewed": "Authors and dates of reports reviewed before assessment.",
    "Translator": "Translator details only appear when translator included is Yes.",
}

def home_page(message=""):
    msg_html = f'<div class="message">{esc(message)}</div>' if message else ""
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{APP_TITLE}</title>
<style>
:root{{--navy:#12324a;--blue:#246b9f;--bg:#f3f7fa;--card:#fff;--ink:#1f2937;--muted:#667085;--line:#d8e1e8;}}
*{{box-sizing:border-box}} body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;margin:0;background:linear-gradient(180deg,#edf6fb 0,#f7f9fb 320px);color:var(--ink);}}
header{{background:var(--navy);color:#fff;padding:24px 30px;box-shadow:0 4px 18px rgba(16,24,40,.18);}} .header-inner{{max-width:1120px;margin:0 auto;display:flex;gap:14px;align-items:center}}.logo{{width:52px;height:52px;border-radius:16px;background:linear-gradient(135deg,#fff,#b9dff5);color:var(--navy);display:flex;align-items:center;justify-content:center;font-weight:900}}h1{{margin:0;font-size:25px}}.small{{color:#cde4f4;font-size:13px;margin-top:4px}}
main{{max-width:1120px;margin:34px auto 70px;padding:0 22px}}.intro{{background:rgba(255,255,255,.96);border:1px solid var(--line);border-radius:20px;box-shadow:0 12px 30px rgba(16,24,40,.07);padding:24px;margin-bottom:22px}}.intro h2{{margin:0 0 8px;font-size:22px}}.intro p{{margin:0;color:var(--muted);line-height:1.5}}.tile-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:18px}}.tile{{background:#fff;border:1px solid var(--line);border-radius:20px;box-shadow:0 12px 30px rgba(16,24,40,.07);padding:24px;text-decoration:none;color:var(--ink);display:block;transition:transform .12s ease,box-shadow .12s ease,border-color .12s ease}}.tile:hover{{transform:translateY(-2px);box-shadow:0 16px 38px rgba(16,24,40,.12);border-color:#9bc8e6}}.tile .badge{{display:inline-block;background:#eef4ff;color:#175cd3;padding:5px 10px;border-radius:999px;font-size:12px;font-weight:800;margin-bottom:12px}}.tile h3{{margin:0 0 8px;font-size:20px}}.tile p{{margin:0 0 18px;color:var(--muted);line-height:1.45}}.button{{display:inline-block;background:var(--blue);color:#fff;border:none;padding:11px 15px;border-radius:11px;text-decoration:none;font-weight:800;box-shadow:0 4px 10px rgba(36,107,159,.2)}}.message{{background:#ecfdf3;border:1px solid #abefc6;color:#05603a;padding:13px 15px;border-radius:14px;margin-bottom:16px}}.footer-note{{margin:24px 4px 0;color:var(--muted);font-size:12px}}
</style>
</head>
<body>
<header><div class="header-inner"><div class="logo">CP</div><div><h1>Child Psych Home Page</h1><div class="small">{esc(DISPLAY_LOCATION)}</div></div></div></header>
<main>
{msg_html}
<section class="intro"><h2>Report Merge Browser App</h2><p>Select which report you want to generate. Each option opens a form, fills the matching Word template, and creates a completed Word document on this computer.</p></section>
<section class="tile-grid">
  <a class="tile" href="/medicolegal"><span class="badge">Main report</span><h3>Medicolegal Report Generation</h3><p>Open the existing medicolegal report generation form using the current Report_Temp.docx template.</p><span class="button">Open medicolegal report generator</span></a>
  <a class="tile" href="/supplementary"><span class="badge">Supplementary</span><h3>Supplementary Medicolegal Report</h3><p>Open a duplicate report generator for a supplementary medicolegal report. Replace the supplementary Word template when ready.</p><span class="button">Open supplementary report generator</span></a>
</section>
<p class="footer-note">Trial tool only. Do not enter real patient/client data on free hosting. Templates and generated reports remain within the running app environment unless you email, upload or share them.</p>
</main>
</body>
</html>"""


def page(report_type="medicolegal", message="", download_file=""):
    report_type = normalise_report_type(report_type)
    config = REPORT_CONFIG[report_type]
    fields = fields_for(report_type)
    email_subject = SETTINGS.get(settings_key(report_type, "email_subject"), "Medicolegal report")
    saved_values = SETTINGS.get(settings_key(report_type, "last_values"), {})
    saved_extra = SETTINGS.get(settings_key(report_type, "last_extra"), {})
    recipient_saved = SETTINGS.get(settings_key(report_type, "recipient"), DEFAULT_RECIPIENTS[0])
    translator_included = saved_extra.get("translator_included", "No")

    rendered_fields = {}
    for f in fields:
        rendered_fields[f] = render_field(f, saved_values)

    sections_html = []
    shown = set()
    for section_title, section_fields in SECTIONS:
        visible = [rendered_fields[f] for f in section_fields if f in rendered_fields]
        if not visible:
            continue
        shown.update([f for f in section_fields if f in rendered_fields])
        extra = ""
        if section_title == "Translator":
            extra = f'''
            <label class="field"><span>Translator included?</span>
              <select name="translator_included" id="translator_included">{option_html(["No", "Yes", "Maybe"], translator_included)}</select>
            </label>'''
        sections_html.append(f'<section class="card"><div class="card-head"><div><h2>{esc(section_title)}</h2><p>{esc(SECTION_HINTS.get(section_title, ""))}</p></div><span class="pill">Section</span></div><div class="grid">{extra}{"".join(visible)}</div></section>')

    other_fields = [rendered_fields[f] for f in fields if f not in shown]
    if other_fields:
        sections_html.append(f'<section class="card"><div class="card-head"><div><h2>Other template fields</h2><p>Additional merge fields detected in the template.</p></div><span class="pill optional">Optional</span></div><div class="grid">{"".join(other_fields)}</div></section>')

    recipient_options = "".join(
        f'<option value="{esc(r)}" {"selected" if r == recipient_saved else ""}>{esc(r)}</option>'
        for r in DEFAULT_RECIPIENTS
    )
    download = ""
    if download_file:
        download = f'<section class="success-panel"><div><h2>Report generated successfully</h2><p>Filename: <code>{esc(download_file)}</code></p><p class="small">Download the Word report and attach it before sending.</p></div><div class="actions"><a class="button" href="/download?file={urllib.parse.quote(download_file)}">Download Word Report</a><a class="button secondary" href="/">Home</a><a class="button secondary" href="/new?type={report_type}">New Report</a></div></section>'

    msg_html = f'<div class="message">{esc(message)}</div>' if message else ""
    ireland_json = json.dumps(IRELAND_COUNTIES)
    uk_json = json.dumps(UK_COUNTIES)
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(config["title"])} - {APP_TITLE}</title>
<style>
:root{{--navy:#12324a;--blue:#246b9f;--bg:#f3f7fa;--card:#fff;--ink:#1f2937;--muted:#667085;--line:#d8e1e8;}}
*{{box-sizing:border-box}} body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;margin:0;background:linear-gradient(180deg,#edf6fb 0,#f7f9fb 280px);color:var(--ink);}}
header{{background:var(--navy);color:#fff;padding:18px 30px;position:sticky;top:0;z-index:10;box-shadow:0 4px 18px rgba(16,24,40,.18);}} .header-inner{{max-width:1180px;margin:0 auto;display:flex;justify-content:space-between;align-items:center;gap:16px}}.brand{{display:flex;gap:14px;align-items:center}}.logo{{width:46px;height:46px;border-radius:14px;background:linear-gradient(135deg,#fff,#b9dff5);color:var(--navy);display:flex;align-items:center;justify-content:center;font-weight:900}}.top-actions{{display:flex;gap:10px;flex-wrap:wrap}}
main{{max-width:1180px;margin:26px auto 70px;padding:0 22px;background:transparent;box-shadow:none;border-radius:0}} h1{{margin:0;font-size:22px}}.small{{color:#cde4f4;font-size:13px}}.intro{{display:grid;grid-template-columns:1.4fr .8fr;gap:18px;margin-bottom:18px}}.panel,.card{{background:rgba(255,255,255,.96);border:1px solid var(--line);border-radius:18px;box-shadow:0 12px 30px rgba(16,24,40,.07)}}.panel{{padding:22px}}.panel h2{{margin:0 0 8px;font-size:20px}}.panel p{{margin:0;color:var(--muted);font-size:13px;line-height:1.45}}.progress-label{{display:flex;justify-content:space-between;color:var(--muted);font-size:13px;margin-bottom:8px}}.progress-track{{height:12px;background:#e6edf2;border-radius:999px;overflow:hidden}}.progress-bar{{height:100%;width:0%;background:linear-gradient(90deg,var(--blue),#3a9ed4);transition:width .2s ease}}.progress-note{{margin-top:10px!important;color:var(--muted)!important}}
.card{{margin:14px 0;overflow:hidden}}.card-head{{display:flex;justify-content:space-between;align-items:center;gap:12px;padding:18px 20px;border-bottom:1px solid #eef2f5}}.card-head h2{{margin:0;font-size:18px;border:0;padding:0}}.card-head p{{margin:4px 0 0;color:var(--muted);font-size:12.5px}}.pill{{white-space:nowrap;background:#eef4ff;color:#175cd3;padding:5px 10px;border-radius:999px;font-size:12px;font-weight:800}}.pill.optional{{background:#f2f4f7;color:#475467}}.grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:16px 18px;padding:20px}}
.field span{{display:block;font-weight:700;font-size:13px;margin-bottom:5px;color:#344054}} input,select,textarea{{width:100%;padding:10px 11px;border:1px solid #cfd8e3;border-radius:10px;font-size:14px;background:#fff;outline:none;transition:border-color .12s ease,box-shadow .12s ease}} input:focus,select:focus,textarea:focus{{border-color:var(--blue);box-shadow:0 0 0 4px rgba(36,107,159,.12)}} textarea{{min-height:92px;resize:vertical}}.readonly-display{{width:100%;padding:10px 11px;border:1px solid #d6dde5;border-radius:10px;min-height:41px;background:#f2f5f8;color:#344054}}.readonly-field span:after{{content:' 🔒';font-weight:400}} small{{display:block;color:var(--muted);font-size:12px;margin-top:5px}}
.actions{{margin-top:18px;display:flex;gap:10px;flex-wrap:wrap}}button,.button{{background:var(--blue);color:#fff;border:none;padding:11px 15px;border-radius:11px;text-decoration:none;font-weight:800;cursor:pointer;box-shadow:0 4px 10px rgba(36,107,159,.2)}}button.secondary,.button.secondary{{background:#475467}}.message{{background:#ecfdf3;border:1px solid #abefc6;color:#05603a;padding:13px 15px;border-radius:14px;margin-bottom:16px}}.note{{background:#fffaeb;border:1px solid #fedf89;color:#7a2e0e;padding:13px 14px;border-radius:14px;margin:14px 20px}}.success-panel{{display:flex;justify-content:space-between;align-items:center;background:#ecfdf3;border:1px solid #abefc6;border-radius:18px;padding:18px 20px;margin:16px 0}}.success-panel h2{{margin:0 0 6px;font-size:18px}}.hidden{{display:none!important}}.email-box{{padding:0}}.email-grid{{display:grid;grid-template-columns:1fr 1fr;gap:16px 18px;padding:20px}}.email-grid .wide{{grid-column:1/-1}}.footer-note{{margin:24px 4px 0;color:var(--muted);font-size:12px}}@media(max-width:820px){{.header-inner,.intro,.email-grid{{display:block}}.top-actions{{margin-top:14px}}.success-panel{{display:block}}.success-panel .button{{display:inline-block;margin-top:12px}}}}
</style>
</head>
<body>
<header><div class="header-inner"><div class="brand"><div class="logo">CP</div><div><h1>Child Psych Home Page - Report Merge Browser App</h1><div class="small">{esc(config["title"])} · {esc(DISPLAY_LOCATION)}</div></div></div><div class="top-actions"><a class="button secondary" href="/">Home</a><a class="button secondary" href="/new?type={report_type}">New Report</a><button form="reportForm" type="submit" name="action" value="generate">Generate Report</button><button form="reportForm" type="reset" class="secondary">Clear Form</button></div></div></header>
<main>
{msg_html}
{download}
<section class="intro"><div class="panel"><h2>{esc(config["title"])}</h2><p>{esc(config["subtitle"])} Enter the report merge details below. Pronouns and calculated date fields are locked and update automatically.</p></div><div class="panel"><div class="progress-label"><strong>Completion</strong><span id="completionText">0 fields completed</span></div><div class="progress-track"><div class="progress-bar" id="completionBar"></div></div><p class="progress-note">Use the section cards below to work through the report in order. Click New Report to clear previous case details.</p></div></section>
<form id="reportForm" method="post" action="/generate">
  <input type="hidden" name="report_type" value="{esc(report_type)}">
  {''.join(sections_html)}

  <section class="card email-box">
    <div class="card-head"><div><h2>Email / Output</h2><p>Generate the Word report and optionally open a prepared email draft.</p></div><span class="pill">Final step</span></div>
    <div class="email-grid">
      <label class="field"><span>Send to</span><select name="recipient">{recipient_options}</select></label>
      <label class="field"><span>Email subject</span><input name="email_subject" value="{esc(email_subject)}"></label>
      <label class="field wide"><span>Email message</span><textarea name="email_body">Please find attached the completed medicolegal report.</textarea></label>
    </div>
    <div class="note">The app can open Gmail or Outlook Web with the recipient, subject and message filled in. Browsers do not allow a local app to silently attach a generated Word file to webmail, so attach the downloaded report manually before sending.</div>
    <div class="actions" style="padding:0 20px 20px;">
      <button type="submit" name="action" value="generate">Generate Word report only</button>
      <button type="submit" name="action" value="gmail">Generate and open Gmail draft</button>
      <button type="submit" name="action" value="outlook">Generate and open Outlook Web draft</button>
      <button type="submit" name="action" value="mailto" class="secondary">Generate and open default email app</button>
    </div>
  </section>
</form>
<p class="footer-note">Trial tool only. Do not enter real patient/client data on free hosting. Generated reports are temporary and should be downloaded immediately.</p>
</main>
<script>
const irelandCounties = {ireland_json};
const ukCounties = {uk_json};

function field(id) {{ return document.getElementById('field_' + id); }}
function setAuto(id, value) {{
  const input = field(id);
  const display = document.getElementById('display_' + id);
  if (input) input.value = value || '';
  if (display) display.textContent = value || '';
}}
function parseDate(value) {{
  if (!value) return null;
  const d = new Date(value + 'T00:00:00');
  return isNaN(d.getTime()) ? null : d;
}}
function yearsMonthsBetween(start, end) {{
  if (!start || !end || end < start) return '';
  let months = (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth());
  if (end.getDate() < start.getDate()) months -= 1;
  const years = Math.floor(months / 12);
  const remMonths = months % 12;
  return `${{years}} ${{years === 1 ? 'year' : 'years'}}, ${{remMonths}} ${{remMonths === 1 ? 'month' : 'months'}}`;
}}
function updatePronounsAndAges() {{
  const sex = field('Sex') ? field('Sex').value : '';
  if (sex === 'Male') {{
    setAuto('Pronoun1', 'he'); setAuto('Pronoun2', 'him'); setAuto('Pronoun3', 'his');
  }} else if (sex === 'Female') {{
    setAuto('Pronoun1', 'she'); setAuto('Pronoun2', 'her'); setAuto('Pronoun3', 'her');
  }} else {{
    setAuto('Pronoun1', ''); setAuto('Pronoun2', ''); setAuto('Pronoun3', '');
  }}
  const dob = parseDate(field('DOB') ? field('DOB').value : '');
  const accident = parseDate(field('Date_accident') ? field('Date_accident').value : '');
  const assessment = parseDate(field('Ass_date') ? field('Ass_date').value : '');
  setAuto('Age_today', yearsMonthsBetween(dob, new Date()));
  setAuto('Age_accident', yearsMonthsBetween(dob, accident));
  setAuto('Time_from_accident_today', yearsMonthsBetween(accident, assessment));
}}
function updateCountyOptions() {{
  const countryEl = field('Add_country');
  const countyEl = field('Add_county');
  if (!countryEl || !countyEl) return;
  const existing = countyEl.value || countyEl.dataset.saved || '';
  let options = [''];
  if (countryEl.value === 'Ireland') options = irelandCounties;
  else if (countryEl.value === 'United Kingdom') options = ukCounties;
  countyEl.innerHTML = '';
  options.forEach(opt => {{
    const el = document.createElement('option');
    el.value = opt; el.textContent = opt;
    if (opt === existing) el.selected = true;
    countyEl.appendChild(el);
  }});
}}
function updateTranslatorVisibility() {{
  const sel = document.getElementById('translator_included');
  const show = sel && sel.value === 'Yes';
  document.querySelectorAll('.translator-detail').forEach(el => {{ el.classList.toggle('hidden', !show); }});
  if (!show) {{
    if (field('Translator_name')) field('Translator_name').value = '';
    if (field('Language')) field('Language').value = '';
  }}
}}
['Sex', 'DOB', 'Date_accident', 'Ass_date'].forEach(id => {{ if (field(id)) field(id).addEventListener('change', updatePronounsAndAges); }});
if (field('Add_country')) field('Add_country').addEventListener('change', updateCountyOptions);
const translatorSel = document.getElementById('translator_included');
if (translatorSel) translatorSel.addEventListener('change', updateTranslatorVisibility);
updateCountyOptions();
updatePronounsAndAges();
updateTranslatorVisibility();

function updateCompletion() {{
  const fields = Array.from(document.querySelectorAll('input[type="text"], input[type="date"], select'));
  const visibleFields = fields.filter(el => el.offsetParent !== null && !el.name.includes('Pronoun') && !el.name.includes('Age_') && !el.name.includes('Time_from'));
  const completed = visibleFields.filter(el => (el.value || '').trim() !== '').length;
  const total = visibleFields.length || 1;
  const pct = Math.round((completed / total) * 100);
  const t = document.getElementById('completionText');
  const b = document.getElementById('completionBar');
  if (t) t.textContent = `${{completed}} of ${{total}} visible fields completed`;
  if (b) b.style.width = pct + '%';
}}
document.querySelectorAll('input, select, textarea').forEach(el => el.addEventListener('input', updateCompletion));
updateCompletion();
</script>
</body>
</html>"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        return

    def send_html(self, content, status=200):
        data = content.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        qs = urllib.parse.parse_qs(parsed.query)
        if parsed.path == "/":
            self.send_html(home_page())
            return
        if parsed.path in {"/medicolegal", "/supplementary"}:
            report_type = parsed.path.strip("/")
            self.send_html(page(report_type))
            return
        if parsed.path == "/new":
            report_type = normalise_report_type(qs.get("type", ["medicolegal"])[0])
            SETTINGS.pop(settings_key(report_type, "last_values"), None)
            SETTINGS.pop(settings_key(report_type, "last_extra"), None)
            save_settings(SETTINGS)
            self.send_html(page(report_type, "New report started. Previous form values have been cleared."))
            return
        if parsed.path == "/download":
            filename = os.path.basename(qs.get("file", [""])[0])
            path = OUTPUT_DIR / filename
            if not filename or not path.exists():
                self.send_html(page("medicolegal", "File not found."), 404)
                return
            data = path.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        self.send_html(home_page())

    def do_POST(self):
        if self.path != "/generate":
            self.send_html(home_page("Unknown action."), 404)
            return
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length).decode("utf-8", errors="replace")
        form = urllib.parse.parse_qs(raw)
        get = lambda name, default="": form.get(name, [default])[0]
        report_type = normalise_report_type(get("report_type", "medicolegal"))
        fields = fields_for(report_type)
        values = {f: get(f"field_{f}") for f in fields}
        values = computed_values(values)
        translator_included = get("translator_included", "No")
        if translator_included != "Yes":
            values["Translator_name"] = ""
            values["Language"] = ""
        recipient = get("recipient", DEFAULT_RECIPIENTS[0])
        if recipient not in DEFAULT_RECIPIENTS:
            self.send_html(page(report_type, "Invalid recipient selected."), 400)
            return
        OUTPUT_DIR.mkdir(exist_ok=True)
        claimant = safe_filename_part((values.get("Surname", "") + "_" + values.get("Forename", "")).strip("_"))
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"{REPORT_CONFIG[report_type]['filename_prefix']}_{claimant}_{timestamp}.docx"
        output_path = OUTPUT_DIR / output_file
        try:
            populate_docx(template_for(report_type), output_path, values)
            if not HOSTED_MODE:
                SETTINGS.update({
                    settings_key(report_type, "recipient"): recipient,
                    settings_key(report_type, "email_subject"): get("email_subject", "Medicolegal report"),
                    settings_key(report_type, "last_values"): values,
                    settings_key(report_type, "last_extra"): {"translator_included": translator_included},
                })
                save_settings(SETTINGS)
            action = get("action")
            if action in {"gmail", "outlook", "mailto"}:
                provider = {"gmail": "gmail", "outlook": "outlook", "mailto": "mailto"}[action]
                body = get("email_body") + "\n\nAttachment to add manually: " + output_file
                subject = get("email_subject", "Medicolegal report")
                url = webmail_compose_url(provider, recipient, subject, body)
                provider_name = {"gmail": "Gmail", "outlook": "Outlook Web", "mailto": "your default email app"}[action]
                html_page = page(report_type, f"Report generated. Open {provider_name} with a draft addressed to {recipient}. Download the Word report and attach it before sending.", output_file)
                draft_panel = f'<section class="success-panel"><div><h2>Email draft ready</h2><p>Open the prepared draft, then manually attach the generated Word report before sending.</p></div><a class="button" target="_blank" rel="noopener" href="{html.escape(url, quote=True)}">Open {html.escape(provider_name)} draft</a></section>'
                html_page = html_page.replace("</main>", draft_panel + "</main>")
                html_page = html_page.replace("</body>", f'<script>window.open({json.dumps(url)}, "_blank");</script></body>')
                self.send_html(html_page)
            else:
                self.send_html(page(report_type, "Report generated successfully.", output_file))
        except Exception as e:
            self.send_html(page(report_type, f"Error: {e}"), 500)


def open_browser_later():
    time.sleep(1)
    webbrowser.open(f"http://{HOST}:{PORT}/")


def main():
    if not DEFAULT_TEMPLATE.exists():
        print(f"Cannot find template: {DEFAULT_TEMPLATE}")
        sys.exit(1)
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"\n{APP_TITLE} is running.")
    if HOSTED_MODE:
        print(f"Render is running the app on port {PORT}.")
    else:
        print(f"Open this address in your browser: http://{HOST}:{PORT}/")
        print("Leave this Terminal window open while using the app.")
        print("Press Ctrl+C in this Terminal window to stop it.\n")
        threading.Thread(target=open_browser_later, daemon=True).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping app.")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
drive_pull.py — download Cashify source files from Drive using a Google service
account, so the dashboard can refresh with no manual steps.

AUTH
----
Set env var GOOGLE_SERVICE_ACCOUNT_JSON to the service-account key JSON (the whole
file contents). The folder must be shared (Viewer) with that service account's
email. Scope used: drive.readonly.

USAGE
    python3 drive_pull.py
    CASHIFY_ROOT_FOLDER_ID=... python3 drive_pull.py

Downloads:
  * mis.xlsx                = MIS file with the latest month in filename
  * model.xlsx              = latest file containing "Model" in Growth Huddle folder
  * captable.pdf            = latest PDF/table file in Cap Tables folder
  * investment_profile.xlsx = latest Investment Profile / valuation file in Cap Tables folder
  * growth_huddle.xlsx      = latest Growth Huddle workbook
Handles both uploaded files and native Google Sheets (exported to .xlsx).
"""
import os, io, json, re, sys
from datetime import datetime
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaIoBaseDownload

COMPANY_NAME = os.environ.get("COMPANY_NAME", "Ultrahuman")
ROOT_FOLDER_ID = os.environ.get("CASHIFY_ROOT_FOLDER_ID", os.environ.get("DRIVE_FOLDER_ID", "1TlP_wE_sZ15qi0qdod0NUssqXcC50e-6"))
SHEET_MIME = "application/vnd.google-apps.spreadsheet"
XLSX_MIME  = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

FOLDER_ENV = {
    "mis": os.environ.get("CASHIFY_MIS_FOLDER_ID"),
    "growth": os.environ.get("CASHIFY_GROWTH_HUDDLE_FOLDER_ID"),
    "cap": os.environ.get("CASHIFY_CAP_TABLES_FOLDER_ID"),
}

FILE_ENV = {
    "mis": os.environ.get("CASHIFY_MIS_FILE_ID"),
    "model": os.environ.get("CASHIFY_MODEL_FILE_ID"),
    "growth": os.environ.get("CASHIFY_GROWTH_HUDDLE_FILE_ID"),
    "investment": os.environ.get("CASHIFY_INVESTMENT_PROFILE_FILE_ID"),
}

def env_truthy(name):
    return str(os.environ.get(name, "")).strip().lower() in {"1", "true", "yes", "y"}

def client():
    raw = (
        os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
        or os.environ.get("GDRIVE_CREDENTIALS")
        or os.environ.get("G_DRIVESECRET")
    )
    if not raw:
        sys.exit("ERROR: set GOOGLE_SERVICE_ACCOUNT_JSON, GDRIVE_CREDENTIALS or G_DRIVESECRET (service-account key contents).")
    info = json.loads(raw)
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive.readonly"])
    return build("drive", "v3", credentials=creds, cache_discovery=False)

def list_files(svc, folder_id, extra_q="", page_size=100):
    q = f"'{folder_id}' in parents and trashed=false"
    if extra_q:
        q += f" and ({extra_q})"
    res = svc.files().list(q=q, orderBy="modifiedTime desc", pageSize=page_size,
                           fields="files(id,name,mimeType,modifiedTime)",
                           supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    return res.get("files", [])

def get_file(svc, file_id):
    return svc.files().get(
        fileId=file_id,
        fields="id,name,mimeType,modifiedTime,parents",
        supportsAllDrives=True,
    ).execute()

def find_child_folder(svc, root_id, *needles):
    folders = list_files(svc, root_id, "mimeType='application/vnd.google-apps.folder'")
    for f in folders:
        name = f["name"].lower()
        if all(n.lower() in name for n in needles):
            return f["id"]
    seen = ", ".join(f["name"] for f in folders) or "none"
    sys.exit(f"ERROR: could not find folder containing {needles} under {root_id}. Seen: {seen}")

MONTHS = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}

def month_key(name):
    low = name.lower()
    month = None
    for token, n in MONTHS.items():
        if re.search(rf"(^|[^a-z]){token}([^a-z]|$)", low):
            month = n
            break
    if not month:
        return None
    years = re.findall(r"(20\d{2}|'\d{2}|[^0-9](\d{2})(?=[^0-9]|$))", name)
    year = None
    for full, short in years:
        raw = full if full.startswith("20") else short or full.replace("'", "")
        if raw:
            year = int(raw) if len(raw) == 4 else 2000 + int(raw)
    if not year:
        return None
    return (year, month)

def modified_key(f):
    try:
        return datetime.fromisoformat(f["modifiedTime"].replace("Z", "+00:00"))
    except Exception:
        return datetime.min

def newest_by_modified(svc, folder_id, predicate, label):
    visible = list_files(svc, folder_id)
    files = [f for f in visible if predicate(f)]
    if not files:
        names = ", ".join(f["name"] for f in visible) or "none"
        sys.exit(f"ERROR: no {label} found in folder {folder_id}. Seen: {names}")
    return sorted(files, key=modified_key, reverse=True)[0]

def latest_mis_by_filename_month(svc, folder_id):
    candidates = []
    for f in list_files(svc, folder_id):
        if "mis" not in f["name"].lower():
            continue
        mk = month_key(f["name"])
        if mk:
            candidates.append((mk, modified_key(f), f))
    if not candidates:
        names = ", ".join(f["name"] for f in list_files(svc, folder_id))
        sys.exit(f"ERROR: no MIS file with parseable month/year in filename. Seen: {names}")
    ranked = sorted(candidates, key=lambda x: (x[0], x[1]), reverse=True)
    print("  MIS candidates visible to service account:")
    for mk, mod, f in ranked[:8]:
        print(f"    {mk[0]:04d}-{mk[1]:02d} · {f['name']} · modified {f.get('modifiedTime')}")
    return ranked[0][2]

def download(svc, f, out_path):
    if f["mimeType"] == SHEET_MIME:
        req = svc.files().export_media(fileId=f["id"], mimeType=XLSX_MIME)
    else:
        req = svc.files().get_media(fileId=f["id"], supportsAllDrives=True)
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = dl.next_chunk()
    with open(out_path, "wb") as fh:
        fh.write(buf.getvalue())
    print(f"  ↓ {f.get('name', out_path)}  -> {out_path}")

def main():
    svc = client()
    mis_folder = FOLDER_ENV["mis"] or find_child_folder(svc, ROOT_FOLDER_ID, "mis")
    growth_folder = FOLDER_ENV["growth"] or find_child_folder(svc, ROOT_FOLDER_ID, "growth", "huddle")
    cap_folder = FOLDER_ENV["cap"] or find_child_folder(svc, ROOT_FOLDER_ID, "cap")
    print(f"  folders: mis={mis_folder} growth={growth_folder} cap={cap_folder}")

    if FILE_ENV["mis"] and env_truthy("CASHIFY_FORCE_MIS_FILE_ID"):
        mis = get_file(svc, FILE_ENV["mis"])
        print(f"  using exact MIS file override: {mis['name']} ({mis['id']})")
    else:
        if FILE_ENV["mis"]:
            print("  ignoring CASHIFY_MIS_FILE_ID because CASHIFY_FORCE_MIS_FILE_ID is not true; selecting latest MIS by filename month")
        mis = latest_mis_by_filename_month(svc, mis_folder)
    if FILE_ENV["model"]:
        model = get_file(svc, FILE_ENV["model"])
        print(f"  using exact model file: {model['name']} ({model['id']})")
    else:
        model = newest_by_modified(svc, growth_folder, lambda f: "model" in f["name"].lower(), "model workbook")
    if FILE_ENV["growth"]:
        growth = get_file(svc, FILE_ENV["growth"])
        print(f"  using exact Growth Huddle file: {growth['name']} ({growth['id']})")
    else:
        growth = newest_by_modified(svc, growth_folder, lambda f: "growth" in f["name"].lower() and "huddle" in f["name"].lower(), "growth huddle workbook")
    captable = newest_by_modified(svc, cap_folder, lambda f: f["mimeType"] == "application/pdf" or "cap" in f["name"].lower(), "cap table PDF")
    if FILE_ENV["investment"]:
        investment = get_file(svc, FILE_ENV["investment"])
        print(f"  using exact investment profile file: {investment['name']} ({investment['id']})")
    else:
        investment = newest_by_modified(
            svc,
            cap_folder,
            lambda f: any(token in f["name"].lower() for token in ("investment", "valuation", "profile")),
            "investment profile workbook",
        )

    download(svc, mis, "mis.xlsx")
    download(svc, model, "model.xlsx")
    download(svc, growth, "growth_huddle.xlsx")
    download(svc, captable, "captable.pdf")
    download(svc, investment, "investment_profile.xlsx")
    print(f"✓ pulled latest {COMPANY_NAME} source files from Drive")

if __name__ == "__main__":
    main()

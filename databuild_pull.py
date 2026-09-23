"""Databuild (dbonline.co.za) project-data puller.

Logs into your Databuild account, pulls approved projects from the internal
dashboard API, and writes a tidy dataset (xlsx/csv/json) for use as grounding
by Microsoft 365 Copilot agents.

Credentials are read from a local .env file only - never from the command line
or hardcoded. See README.md for Windows setup.

Usage:
    python databuild_pull.py                 # normal pull
    python databuild_pull.py --probe         # discover endpoints first
    python databuild_pull.py --all           # include non-approved statuses
    python databuild_pull.py --output-dir C:/path/to/folder
    python databuild_pull.py --dotenv C:/path/to/.env
"""

import argparse
import csv
import datetime as dt
import json
import os
import re
import sys
from pathlib import Path

import requests

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None

BASE_URL = "https://www.dbonline.co.za"
LOGIN_URL = f"{BASE_URL}/Account/LogIn"
DEFAULT_RETURN = "/api/Project/Dashboard"

PROJECT_ENDPOINTS = [
    "Dashboard",
    "List",
    "GetProjects",
    "ApprovedProjects",
    "Approved",
    "Search",
    "SearchProjects",
    "GetAll",
    "Projects",
    "Data",
    "GridData",
    "GetProjectList",
    "Index",
]

FIELD_SYNONYMS = {
    "ref": ["ProjectRef", "ProjectReference", "Reference", "Ref", "ProjectId", "Id", "ProjectNo"],
    "name": ["ProjectName", "Name", "Title", "Project", "Description"],
    "client": ["Client", "ClientName", "Owner", "OwnerName", "ClientOwner", "Employer"],
    "location": ["Location", "Region", "Town", "City", "Province", "Area"],
    "sector": ["Sector", "Category", "Discipline", "Industry"],
    "status": ["Status", "ProjectStatus", "StatusCode", "TenderStatus"],
    "approved_date": ["DateApproved", "ApprovedDate", "ApprovalDate", "AwardDate", "AwardedDate", "DateAwarded"],
    "stage": ["Stage", "StageName", "CurrentStage", "Phase"],
    "value": ["EstimatedValue", "Value", "ProjectValue", "TenderValue", "Amount", "EstimatedCost"],
    "contact": ["Contact", "ContactPerson", "TenderContact", "ContactName", "Consultant"],
    "notes": ["Notes", "Description", "Comments", "Remarks"],
}

APPROVED_STATUSES = {"approved", "awarded", "award"}

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")


class DatabuildError(Exception):
    pass


def log(msg):
    print(f"[{dt.datetime.now().strftime('%H:%M:%S')}] {msg}")


def new_session():
    s = requests.Session()
    s.headers["User-Agent"] = UA
    s.headers["Accept"] = "application/json, text/plain, */*"
    return s


def find_token(text):
    m = re.search(r'name="__RequestVerificationToken"\s+type="hidden"\s+value="([^"]+)"', text)
    if not m:
        m = re.search(r'__RequestVerificationToken[^>]*value="([^"]+)"', text)
    return m.group(1) if m else None


def read_env(env_path=None):
    candidates = []
    if env_path:
        candidates.append(Path(env_path))
    candidates.append(Path(__file__).resolve().with_name(".env"))
    candidates.append(Path.cwd() / ".env")
    chosen = None
    for c in candidates:
        if c.exists():
            chosen = c
            break
    if chosen is None:
        return {}, None
    if load_dotenv is not None:
        load_dotenv(dotenv_path=chosen, override=False)
        env = {k: os.environ.get(k, "") for k in ("DATABUILD_USERNAME", "DATABUILD_PASSWORD", "DATABUILD_OUTPUT_DIR")}
        return env, chosen
    env = {}
    for line in chosen.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, _, v = line.partition("=")
        env[k.strip().upper()] = v.strip().strip('"').strip("'")
    return env, chosen


def output_dir_from_env(env):
    override = env.get("DATABUILD_OUTPUT_DIR")
    if override:
        return Path(override)
    onedrive = os.environ.get("ONEDRIVE")
    if onedrive:
        return Path(onedrive) / "Databuild"
    home_onedrive = Path.home() / "OneDrive" / "Databuild"
    if home_onedrive.parent.exists():
        return home_onedrive
    return Path.home() / "Databuild"


def login(session, username, password):
    log("Fetching login page for session token...")
    r = session.get(LOGIN_URL, timeout=30)
    r.raise_for_status()
    token = find_token(r.text)
    if not token:
        raise DatabuildError("Could not read __RequestVerificationToken from login page.")

    data = {
        "UserName": username,
        "Password": password,
        "RememberMe": "false",
        "ReturnUrl": DEFAULT_RETURN,
        "__RequestVerificationToken": token,
    }
    log("Submitting credentials...")
    r = session.post(LOGIN_URL, data=data, allow_redirects=False, timeout=30)

    if r.status_code in (301, 302) or (r.status_code == 200 and DEFAULT_RETURN in r.text):
        location = r.headers.get("Location", DEFAULT_RETURN)
        log(f"Login accepted (redirect: {location}).")
        session.get(BASE_URL + DEFAULT_RETURN, timeout=30)
        return
    if r.status_code == 200:
        lowered = r.text.lower()
        if "password" in lowered and ("incorrect" in lowered or "invalid" in lowered or "fail" in lowered):
            raise DatabuildError("Login rejected: check DATABUILD_USERNAME / DATABUILD_PASSWORD in .env.")
        if "login" in lowered or "forgotpassword" in lowered:
            raise DatabuildError("Login failed (login page returned). Verify credentials.")
    raise DatabuildError(f"Unexpected login response: HTTP {r.status_code}")


def probe(session, base, prefix, report):
    log("Probing candidate project endpoints...")
    out = {}
    for name in prefix:
        url = f"{base}/api/Project/{name}"
        try:
            r = session.get(url, timeout=60)
            sample = None
            if "json" in r.headers.get("Content-Type", ""):
                try:
                    j = r.json()
                    sample = summarize_json(j)
                except ValueError:
                    sample = r.text[:250]
            else:
                sample = r.text[:250]
            out[name] = {"url": url, "http": r.status_code, "sample": sample}
            log(f"{name}: HTTP {r.status_code}")
        except requests.RequestException as e:
            out[name] = {"url": url, "http": "error", "error": str(e)}
            log(f"{name}: error {e}")
    report["endpoints"] = out
    return out


def summarize_json(j, depth=0):
    if depth > 3:
        return type(j).__name__
    if isinstance(j, dict):
        keys = list(j.keys())[:12]
        return {k: summarize_json(j[k], depth + 1) for k in keys}
    if isinstance(j, list):
        counts = {}
        for item in j[:5]:
            t = summarize_json(item, depth + 1)
            if isinstance(t, dict):
                counts.setdefault("dict", []).append(t)
            else:
                counts.setdefault(type(j[0]).__name__, counts.get(type(j[0]).__name__, 0) + 1)
        return {"len": len(j), "sample": counts}
    return j


def extract_projects(payload):
    projects = None
    if isinstance(payload, dict):
        for key in ("projects", "data", "items", "results", "rows", "ProjectList", "List", "result"):
            v = payload.get(key)
            if isinstance(v, list):
                projects = v
                break
        if projects is None:
            for k, v in payload.items():
                if isinstance(v, list):
                    projects = v
                    break
        if projects is None and isinstance(payload.get("data"), dict):
            projects = extract_projects(payload["data"])
        if projects is None:
            projects = [payload]
    elif isinstance(payload, list):
        projects = payload
    return projects or []


def get_field(project, field):
    for key in FIELD_SYNONYMS[field]:
        if isinstance(project, dict) and key in project:
            return stringify(project[key])
        if isinstance(project, dict):
            for subkey, val in project.items():
                if isinstance(val, dict) and key in val:
                    return stringify(val[key])
    return ""


def stringify(v):
    if v is None:
        return ""
    if isinstance(v, (dict, list)):
        return json.dumps(v, ensure_ascii=False)
    return str(v)


def normalize_project(p):
    if not isinstance(p, dict):
        p = {"Item": p}
    return {
        "Ref": get_field(p, "ref"),
        "Project Name": get_field(p, "name"),
        "Client / Owner": get_field(p, "client"),
        "Location": get_field(p, "location"),
        "Sector": get_field(p, "sector"),
        "Status": get_field(p, "status"),
        "Date Approved": get_field(p, "approved_date"),
        "Stage": get_field(p, "stage"),
        "Estimated Value": get_field(p, "value"),
        "Contact": get_field(p, "contact"),
        "Notes": get_field(p, "notes"),
    }


def is_approved(project):
    status = get_field(project, "status").strip().lower()
    return status in APPROVED_STATUSES or "approved" in status or "award" in status


def fetch_projects(session, base, report):
    log("Trying dashboard data endpoints...")
    endpoints_report = report.setdefault("endpoints", {})
    combined = []
    used = "Dashboard"
    tried = []
    for name in PROJECT_ENDPOINTS:
        url = f"{base}/api/Project/{name}"
        tried.append(name)
        try:
            r = session.get(url, timeout=60)
        except requests.RequestException:
            continue
        if r.status_code != 200 or "json" not in r.headers.get("Content-Type", ""):
            endpoints_report.setdefault(name, {}).setdefault("poll", None)
            continue
        try:
            payload = r.json()
        except ValueError:
            continue
        projects = extract_projects(payload)
        if projects:
            combined = projects
            used = name
            endpoints_report[name] = {
                "url": url,
                "http": r.status_code,
                "project_count": len(combined),
            }
            break
    if not combined:
        raise DatabuildError(
            "No JSON project list found at /api/Project/*. Run with --probe to inspect "
            "available endpoints, or capture a browser network (HAR) log of the "
            "dashboard so the endpoint map can be updated."
        )
    log(f"Using endpoint /api/Project/{used} ({len(combined)} projects).")
    report["endpoint"] = f"/api/Project/{used}"
    report["endpoints_tried"] = tried
    return combined


def write_csv(path, rows):
    if not rows:
        return
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)


def write_xlsx(path, rows):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Approved Projects"
    headers = list(rows[0].keys()) if rows else ["Ref", "Project Name", "Client / Owner"]
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="0E7C66")
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
    for row in rows:
        ws.append([row.get(h, "") for h in headers])
    for idx, h in enumerate(headers, start=1):
        width = max(len(h), max((len(str(row.get(h, ""))) for row in rows), default=0))
        ws.column_dimensions[get_column_letter(idx)].width = min(max(width + 2, 10), 45)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def main():
    parser = argparse.ArgumentParser(description="Databuild project puller")
    parser.add_argument("--probe", action="store_true", help="discover endpoints and stop")
    parser.add_argument("--all", action="store_true", help="include non-approved statuses too")
    parser.add_argument("--output-dir", help="override output folder")
    parser.add_argument("--dotenv", help="path to a specific .env file")
    parser.add_argument("--statuses", help="comma-separated approved status filters (default: approved,awarded)")
    args = parser.parse_args()

    env, env_file = read_env(args.dotenv)
    username = env.get("DATABUILD_USERNAME", "")
    password = env.get("DATABUILD_PASSWORD", "")

    if env_file is None:
        parser.error("No .env file found. Copy .env.example to .env and fill in credentials.")

    if not username or not password:
        parser.error("DATABUILD_USERNAME / DATABUILD_PASSWORD missing in .env.")

    out_dir = Path(args.output_dir) if args.output_dir else output_dir_from_env(env)
    out_dir = out_dir.expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    global APPROVED_STATUSES
    if args.statuses:
        APPROVED_STATUSES = {s.strip().lower() for s in args.statuses.split(",") if s.strip()}

    session = new_session()
    login(session, username, password)

    stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    report = {"run_at": stamp, "user": username, "output_dir": str(out_dir)}

    if args.probe:
        probe(session, BASE_URL, PROJECT_ENDPOINTS, report)
        out = out_dir / "probe_report.json"
        out.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
        log(f"Probe report written: {out}")
        sys.exit(0)

    projects = fetch_projects(session, BASE_URL, report)

    if args.all:
        kept = projects
    else:
        kept = [p for p in projects if is_approved(p)]
        report["filtered"] = {"approved": len(kept), "total_pulled": len(projects)}
        log(f"Filtered to {len(kept)} approved project(s) of {len(projects)} pulled.")

    rows = [normalize_project(p) for p in kept]
    rows.sort(key=lambda r: (r.get("Date Approved", "") or "", r.get("Project Name", "") or ""), reverse=False)

    base_name = "DataBuild_Approved_Projects"
    paths = {}
    xlsx_path = out_dir / f"{base_name}.xlsx"
    csv_path = out_dir / f"{base_name}.csv"
    json_path = out_dir / f"{base_name}.json"

    write_xlsx(xlsx_path, rows)
    write_csv(csv_path, rows)
    json_path.write_text(json.dumps(rows, indent=2, ensure_ascii=False), encoding="utf-8")

    report["files"] = {
        "xlsx": str(xlsx_path),
        "csv": str(csv_path),
        "json": str(json_path),
        "project_count": len(rows),
    }
    (out_dir / "last_run.json").write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")

    log(f"Wrote {len(rows)} project(s):")
    log(f"  Excel : {xlsx_path}")
    log(f"  CSV   : {csv_path}")
    log(f"  JSON  : {json_path}")
    log("Done.")


if __name__ == "__main__":
    try:
        main()
    except DatabuildError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    except KeyboardInterrupt:
        print("Interrupted.", file=sys.stderr)
        sys.exit(130)
    except requests.RequestException as e:
        print(f"Network error: {e}", file=sys.stderr)
        sys.exit(1)
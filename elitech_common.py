# ============================================================
# elitech_common.py - shared config and helpers for
# elitech_live.py (every 5 min) and elitech_summary.py (daily)
# ============================================================

import os, sys, re, json, time, smtplib, html
from datetime import datetime, timedelta
from email.message import EmailMessage

import requests
import pandas as pd
import pytz
import gspread

# ── CONFIG ────────────────────────────────────────────────
env = lambda k, d="": (os.environ.get(k) or d).strip()
KEY_ID     = env("ELITECH_KEY_ID")
KEY_SECRET = env("ELITECH_KEY_SECRET")
USERNAME   = env("ELITECH_USERNAME")
PASSWORD   = env("ELITECH_PASSWORD")
GMAIL_PASS = env("GMAIL_APP_PASSWORD")
TEST_MODE  = env("TEST_MODE", "false").lower() == "true"
SEND_EMAIL = env("SEND_EMAIL", "true").lower() == "true"
SUMMARY_DAY_OFFSET = int(env("SUMMARY_DAY_OFFSET", "1"))
IN_COLAB = False

GOOGLE_SHEET_ID = "19hPHSr6dmFMshW2S4G6T16ZZkcL3B0SdkvkSDpQ11W4"
TAB_PREFIX = "TEST " if TEST_MODE else ""
STATE_TAB    = TAB_PREFIX + "Elitech State"
READINGS_TAB = TAB_PREFIX + "Elitech Readings"
ALERTS_TAB   = TAB_PREFIX + "Elitech Alerts"

TZ = pytz.timezone("Asia/Riyadh")

SENDER_EMAIL  = "aof.group.auto@gmail.com"
SENDER_NAME   = "Business Intelligence"
# ---- Instant alerts (temperature only), sent by elitech_live.py ----
ALERT_TO = [
    "a.omara@aofgroup.com",
    "m.alhuaydar@aofgroup.com",
    "h.modwi@aofgroup.com",
    "cctvdep@aofgroup.com",
]
ALERT_CC = [
    "m.abouelmaaty@albawasiq.com",   # Maher
    "a.alsalem@aofgroup.com",
    "o.salahaddin@aofgroup.com",
    "omar@aofgroup.com",
    "faisal@aofgroup.com",
]

# ---- Morning reports for yesterday (temperature summary + offline / no-reading report), sent by elitech_summary.py ----
SUMMARY_TO = [
    "a.omara@aofgroup.com",
    "m.houider@aofgroup.com",
    "h.modwi@aofgroup.com",
    "cctvdep@aofgroup.com",
]
SUMMARY_CC = [
    "m.abouelmaaty@albawasiq.com",   # Maher
    "a.alsalem@aofgroup.com",
    "o.salahaddin@aofgroup.com",
    "omar@aofgroup.com",
    "faisal@aofgroup.com",
]

# Which list send_email() uses; each script sets this at start.
RECIPIENTS_TO, RECIPIENTS_CC = ALERT_TO, ALERT_CC

def use_recipients(kind):
    """kind = 'alert' or 'summary'."""
    global RECIPIENTS_TO, RECIPIENTS_CC
    RECIPIENTS_TO, RECIPIENTS_CC = (SUMMARY_TO, SUMMARY_CC) if kind == "summary" else (ALERT_TO, ALERT_CC)

TEST_RECIPIENTS = ["a.alsalem@aofgroup.com"]

POLL_MIN          = 5     # how often elitech_live.py is triggered
CONFIRM_MIN       = int(env("CONFIRM_MIN", "15"))  # out of range this long (4 readings in a row) before the first alert
REMIND_HOURS      = 4     # re-alert if still broken after this long
RESOLVE_MARGIN    = 0.5   # must be this far inside the range to count as back to normal
RESOLVE_READINGS  = 3     # "resolved" only after this many good readings in a row (15 min)
OFFLINE_AFTER_MIN = 60    # no upload for this long = offline
INACTIVE_DAYS     = 7     # devices silent longer than this are not tracked (listed in summary only)
KEEP_DAYS         = 7     # readings older than this are removed from the sheet

# Ranges by type. Chiller 0..5. Freezer: alert when warmer than -15 (colder is never an alarm).
TEMP_RANGES = {
    "chiller": (0.0, 5.0),
    "freezer": (None, -15.0),
}

# Device -> type. A channel whose name contains "Freezer"/"Chiller" (HAJED, MAJED) is typed automatically.
# Values below are GUESSED from the 4 Nov readings - confirm them.
# Use a dict for multi-probe devices, "ignore" to skip a device entirely.
DEVICE_TYPE = {
    # Branch freezers (reading -10 .. -25)
    "ANRUH": "freezer", "AZRUH": "freezer", "BDRUH": "freezer", "DARUH": "freezer",
    "FYJED": "freezer", "GaratissB01": "freezer", "GaratissB02": "freezer", "HSRUH": "freezer",
    "IRRUH": "freezer", "KHRUH": "freezer", "KRRUH": "freezer", "LBRUH": "freezer",
    "NDRUH": "freezer", "NSRUH": "freezer", "PSJED": "freezer", "Plaster Freez1": "freezer",
    "Qadisiyah": "freezer", "RBRUH": "freezer", "RWAHS": "freezer", "RWRUH": "freezer",
    "SWRUH": "freezer", "Sifa": "freezer", "TKRUH": "freezer", "URURUH": "freezer",
    "SARUH": "freezer", "SLAHS": "freezer", "TWRUH": "freezer",
    "SFJED": "freezer",
    "HIJED": {"Probe 1": "freezer", "Probe 3": "chiller"},
    # Chillers (reading around 0 .. 9)
    "QRRUH": "freezer",   # read -17.2 on 6 Nov
    "Lubda B01": "chiller", "Lubda B02": "chiller", "Sauces Storage": "chiller",
    # Unknown - fill in
    # "NRRUH": "",          # read 16.8 C on 4 Nov
    # "Defrosting Egg": "",
    "Big Truck Plate (5406)": "ignore",
}

# Branch labels from the Zenput temperature report; matched on the first word (device name).
BRANCH_LABELS = [
    "LBRUH B07", "FAYJED B32", "ANRUH B31", "SLAHS B23", "NDRUH B15", "SWRUH B08", "SARUH B27",
    "DARUH B06", "KRRUH B21", "SFJED B24", "RBRUH B14", "RAWRUH B05", "NSRUH B04", "SHRUH B10",
    "KHRUH B02", "NURUH B01", "TWRUH B12", "AZRUH B09", "RWAHS B25", "HIRJED B33", "NRRUH B11",
    "MURUH B19", "QADRUH B30", "PSJED B36", "OBJED B22", "URURUH B34", "HAJED B26", "MAJED B28",
    "IRRUH B35", "BDRUH B16", "AQRUH B13", "TKRUH B18", "QRRUH B17", "UHDMM B38", "HSRUH B39",
]
BRANCH_BY_CODE = {b.split()[0]: b for b in BRANCH_LABELS}
BRANCH_ALIASES = {"FYJED": "FAYJED B32", "HIJED": "HIRJED B33"}  # device name differs from Zenput code

BASE = "https://www.i-elitech.com"
TOKEN_URL = f"{BASE}/api/data-api/elitechAccess/getToken"
LIST_URL  = f"{BASE}/api/device/Device/get/deviceList"
FMT = "%Y-%m-%d %H:%M:%S"

NOW = datetime.now(TZ).replace(tzinfo=None, microsecond=0)
NOW_S = NOW.strftime(FMT)


# ── HELPERS ───────────────────────────────────────────────
def branch_of(device):
    return BRANCH_ALIASES.get(device) or BRANCH_BY_CODE.get(device) or device

def channel_type(device, channel):
    c = str(channel).lower()
    if "freez" in c:
        return "freezer"
    if "chill" in c:
        return "chiller"
    t = DEVICE_TYPE.get(device, "")
    if isinstance(t, dict):
        return t.get(channel, "")
    return t

def out_of_range(v, typ, margin=0.0):
    """margin > 0 narrows the range: used to decide an alert is really resolved."""
    lo, hi = TEMP_RANGES.get(typ, (None, None))
    return (lo is not None and v < lo + margin) or (hi is not None and v > hi - margin)

def worst_reading(values, typ):
    """The reading furthest outside the allowed range."""
    lo, hi = TEMP_RANGES.get(typ, (None, None))
    def dist(v):
        return max((v - hi) if hi is not None else 0, (lo - v) if lo is not None else 0)
    return max(values, key=dist)

def branch_sort_key(branch, device="", channel=""):
    """Order: Classic B01..Bnn, then Garatis QB.., then Lubda LB.., then everything else A-Z.
    Devices/probes of the same branch stay together."""
    b = str(branch or "")
    low = b.lower()
    m = re.search(r"(\d+)\s*$", b) or re.search(r"[A-Za-z]*\s*B\s*0*(\d+)", b)
    num = int(m.group(1)) if m else 9999
    if "garat" in low or re.search(r"\bQB\s*\d", b):
        group = 1
    elif "lubda" in low or re.search(r"\bLB\s*\d", b):
        group = 2
    elif re.search(r"\bB\s*\d+\s*$", b):
        group = 0
    else:
        group, num = 3, 0
    return (group, num, low, str(device or "").lower(), str(channel or "").lower())

def range_text(typ):
    lo, hi = TEMP_RANGES.get(typ, (None, None))
    if lo is None and hi is None:
        return "-"
    if lo is None:
        return f"<= {hi:g}"
    if hi is None:
        return f">= {lo:g}"
    return f"{lo:g} .. {hi:g}"

def fnum(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None

def utc_to_ksa(s):
    try:
        return datetime.strptime(str(s).replace("T", " ")[:19], FMT) + timedelta(hours=3) if s else None
    except ValueError:
        return None

def parse_ts(s):
    try:
        return datetime.strptime(str(s)[:19], FMT)
    except ValueError:
        return None

def fmt_minutes(m):
    m = int(round(m or 0))
    h, mm = divmod(m, 60)
    return f"{h}h {mm:02d}m" if h else f"{mm}m"

TYPE_AR  = {"freezer": "فريزر", "chiller": "ثلاجة", "": "-"}
KIND_AR  = {"Temperature": "درجة حرارة خارج النطاق", "Offline": "الجهاز غير متصل",
            "No reading": "لا توجد قراءة حرارة (حساس)"}
EVENT_AR = {"NEW": "تنبيه جديد", "REMINDER": "ما زال قائما", "RESOLVED": "تم الحل"}


# ── ELITECH ───────────────────────────────────────────────
def get_token():
    body = {"keyId": KEY_ID, "keySecret": KEY_SECRET, "userName": USERNAME, "password": PASSWORD}
    for attempt in range(3):
        j = requests.post(TOKEN_URL, json=body, timeout=30).json()
        code = str(j.get("code"))
        print(f"getToken attempt {attempt + 1}: code={code} msg={j.get('msg') or j.get('message')}")
        if code == "0" and j.get("data"):
            return j["data"]
        if code != "5110":
            break
        time.sleep(65)
    sys.exit("No token from Elitech (5108 = API service expired, 5100 = login wrong).")

def fetch_devices(token):
    devices, seen, page = [], set(), 1
    while True:
        r = requests.post(LIST_URL, json={"start": page, "pageSize": 50, "typeScenes": 1},
                          headers={"Authorization": token, "Content-Type": "application/json"}, timeout=60)
        j = r.json()
        if str(j.get("code")) != "0":
            sys.exit(f"deviceList failed: HTTP {r.status_code} code={j.get('code')} msg={j.get('message')}")
        data = j.get("data") or {}
        new = [d for d in (data.get("list") or []) if d.get("id") not in seen]
        seen.update(d.get("id") for d in new)
        devices.extend(new)
        if not new or data.get("isLastPage") or len(devices) >= (data.get("total") or 0):
            print(f"Devices fetched: {len(devices)}")
            return devices
        page += 1
        time.sleep(1)

def snapshot(devices):
    """One record per device: status, temperature channels, last upload."""
    out = []
    for d in devices:
        name = d.get("deviceName")
        last = utc_to_ksa(d.get("lastSessionTime"))
        mins = (NOW - last).total_seconds() / 60 if last else None
        channels, faults = [], []
        for probe in d.get("data") or []:
            for p in probe.get("dataDetailList") or []:
                if p.get("unitCode") != "℃":
                    continue
                ch = p.get("paramName") or probe.get("probeName")
                if p.get("value") == "FAULT":
                    faults.append(ch)
                    continue
                v = fnum(p.get("value"))
                if v is not None:
                    channels.append({"channel": ch, "temp": v, "type": channel_type(name, ch),
                                     "time": utc_to_ksa(p.get("monitorTime"))})
        if mins is None or mins > INACTIVE_DAYS * 1440:
            status = "Inactive"
        elif not d.get("onlineFlag") or mins > OFFLINE_AFTER_MIN:
            status = "Offline"
        elif not channels:
            status = "No reading"
        else:
            status = "Online"
        out.append({"device": name, "branch": branch_of(name), "status": status, "last": last,
                    "minutes": mins, "channels": channels, "faults": faults,
                    "battery": d.get("devicePower"), "ignore": DEVICE_TYPE.get(name) == "ignore"})
    return out


# ── GOOGLE SHEET ──────────────────────────────────────────
def google_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    raw = os.environ.get("GOOGLE_CREDENTIALS")
    if raw:
        from google.oauth2.service_account import Credentials
        return gspread.authorize(Credentials.from_service_account_info(json.loads(raw), scopes=scopes))
    if IN_COLAB:
        from google.colab import auth
        from google.auth import default
        auth.authenticate_user()
        creds, _ = default()
        return gspread.authorize(creds)
    sys.exit("No Google credentials (GOOGLE_CREDENTIALS secret missing).")

def get_tab(sh, title, header):
    try:
        return sh.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=1000, cols=max(len(header), 2))
        if header:
            ws.append_row(header, value_input_option="RAW")
        print(f"Created tab: {title}")
        return ws

READ_HEADER  = ["Poll time (KSA)", "Device", "Branch", "Channel", "Type", "Temp", "Status", "Reading time (KSA)"]
ALERT_HEADER = ["Time (KSA)", "Event", "Kind", "Device", "Branch", "Channel", "Detail"]

def load_state(ws):
    raw = "".join(ws.col_values(1))
    try:
        return json.loads(raw) if raw else {}
    except json.JSONDecodeError:
        print("State unreadable, starting fresh.")
        return {}

def save_state(ws, state):
    s = json.dumps(state, ensure_ascii=False, separators=(",", ":"))
    chunks = [s[i:i + 40000] for i in range(0, len(s), 40000)] or ["{}"]
    ws.clear()
    ws.update(range_name="A1", values=[[c] for c in chunks], value_input_option="RAW")


# ── EMAIL ─────────────────────────────────────────────────
def table(headers, rows):
    if not rows:
        return "<p style='color:#2e7d32;margin:4px 0 14px;'>لا يوجد</p>"
    th = "".join(f"<th style='padding:8px 10px;border:1px solid #ddd;text-align:right;background:#2C3E50;color:#fff;'>{h}</th>" for h in headers)
    trs = ""
    for i, r in enumerate(rows):
        bg = "#FFEEEE" if r and str(r[0]).startswith("!") else ("#F7F7F7" if i % 2 else "#FFFFFF")
        tds = "".join(f"<td style='padding:7px 10px;border:1px solid #ddd;'>{html.escape(str(c).lstrip('!'))}</td>" for c in r)
        trs += f"<tr style='background:{bg};'>{tds}</tr>"
    return f"<table style='border-collapse:collapse;width:100%;font-size:13px;margin-bottom:18px;'><tr>{th}</tr>{trs}</table>"

def wrap(title, subtitle, inner):
    return f"""<html><head><meta charset="UTF-8"></head>
    <body dir="rtl" style="font-family:Arial,sans-serif;font-size:14px;line-height:1.7;color:#000;max-width:950px;margin:auto;padding:20px;">
      <div style="background:#2C3E50;color:#fff;padding:16px 20px;border-radius:8px;margin-bottom:18px;">
        <h2 style="margin:0;">{title}</h2>
        <p style="margin:4px 0 0;font-size:13px;opacity:0.85;">{subtitle}</p>
      </div>
      {inner}
      <p style="color:#888;font-size:12px;">Business Intelligence - AOF Group</p>
    </body></html>"""

def send_email(subject, body, attachment=None):
    if TEST_MODE:
        subject = "[TEST] " + subject
    print("Subject:", subject)
    with open("email_preview.html", "w", encoding="utf-8") as f:
        f.write(body)
    if not SEND_EMAIL:
        print("SEND_EMAIL=false, not sent (preview saved to email_preview.html).")
        return
    if not GMAIL_PASS:
        sys.exit("GMAIL_APP_PASSWORD missing.")
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = f"{SENDER_NAME} <{SENDER_EMAIL}>"
    msg["To"] = ", ".join(TEST_RECIPIENTS if TEST_MODE else RECIPIENTS_TO)
    if not TEST_MODE and RECIPIENTS_CC:
        msg["Cc"] = ", ".join(RECIPIENTS_CC)
    msg.set_content("This report needs an HTML-capable email client.")
    msg.add_alternative(body, subtype="html")
    files = attachment if isinstance(attachment, (list, tuple)) else [attachment]
    for path in files:
        if not path or not os.path.exists(path):
            continue
        if path.endswith(".pdf"):
            mt, st = "application", "pdf"
        else:
            mt, st = "application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        with open(path, "rb") as f:
            msg.add_attachment(f.read(), maintype=mt, subtype=st, filename=os.path.basename(path))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(SENDER_EMAIL, GMAIL_PASS)
        s.send_message(msg)
    print("Email sent.")



def open_store():
    print(f"Run {NOW_S} KSA | test={TEST_MODE} | confirm={CONFIRM_MIN}m")
    missing = [n for n, v in {"ELITECH_KEY_ID": KEY_ID, "ELITECH_KEY_SECRET": KEY_SECRET,
                              "ELITECH_USERNAME": USERNAME, "ELITECH_PASSWORD": PASSWORD}.items() if not v]
    if missing:
        sys.exit(f"Missing secrets: {missing}")
    sh = google_client().open_by_key(GOOGLE_SHEET_ID)
    ws_state = get_tab(sh, STATE_TAB, [])
    return sh, ws_state, load_state(ws_state)

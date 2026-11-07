# ============================================================
# elitech_live.py - runs every 5 minutes.
# Reads every logger, logs the readings to the Google Sheet and sends an
# instant email only when something new goes wrong, stays wrong for
# REMIND_HOURS, or recovers.
# ============================================================
from elitech_common import *

# ── LIVE ──────────────────────────────────────────────────
def run_live(snap, sh, state):
    issues = state.setdefault("issues", {})
    baseline = not state.get("initialized")
    events, touched = [], set()

    def track(key, active, info, confirm_min, clear=True):
        touched.add(key)
        rec = issues.get(key)
        if active:
            if rec is None:
                rec = issues[key] = {**info, "since": NOW_S, "alerted": False, "last_alert": "",
                                     "silent": baseline and info["kind"] != "Temperature", "peak": info.get("value")}
            rec["ok_count"] = 0
            rec["detail"], rec["value"] = info["detail"], info.get("value")
            rec["last_bad"] = NOW_S
            if info.get("value") is not None:
                prev = rec.get("peak")
                rec["peak"] = info["value"] if prev is None else worst_reading([prev, info["value"]], info.get("type", ""))
            if rec["silent"]:
                return
            age = (NOW - parse_ts(rec["since"])).total_seconds() / 60
            if not rec["alerted"] and age >= confirm_min:
                rec["alerted"], rec["last_alert"] = True, NOW_S
                events.append(("NEW", dict(rec)))
            elif rec["alerted"] and (NOW - parse_ts(rec["last_alert"])).total_seconds() / 3600 >= REMIND_HOURS:
                rec["last_alert"] = NOW_S
                events.append(("REMINDER", dict(rec)))
        elif rec is not None:
            if not rec["alerted"] or rec["silent"]:
                issues.pop(key, None)          # never alerted: just forget it
                return
            if not clear:                      # in range but still near the limit: keep it open
                return
            rec["ok_count"] = rec.get("ok_count", 0) + 1
            if rec["ok_count"] >= RESOLVE_READINGS:  # back in range several readings in a row
                events.append(("RESOLVED", {**rec, "detail": info.get("ok_detail", rec["detail"]), "now": info.get("value")}))
                issues.pop(key, None)

    log_rows = []
    for d in snap:
        if d["ignore"] or d["status"] == "Inactive":
            continue
        dev, br = d["device"], d["branch"]
        base = {"device": dev, "branch": br, "channel": "",
                "battery": f"{d['battery']}%" if d.get("battery") is not None else "-"}

        track(f"{dev}|offline", d["status"] == "Offline",
              {**base, "kind": "Offline", "detail": f"آخر اتصال {d['last']:%Y-%m-%d %H:%M}" if d["last"] else "",
               "ok_detail": "عاد الاتصال"}, 0)
        track(f"{dev}|noread", d["status"] == "No reading",
              {**base, "kind": "No reading",
               "detail": ("FAULT: " + ", ".join(d["faults"])) if d["faults"] else "لا توجد قيمة",
               "ok_detail": "عادت القراءة"}, CONFIRM_MIN)

        if d["status"] in ("Offline", "No reading"):
            log_rows.append([NOW_S, dev, br, "", "", "", d["status"],
                             d["last"].strftime(FMT) if d["last"] else ""])
            continue

        for c in d["channels"]:
            bad = bool(c["type"]) and out_of_range(c["temp"], c["type"])
            clear = not (c["type"] and out_of_range(c["temp"], c["type"], RESOLVE_MARGIN))
            track(f"{dev}|{c['channel']}", bad,
                  {**base, "channel": c["channel"], "kind": "Temperature", "type": c["type"], "value": c["temp"],
                   "detail": f"{c['temp']:g} C (المطلوب {range_text(c['type'])})",
                   "ok_detail": f"{c['temp']:g} C ضمن النطاق"}, CONFIRM_MIN, clear)
            status = "OUT" if bad else ("OK" if c["type"] else "Untyped")
            log_rows.append([NOW_S, dev, br, c["channel"], c["type"], c["temp"], status,
                             c["time"].strftime(FMT) if c["time"] else ""])

    for key in [k for k in issues if k not in touched]:
        issues.pop(key, None)  # device went inactive/ignored: drop quietly

    state["initialized"] = True
    state["last_live"] = NOW_S

    ws_read = get_tab(sh, READINGS_TAB, READ_HEADER)
    if log_rows:
        ws_read.append_rows(log_rows, value_input_option="RAW")
    print(f"Logged {len(log_rows)} readings.")

    if events:
        ws_alert = get_tab(sh, ALERTS_TAB, ALERT_HEADER)
        ws_alert.append_rows([[NOW_S, e, r["kind"], r["device"], r["branch"], r.get("channel", ""), r["detail"]]
                              for e, r in events], value_input_option="RAW")

    counts = pd.Series([s["status"] for s in snap]).value_counts().to_dict()
    print(f"Status: {counts} | open issues: {sum(1 for r in issues.values() if r['alerted'])} | "
          f"events this run: {len(events)}" + (" (baseline run, existing offline/no-reading not alerted)" if baseline else ""))

    if events:
        temp_events = [e for e in events if e[1]["kind"] == "Temperature"]
        dev_events  = [e for e in events if e[1]["kind"] != "Temperature"]
        if temp_events:
            send_instant_email(temp_events, issues, "temp")
        if dev_events:
            send_instant_email(dev_events, issues, "device")
    else:
        print("Nothing new - no instant email.")


RESOLVED_AR = {"Temperature": "كانت درجة الحرارة خارج النطاق", "Offline": "كان الجهاز غير متصل",
               "No reading": "كانت بدون قراءة حرارة"}

def send_instant_email(events, issues, report):
    def c(v):
        return "-" if v is None else f"{v:g} C"

    def open_rows(kind_events):
        out = []
        for e, r in kind_events:
            mins = (NOW - parse_ts(r["since"])).total_seconds() / 60
            temp = r["kind"] == "Temperature"
            out.append(["!" + r["branch"], r.get("channel") or "-", TYPE_AR.get(r.get("type", ""), "-"),
                        KIND_AR[r["kind"]],
                        c(r.get("value")) if temp else r["detail"],
                        c(r.get("peak")) if temp else "-",
                        range_text(r.get("type", "")) if temp else "-",
                        r["since"][11:16], fmt_minutes(mins)])
        return out

    def resolved_rows(kind_events):
        out = []
        for e, r in kind_events:
            start = parse_ts(r["since"])
            last_bad = parse_ts(r.get("last_bad") or "") or NOW
            mins = (last_bad - start).total_seconds() / 60 + POLL_MIN
            temp = r["kind"] == "Temperature"
            out.append([r["branch"], r.get("channel") or "-", TYPE_AR.get(r.get("type", ""), "-"),
                        RESOLVED_AR[r["kind"]],
                        c(r.get("peak")) if temp else "-",
                        c(r.get("now")) if temp else "عاد",
                        range_text(r.get("type", "")) if temp else "-",
                        r["since"][11:16], last_bad.strftime("%H:%M"), fmt_minutes(mins)])
        return out

    new = [x for x in events if x[0] == "NEW"]
    rem = [x for x in events if x[0] == "REMINDER"]
    res = [x for x in events if x[0] == "RESOLVED"]
    is_temp = report == "temp"
    open_n = sum(1 for r in issues.values() if r["alerted"] and not r["silent"]
                 and (r["kind"] == "Temperature") == is_temp)
    if not is_temp:
        hdr_open = ["الفرع / الجهاز", "المشكلة", "التفاصيل", "البطارية", "بدأت", "المدة"]
        hdr_res  = ["الفرع / الجهاز", "المشكلة", "بدأت", "آخر مرة", "مدة المشكلة"]
    hdr_open = ["الفرع / الجهاز", "الحساس", "النوع", "المشكلة", "القراءة الحالية", "أسوأ قراءة", "المطلوب", "بدأت", "المدة"]
    hdr_res  = ["الفرع / الجهاز", "الحساس", "النوع", "المشكلة", "أسوأ قراءة", "القراءة الآن", "المطلوب",
                "بدأت", "آخر قراءة سيئة", "مدة المشكلة"]

    if not is_temp:
        def open_rows(kind_events):
            return [["!" + r["branch"], KIND_AR[r["kind"]], r["detail"], r.get("battery", "-"),
                     r["since"][11:16], fmt_minutes((NOW - parse_ts(r["since"])).total_seconds() / 60)]
                    for e, r in kind_events]

        def resolved_rows(kind_events):
            out = []
            for e, r in kind_events:
                last_bad = parse_ts(r.get("last_bad") or "") or NOW
                out.append([r["branch"], RESOLVED_AR[r["kind"]], r["since"][11:16], last_bad.strftime("%H:%M"),
                            fmt_minutes((last_bad - parse_ts(r["since"])).total_seconds() / 60 + POLL_MIN)])
            return out

    inner = ""
    if new:
        inner += f"<h3 style='color:#C62828;margin:6px 0;'>تنبيهات جديدة ({len(new)})</h3>" + table(hdr_open, open_rows(new))
    if rem:
        inner += f"<h3 style='color:#EF6C00;margin:6px 0;'>ما زالت قائمة منذ أكثر من {REMIND_HOURS} ساعات ({len(rem)})</h3>" + table(hdr_open, open_rows(rem))
    if res:
        inner += f"<h3 style='color:#2E7D32;margin:6px 0;'>تم الحل ({len(res)})</h3>" + table(hdr_res, resolved_rows(res))

    parts = []
    if new: parts.append(f"{len(new)} جديد")
    if rem: parts.append(f"{len(rem)} مستمر")
    if res: parts.append(f"{len(res)} تم الحل")
    sub = f"<span dir='ltr'>{NOW:%Y-%m-%d %H:%M}</span> KSA | التنبيهات المفتوحة حاليا: {open_n}"
    if is_temp:
        subject = f"تنبيه فوري - درجات حرارة الثلاجات | {' - '.join(parts)} | {NOW:%H:%M}"
        send_email(subject, wrap("تنبيه فوري - درجات الحرارة Elitech", sub, inner))
    else:
        subject = f"تنبيه أجهزة المراقبة - انقطاع أو بدون قراءة | {' - '.join(parts)} | {NOW:%H:%M}"
        send_email(subject, wrap("تنبيه أجهزة المراقبة - انقطاع الاتصال أو القراءة Elitech", sub, inner))



if __name__ == "__main__":
    sh, ws_state, state = open_store()
    snap = snapshot(fetch_devices(get_token()))
    run_live(snap, sh, state)
    save_state(ws_state, state)
    print("Done.")

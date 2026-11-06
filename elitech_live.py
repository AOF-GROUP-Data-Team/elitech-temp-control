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
            if info.get("value") is not None and rec.get("peak") is not None:
                lo, hi = TEMP_RANGES.get(info.get("type", ""), (None, None))
                rec["peak"] = max(rec["peak"], info["value"]) if hi is not None and info["value"] > hi else min(rec["peak"], info["value"])
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
                events.append(("RESOLVED", {**rec, "detail": info.get("ok_detail", rec["detail"])}))
                issues.pop(key, None)

    log_rows = []
    for d in snap:
        if d["ignore"] or d["status"] == "Inactive":
            continue
        dev, br = d["device"], d["branch"]
        base = {"device": dev, "branch": br, "channel": ""}

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
        send_instant_email(events, issues)
    else:
        print("Nothing new - no instant email.")


def send_instant_email(events, issues):
    def rows_for(kind_events, mark):
        out = []
        for e, r in kind_events:
            mins = (NOW - parse_ts(r["since"])).total_seconds() / 60
            out.append([("!" if mark else "") + r["branch"], r.get("channel") or "-", TYPE_AR.get(r.get("type", ""), "-"),
                        KIND_AR[r["kind"]], r["detail"], r["since"][:16], fmt_minutes(mins)])
        return out

    new = [x for x in events if x[0] == "NEW"]
    rem = [x for x in events if x[0] == "REMINDER"]
    res = [x for x in events if x[0] == "RESOLVED"]
    open_n = sum(1 for r in issues.values() if r["alerted"] and not r["silent"])
    hdr = ["الفرع / الجهاز", "الحساس", "النوع", "المشكلة", "التفاصيل", "منذ", "المدة"]

    inner = ""
    if new:
        inner += f"<h3 style='color:#C62828;margin:6px 0;'>تنبيهات جديدة ({len(new)})</h3>" + table(hdr, rows_for(new, True))
    if rem:
        inner += f"<h3 style='color:#EF6C00;margin:6px 0;'>ما زالت قائمة منذ أكثر من {REMIND_HOURS} ساعات ({len(rem)})</h3>" + table(hdr, rows_for(rem, True))
    if res:
        inner += f"<h3 style='color:#2E7D32;margin:6px 0;'>تم الحل ({len(res)})</h3>" + table(hdr, rows_for(res, False))

    parts = []
    if new: parts.append(f"{len(new)} جديد")
    if rem: parts.append(f"{len(rem)} مستمر")
    if res: parts.append(f"{len(res)} تم الحل")
    subject = f"تنبيه فوري - درجات حرارة الثلاجات | {' - '.join(parts)} | {NOW:%H:%M}"
    sub = f"<span dir='ltr'>{NOW:%Y-%m-%d %H:%M}</span> KSA | التنبيهات المفتوحة حاليا: {open_n}"
    send_email(subject, wrap("تنبيه فوري - أجهزة مراقبة الحرارة Elitech", sub, inner))



if __name__ == "__main__":
    sh, ws_state, state = open_store()
    snap = snapshot(fetch_devices(get_token()))
    run_live(snap, sh, state)
    save_state(ws_state, state)
    print("Done.")

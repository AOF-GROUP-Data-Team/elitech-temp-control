# ============================================================
# elitech_summary.py - runs once a day (08:00 KSA).
# Full-day summary of yesterday 00:00-23:59 from every reading the live
# runs logged (one per probe every 5 minutes), emailed with an Excel file.
# ============================================================
from elitech_common import *
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── SUMMARY ───────────────────────────────────────────────
def run_summary(snap, sh, state):
    day = (NOW - timedelta(days=SUMMARY_DAY_OFFSET)).date()
    day_s = day.strftime("%Y-%m-%d")
    print(f"Building summary for {day_s} ...")

    ws_read = get_tab(sh, READINGS_TAB, READ_HEADER)
    values = ws_read.get_all_values()
    df = pd.DataFrame(values[1:], columns=values[0]) if len(values) > 1 else pd.DataFrame(columns=READ_HEADER)
    df = df[df["Poll time (KSA)"].str.startswith(day_s)].copy()
    if df.empty:
        print(f"No readings logged for {day_s}. The live tracker must run during that day first.")
        return False
    df["t"] = pd.to_datetime(df["Poll time (KSA)"])
    df["Temp"] = pd.to_numeric(df["Temp"], errors="coerce")
    df = df.sort_values("t")

    def add_durations(g):
        nxt = g["t"].shift(-1)
        g = g.copy()
        g["dur"] = ((nxt - g["t"]).dt.total_seconds() / 60).fillna(POLL_MIN).clip(upper=2 * POLL_MIN)
        return g

    temps = df[df["Temp"].notna()]
    temps = pd.concat([add_durations(g) for _, g in temps.groupby(["Device", "Channel"])]) if not temps.empty else temps
    devstat = df.drop_duplicates(["Poll time (KSA)", "Device"])
    devstat = pd.concat([add_durations(g) for _, g in devstat.groupby("Device")])

    polls = df["Poll time (KSA)"].nunique()

    ch_rows = []
    for (dev, ch), g in temps.groupby(["Device", "Channel"]):
        typ = g["Type"].iloc[-1]
        out_min = g.loc[g["Status"] == "OUT", "dur"].sum()
        total = g["dur"].sum()
        ch_rows.append({
            "Branch": g["Branch"].iloc[0], "Device": dev, "Channel": ch, "Type": typ,
            "Range": range_text(typ) if typ else "-",
            "Min": round(g["Temp"].min(), 1), "Max": round(g["Temp"].max(), 1), "Avg": round(g["Temp"].mean(), 1),
            "Readings": len(g), "Out of range (min)": round(out_min),
            "Out of range %": round(100 * out_min / total, 1) if total else 0,
        })
    ch_df = pd.DataFrame(ch_rows)
    if not ch_df.empty:
        ch_df = ch_df.sort_values(["Out of range (min)", "Branch"], ascending=[False, True]).reset_index(drop=True)

    off_rows = []
    for dev, g in devstat.groupby("Device"):
        off = g.loc[g["Status"] == "Offline", "dur"].sum()
        nr = g.loc[g["Status"] == "No reading", "dur"].sum()
        if off or nr:
            off_rows.append({"Branch": g["Branch"].iloc[0], "Device": dev,
                             "Offline (min)": round(off), "No reading (min)": round(nr)})
    off_df = pd.DataFrame(off_rows)

    ws_alert = get_tab(sh, ALERTS_TAB, ALERT_HEADER)
    av = ws_alert.get_all_values()
    al_df = pd.DataFrame(av[1:], columns=av[0]) if len(av) > 1 else pd.DataFrame(columns=ALERT_HEADER)
    al_df = al_df[al_df["Time (KSA)"].str.startswith(day_s)]
    dev_al = al_df[al_df["Kind"] != "Temperature"]
    al_df = al_df[al_df["Kind"] == "Temperature"]

    inactive = [s for s in snap if s["status"] == "Inactive" and not s["ignore"]]
    untyped = sorted({f"{r['Device']} / {r['Channel']}" for _, r in ch_df.iterrows() if not r["Type"]}) if not ch_df.empty else []

    breach = ch_df[ch_df["Out of range (min)"] > 0] if not ch_df.empty else ch_df
    healthy = len(ch_df) - len(breach) if not ch_df.empty else 0
    total_out_h = round(breach["Out of range (min)"].sum() / 60, 1) if not breach.empty else 0
    new_alerts = (al_df["Event"] == "NEW").sum()

    # ---------- Excel ----------
    out_path = f"Cold_Chain_Summary_{day_s}.xlsx"
    build_summary_excel(out_path, day_s, ch_df, pd.DataFrame(), al_df, temps, [])
    pdf_path = f"Cold_Chain_Charts_{day_s}.pdf"
    build_charts_pdf(pdf_path, day, temps, ch_df)

    # ---------- Email ----------
    def kpi(label, val, color):
        return (f"<td style='padding:10px 14px;border:1px solid #ddd;text-align:center;'>"
                f"<div style='font-size:22px;font-weight:bold;color:{color};'>{val}</div>"
                f"<div style='font-size:12px;color:#555;'>{label}</div></td>")

    kpis = ("<table style='border-collapse:collapse;margin-bottom:16px;'><tr>"
            + kpi("حساسات تمت مراقبتها", len(ch_df), "#222")
            + kpi("ضمن النطاق طوال اليوم", healthy, "#2E7D32")
            + kpi("خرجت عن النطاق", len(breach), "#C62828")
            + kpi("ساعات خارج النطاق (إجمالي)", total_out_h, "#C62828")
            + kpi("تنبيهات فورية أرسلت", new_alerts, "#EF6C00")
            + "</tr></table>")

    b_rows = [["!" + r["Branch"], r["Channel"], TYPE_AR.get(r["Type"], "-"), r["Range"], r["Min"], r["Max"], r["Avg"],
               fmt_minutes(r["Out of range (min)"]), f"{r['Out of range %']}%"] for _, r in breach.iterrows()]
    ok_rows = [[r["Branch"], r["Channel"], TYPE_AR.get(r["Type"], "-"), r["Min"], r["Max"], r["Avg"]]
               for _, r in (ch_df[ch_df["Out of range (min)"] == 0].iterrows() if not ch_df.empty else [])]

    inner = kpis
    inner += "<h3 style='color:#C62828;margin:6px 0;'>خرجت عن النطاق خلال اليوم</h3>"
    inner += table(["الفرع / الجهاز", "الحساس", "النوع", "النطاق", "أدنى", "أعلى", "المتوسط", "المدة خارج النطاق", "نسبة اليوم"], b_rows)
    inner += "<h3 style='color:#2E7D32;margin:6px 0;'>ضمن النطاق طوال اليوم</h3>"
    inner += table(["الفرع / الجهاز", "الحساس", "النوع", "أدنى", "أعلى", "المتوسط"], ok_rows)
    inner += "<p style='color:#555;font-size:12px;'>الأجهزة غير المتصلة أو بدون قراءة في تقرير منفصل.</p>"
    if untyped:
        inner += ("<p style='color:#555;font-size:12px;'>حساسات بدون نوع محدد (لا يتم فحص النطاق): "
                  + html.escape(", ".join(untyped)) + "</p>")
    inner += f"<p style='color:#555;font-size:12px;'>عدد القراءات المسجلة لليوم: {polls} قراءة لكل جهاز تقريبا. التفاصيل الكاملة والقراءات بالساعة في الملف المرفق.</p>"

    status_word = f"{len(breach)} خارج النطاق" if len(breach) else "الكل ضمن النطاق"
    subject = f"ملخص درجات حرارة الثلاجات ليوم {day_s} | {status_word}"
    send_email(subject, wrap("ملخص اليوم - أجهزة مراقبة الحرارة Elitech",
                             f"اليوم: <span dir='ltr'>{day_s}</span> | من 00:00 إلى 23:59 KSA", inner), [out_path, pdf_path])

    send_device_report(day_s, devstat, dev_al, snap)

    prune_readings(ws_read, values)
    return True


def send_device_report(day_s, devstat, dev_al, snap):
    """Separate daily email: devices that went offline or stopped reading, plus current state."""
    rows = []
    for dev, g in devstat.groupby("Device"):
        off = g.loc[g["Status"] == "Offline", "dur"].sum()
        nr = g.loc[g["Status"] == "No reading", "dur"].sum()
        if off or nr:
            times = int(((dev_al["Device"] == dev) & (dev_al["Event"] == "NEW")).sum()) if not dev_al.empty else 0
            rows.append({"branch": g["Branch"].iloc[0], "off": off, "nr": nr, "times": times})
    rows.sort(key=lambda r: -(r["off"] + r["nr"]))

    now_bad = sorted([s for s in snap if s["status"] in ("Offline", "No reading") and not s["ignore"]],
                     key=lambda s: s["minutes"] or 0)
    inactive = sorted([s for s in snap if s["status"] == "Inactive" and not s["ignore"]],
                      key=lambda s: s["minutes"] or 10**9)
    tracked = sum(1 for s in snap if s["status"] != "Inactive" and not s["ignore"])

    def kpi(label, val, color):
        return (f"<td style='padding:10px 14px;border:1px solid #ddd;text-align:center;'>"
                f"<div style='font-size:22px;font-weight:bold;color:{color};'>{val}</div>"
                f"<div style='font-size:12px;color:#555;'>{label}</div></td>")

    def last_seen(s):
        return s["last"].strftime("%Y-%m-%d %H:%M") if s["last"] else "-"

    def battery(s):
        return f"{s['battery']}%" if s.get("battery") is not None else "-"

    inner = ("<table style='border-collapse:collapse;margin-bottom:16px;'><tr>"
             + kpi("أجهزة تتم مراقبتها", tracked, "#222")
             + kpi("انقطعت أو توقفت قراءتها خلال اليوم", len(rows), "#EF6C00")
             + kpi("غير متصلة / بدون قراءة الآن", len(now_bad), "#C62828")
             + kpi(f"متوقفة أكثر من {INACTIVE_DAYS} أيام", len(inactive), "#777")
             + "</tr></table>")
    inner += "<h3 style='color:#EF6C00;margin:6px 0;'>خلال اليوم</h3>"
    inner += table(["الفرع / الجهاز", "مدة الانقطاع", "مدة بدون قراءة", "عدد المرات"],
                   [["!" + r["branch"], fmt_minutes(r["off"]), fmt_minutes(r["nr"]), r["times"]] for r in rows])
    inner += "<h3 style='color:#C62828;margin:6px 0;'>الحالة الآن</h3>"
    inner += table(["الفرع / الجهاز", "الحالة", "آخر اتصال", "البطارية", "التفاصيل"],
                   [["!" + s["branch"], "غير متصل" if s["status"] == "Offline" else "بدون قراءة",
                     last_seen(s), battery(s),
                     ("FAULT: " + ", ".join(s["faults"])) if s["status"] == "No reading" and s["faults"] else "-"]
                    for s in now_bad])
    inner += f"<h3 style='color:#777;margin:6px 0;'>متوقفة منذ أكثر من {INACTIVE_DAYS} أيام (لا تتم مراقبتها)</h3>"
    inner += table(["الجهاز", "آخر اتصال", "البطارية"], [[s["branch"], last_seen(s), battery(s)] for s in inactive])

    word = f"{len(now_bad)} غير متصل الآن" if now_bad else "كل الأجهزة متصلة"
    send_email(f"تقرير حالة أجهزة المراقبة ليوم {day_s} | {word}",
               wrap("تقرير حالة الأجهزة - انقطاع الاتصال أو القراءة Elitech",
                    f"اليوم: <span dir='ltr'>{day_s}</span> | من 00:00 إلى 23:59 KSA", inner))


def build_charts_pdf(path, day, temps, ch_df):
    """One line chart per branch/device for the whole day: allowed band shaded green,
    out-of-range readings marked red. Branches with problems come first."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    from matplotlib.backends.backend_pdf import PdfPages

    if temps.empty:
        return
    out_by_dev = (ch_df.groupby("Device")["Out of range (min)"].sum().to_dict() if not ch_df.empty else {})
    devices = sorted(temps["Device"].unique(), key=lambda d: (-out_by_dev.get(d, 0), str(temps.loc[temps["Device"] == d, "Branch"].iloc[0])))
    start = datetime.combine(day, datetime.min.time())
    end = start + timedelta(days=1)
    colors = ["#1A73E8", "#8E44AD", "#E67E22", "#16A085"]
    band_color = {"chiller": "#2E7D32", "freezer": "#1565C0"}

    per_page = 3
    with PdfPages(path) as pdf:
        for p in range(0, len(devices), per_page):
            fig, axes = plt.subplots(per_page, 1, figsize=(8.27, 11.69))
            for ax in axes:
                ax.set_visible(False)
            for ax, dev in zip(axes, devices[p:p + per_page]):
                ax.set_visible(True)
                g = temps[temps["Device"] == dev]
                branch = g["Branch"].iloc[0]
                drawn_bands = set()
                for i, (ch, cg) in enumerate(g.groupby("Channel")):
                    cg = cg.sort_values("t")
                    typ = cg["Type"].iloc[-1]
                    lo, hi = TEMP_RANGES.get(typ, (None, None))
                    if typ and typ not in drawn_bands:
                        ymin = lo if lo is not None else -40
                        ymax = hi if hi is not None else 40
                        ax.axhspan(ymin, ymax, color=band_color.get(typ, "#2E7D32"), alpha=0.08)
                        for lim in (lo, hi):
                            if lim is not None:
                                ax.axhline(lim, color=band_color.get(typ, "#2E7D32"), lw=0.8, ls="--")
                        drawn_bands.add(typ)
                    label = f"{ch} ({typ or 'no type'}, range {range_text(typ)})"
                    ax.plot(cg["t"], cg["Temp"], color=colors[i % len(colors)], lw=1.4, label=label)
                    bad = cg[cg["Status"] == "OUT"]
                    if not bad.empty:
                        ax.scatter(bad["t"], bad["Temp"], color="#C62828", s=14, zorder=3)
                mins = out_by_dev.get(dev, 0)
                title = f"{branch}  -  out of range {fmt_minutes(mins)}" if mins else f"{branch}  -  in range all day"
                ax.set_title(title, fontsize=11, loc="left", color="#C62828" if mins else "#222222", fontweight="bold")
                vals = list(g["Temp"].dropna())
                limits = [x for typ in g["Type"].unique() for x in TEMP_RANGES.get(typ, (None, None)) if x is not None]
                ax.set_ylim(min(vals + limits) - 2, max(vals + limits) + 2)
                ax.set_xlim(start, end)
                ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
                ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))
                ax.set_ylabel("C")
                ax.grid(alpha=0.25)
                ax.legend(fontsize=7, loc="upper right")
            fig.suptitle(f"Cold chain temperatures - {day:%Y-%m-%d} (KSA)   red dots = out of range", fontsize=10, color="#555555")
            fig.tight_layout(rect=[0, 0, 1, 0.98])
            pdf.savefig(fig)
            plt.close(fig)
    print(f"PDF saved: {path}")


def build_summary_excel(path, day_s, ch_df, off_df, al_df, temps, inactive):
    HEAD = PatternFill("solid", fgColor="2C3E50")
    HFONT = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    RED = PatternFill("solid", fgColor="FFCCCC")
    GREEN = PatternFill("solid", fgColor="CCFFCC")
    CENTER = Alignment(vertical="center", horizontal="center", wrap_text=True)
    thin = Side(style="thin", color="CCCCCC")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

    def write(ws, df, red_rule=None, widths=None):
        cols = list(df.columns)
        for ci, c in enumerate(cols, 1):
            cell = ws.cell(1, ci, c)
            cell.fill, cell.font, cell.alignment, cell.border = HEAD, HFONT, CENTER, BORDER
            ws.column_dimensions[get_column_letter(ci)].width = (widths or {}).get(c, max(12, len(str(c)) + 2))
        for ri, (_, r) in enumerate(df.iterrows(), 2):
            bad = red_rule(r) if red_rule else False
            for ci, c in enumerate(cols, 1):
                v = r[c]
                cell = ws.cell(ri, ci, None if (isinstance(v, float) and pd.isna(v)) else v)
                cell.alignment, cell.border = CENTER, BORDER
                if bad:
                    cell.fill = RED
        ws.freeze_panes = "A2"

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    if not ch_df.empty:
        write(ws, ch_df, lambda r: r["Out of range (min)"] > 0, {"Branch": 22, "Device": 18, "Channel": 16})

    # Hourly: one row per probe, one column per hour.
    # If ANY reading in that hour was out of range, the cell shows the worst reading in red;
    # otherwise it shows the hour's average in green. A 10-minute spike can't hide in an average.
    ws_h = wb.create_sheet("Hourly")
    if not temps.empty:
        t = temps.copy()
        t["Hour"] = t["t"].dt.hour
        hdr = ["Branch", "Channel", "Type"] + [f"{h:02d}:00" for h in range(24)]
        for ci, h in enumerate(hdr, 1):
            c = ws_h.cell(1, ci, h)
            c.fill, c.font, c.alignment, c.border = HEAD, HFONT, CENTER, BORDER
            ws_h.column_dimensions[get_column_letter(ci)].width = 22 if ci == 1 else (14 if ci <= 3 else 8)
        ri = 2
        for (br, ch, typ), g in t.groupby(["Branch", "Channel", "Type"]):
            for ci, v in enumerate([br, ch, typ], 1):
                c = ws_h.cell(ri, ci, v)
                c.alignment, c.border = CENTER, BORDER
            for h in range(24):
                vals = g.loc[g["Hour"] == h, "Temp"].dropna()
                c = ws_h.cell(ri, 4 + h)
                c.alignment, c.border = CENTER, BORDER
                if vals.empty:
                    continue
                bad = [v for v in vals if typ and out_of_range(v, typ)]
                if bad:
                    c.value = round(worst_reading(bad, typ), 1)
                    c.fill, c.font = RED, Font(bold=True, color="9C0006", name="Arial", size=10)
                else:
                    c.value = round(float(vals.mean()), 1)
                    if typ:
                        c.fill = GREEN
            ri += 1
        note = ri + 1
        ws_h.cell(note, 1, "Red = at least one reading that hour was out of range (worst reading shown). "
                           "Green = all readings in range (hourly average shown).").font = Font(italic=True, color="666666")
        ws_h.freeze_panes = "D2"

    if not off_df.empty:
        ws_o = wb.create_sheet("Offline")
        write(ws_o, off_df, None, {"Branch": 22, "Device": 18})

    ws_a = wb.create_sheet("Alerts")
    if not al_df.empty:
        write(ws_a, al_df.reset_index(drop=True), lambda r: r["Event"] == "NEW", {"Detail": 40, "Branch": 22})

    if inactive:
        ws_i = wb.create_sheet("Inactive devices")
        write(ws_i, pd.DataFrame([{"Device": s["device"],
                                   "Last seen (KSA)": s["last"].strftime("%Y-%m-%d %H:%M") if s["last"] else "",
                                   "Battery %": s["battery"]} for s in inactive]), None, {"Device": 30})

    wb.save(path)
    print(f"Excel saved: {path}")


def prune_readings(ws, values):
    cutoff = (NOW - timedelta(days=KEEP_DAYS)).strftime(FMT)
    old = sum(1 for r in values[1:] if r and r[0] < cutoff)
    if old:
        ws.delete_rows(2, old + 1)
        print(f"Pruned {old} readings older than {KEEP_DAYS} days.")



if __name__ == "__main__":
    sh, ws_state, state = open_store()
    snap = snapshot(fetch_devices(get_token()))
    run_summary(snap, sh, state)
    print("Done.")

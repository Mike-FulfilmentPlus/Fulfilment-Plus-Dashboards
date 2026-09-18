"""
Desk365 Customer Services Visibility Report — GitHub Actions edition.

Fetches all tickets from the Desk365 API, computes the same metrics as the
original Claude-hosted artifact (SLA compliance, overdue alerts, volume by
day, category trend, per-customer table), and writes a single self-contained
static HTML file (desk365_visibility_report.html) for GitHub Pages.

Runs unattended on a schedule via .github/workflows/desk365-report.yml — no
AI session involved. Requires one secret in the repo's Actions settings:

    DESK365_API_KEY   Desk365 API key (Settings -> Secrets and variables ->
                       Actions -> New repository secret). Never commit this
                       value or paste it in chat/issues/PRs.

Usage:
    DESK365_API_KEY=xxx python desk365_report.py
"""
import os
import sys
import json
import html
import datetime
import statistics
import requests

DESK365_BASE_URL = "https://fulfilmentplus.desk365.io/apis/v3/tickets"
OUTPUT_FILE = "desk365_visibility_report.html"
NZ_OFFSET_HOURS = 12  # NZST (no DST currently in effect)
CHARTJS_CDN = "https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.5.0/chart.umd.min.js"


# ── Fetch ────────────────────────────────────────────────────────────────

def fetch_all_tickets(api_key):
    tickets = []
    offset = 0
    while True:
        resp = requests.get(
            DESK365_BASE_URL,
            params={
                "ticket_count": 100,
                "order_by": "created_time",
                "order_type": "desc",
                "include_custom_fields": 1,
                "offset": offset,
            },
            headers={"Authorization": api_key},
            timeout=60,
        )
        resp.raise_for_status()
        data = resp.json()
        page = data.get("tickets", [])
        tickets.extend(page)
        count = data.get("count", len(tickets))
        if not page or offset + len(page) >= count:
            break
        offset += 100
    return tickets


# ── Compute ──────────────────────────────────────────────────────────────

def parse_dt(dt_str):
    if not dt_str:
        return None
    return datetime.datetime.strptime(dt_str, "%Y-%m-%d %H:%M:%S").replace(tzinfo=datetime.timezone.utc)


def compute_report(tickets_raw, now=None):
    NOW = now or datetime.datetime.now(datetime.timezone.utc)

    tickets = []
    for t in tickets_raw:
        created = parse_dt(t["created_on"])
        resolved = parse_dt(t["resolved_on"])
        closed = parse_dt(t["closed_on"])
        due = parse_dt(t["due_date"])

        customer = (t.get("custom_fields") or {}).get("cf_Customer") or t.get("company_name") or "Unassigned"

        if not due:
            sla_status = "na"
        else:
            end = resolved or closed
            if end:
                sla_status = "met" if end <= due else "breached"
            else:
                sla_status = "breached_open" if NOW > due else "on_track"

        tickets.append({
            "ticket_number": t["ticket_number"],
            "subject": t["subject"],
            "status": t["status"],
            "customer": customer,
            "category": t.get("category") or "Uncategorized",
            "created": created,
            "resolved": resolved,
            "closed": closed,
            "due": due,
            "sla_status": sla_status,
            "first_replied_duration": t.get("first_replied_duration"),
            "resolved_duration": t.get("resolved_duration"),
        })

    total = len(tickets)
    closed_statuses = {"Closed", "Resolved"}
    open_count = sum(1 for t in tickets if t["status"] not in closed_statuses)
    closed_count = sum(1 for t in tickets if t["status"] in closed_statuses)

    met = sum(1 for t in tickets if t["sla_status"] == "met")
    breached = sum(1 for t in tickets if t["sla_status"] == "breached")
    breached_open = sum(1 for t in tickets if t["sla_status"] == "breached_open")
    on_track = sum(1 for t in tickets if t["sla_status"] == "on_track")
    na = sum(1 for t in tickets if t["sla_status"] == "na")
    measurable = met + breached
    sla_compliance = (met / measurable * 100) if measurable else None

    res_durations = [t["resolved_duration"] for t in tickets if t["status"] in closed_statuses and isinstance(t["resolved_duration"], (int, float))]
    median_resolution_min = statistics.median(res_durations) if res_durations else None
    if len(res_durations) >= 4:
        sorted_durs = sorted(res_durations)
        q1 = statistics.median(sorted_durs[:len(sorted_durs)//2])
        q3 = statistics.median(sorted_durs[(len(sorted_durs)+1)//2:])
        iqr = q3 - q1
        outlier_threshold = q3 + 1.5 * iqr
        outliers = sum(1 for d in res_durations if d > outlier_threshold)
    else:
        outliers = 0

    reply_durations = [t["first_replied_duration"] for t in tickets if isinstance(t["first_replied_duration"], (int, float))]
    median_reply_min = statistics.median(reply_durations) if reply_durations else None

    created_dates = [t["created"] for t in tickets if t["created"]]
    date_range_str = f"{min(created_dates).strftime('%d %b %Y')} – {max(created_dates).strftime('%d %b %Y')}" if created_dates else ""

    overdue = []
    for t in tickets:
        if t["sla_status"] == "breached_open":
            days_overdue = (NOW - t["due"]).total_seconds() / 86400
            overdue.append((days_overdue, t))
    overdue.sort(key=lambda x: -x[0])

    day_labels, day_counts = [], []
    for i in range(13, -1, -1):
        d = (NOW - datetime.timedelta(days=i)).date()
        day_labels.append(d.strftime("%-d %b"))
        day_counts.append(sum(1 for t in tickets if t["created"] and t["created"].date() == d))

    companies = sorted(set(t["customer"] for t in tickets if t["customer"] and t["customer"] != "Unassigned"))
    company_stats = []
    for name in companies:
        rows = [t for t in tickets if t["customer"] == name]
        open_pending = sum(1 for t in rows if t["status"] not in closed_statuses)
        closed_this_month = sum(1 for t in rows if t["closed"] and t["closed"].year == NOW.year and t["closed"].month == NOW.month)
        replies = [t["first_replied_duration"] for t in rows if isinstance(t["first_replied_duration"], (int, float))]
        avg_reply = (sum(replies) / len(replies)) if replies else None
        company_stats.append({
            "name": name, "total": len(rows), "open_pending": open_pending,
            "closed_this_month": closed_this_month, "avg_reply": avg_reply, "avg_reply_n": len(replies),
        })
    company_stats.sort(key=lambda c: (-c["total"], c["name"]))

    WEEKS = 52
    week_labels = [(NOW - datetime.timedelta(weeks=i)).strftime("%d %b") for i in range(WEEKS - 1, -1, -1)]

    def week_index(dt):
        weeks_ago = (NOW - dt).days // 7
        idx = WEEKS - 1 - weeks_ago
        return idx if 0 <= idx < WEEKS else None

    categories = sorted(set(t["category"] for t in tickets))
    category_series = []
    for cat in categories:
        data = [0] * WEEKS
        for t in tickets:
            if t["category"] == cat and t["created"]:
                idx = week_index(t["created"])
                if idx is not None:
                    data[idx] += 1
        category_series.append({"label": cat, "data": data})

    snapshot_local = NOW + datetime.timedelta(hours=NZ_OFFSET_HOURS)
    snapshot_str = snapshot_local.strftime("%Y-%m-%d %H:%M") + " NZST"

    return {
        "total": total, "open_count": open_count, "closed_count": closed_count,
        "date_range_str": date_range_str,
        "met": met, "breached": breached, "breached_open": breached_open, "on_track": on_track, "na": na,
        "measurable": measurable, "sla_compliance": sla_compliance,
        "median_resolution_min": median_resolution_min, "resolved_n": len(res_durations), "outliers": outliers,
        "median_reply_min": median_reply_min, "reply_n": len(reply_durations),
        "overdue": [{"days": d, "ticket_number": t["ticket_number"], "subject": t["subject"],
                     "customer": t["customer"], "category": t["category"]} for d, t in overdue],
        "day_labels": day_labels, "day_counts": day_counts,
        "company_stats": company_stats,
        "week_labels": week_labels, "category_series": category_series,
        "snapshot_str": snapshot_str,
    }


# ── Render ───────────────────────────────────────────────────────────────

def esc(s):
    return html.escape(str(s), quote=True)


def fmt_hrs(minutes):
    return f"{minutes/60:.1f} hrs"


def render_html(d):
    if d["overdue"]:
        parts = []
        for o in d["overdue"]:
            parts.append(f'Ticket #{o["ticket_number"]} (&quot;{esc(o["subject"])}&quot;, {esc(o["customer"])}, {esc(o["category"])}) &mdash; <b>{o["days"]:.1f} days overdue</b>')
        alert_html = f'<div class="alert"><b>SLA breach alert:</b> {len(d["overdue"])} ticket{"s" if len(d["overdue"]) != 1 else ""} overdue and still open &mdash; ' + "; ".join(parts) + '.</div>'
    else:
        alert_html = '<div class="alert alert-ok"><b>No overdue tickets.</b> Every open ticket is still inside its SLA window.</div>'

    sla_pct = f'{d["sla_compliance"]:.1f}%' if d["sla_compliance"] is not None else 'n/a'
    sla_class = 'good' if (d["sla_compliance"] or 0) >= 90 else ('warn' if (d["sla_compliance"] or 0) >= 75 else 'bad')

    median_res = fmt_hrs(d["median_resolution_min"]) if d["median_resolution_min"] is not None else "n/a"
    median_reply = fmt_hrs(d["median_reply_min"]) if d["median_reply_min"] is not None else "n/a"

    company_rows = "".join(f'''
      <tr>
        <td>{esc(c["name"])}</td>
        <td class="num">{c["total"]}</td>
        <td class="num">{c["open_pending"]}</td>
        <td class="num">{c["closed_this_month"]}</td>
        <td class="num">{fmt_hrs(c["avg_reply"]) + f' (n={c["avg_reply_n"]})' if c["avg_reply"] is not None else '&mdash;'}</td>
      </tr>''' for c in d["company_stats"])

    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Desk365 Visibility Report - Fulfilment Plus</title>
<script src="{CHARTJS_CDN}"></script>
<style>
  :root {{
    --bg: #f6f7f9; --surface: #ffffff; --border: #e4e7ec; --text: #1a1d23;
    --muted: #667085; --faint: #98a2b3; --accent: #2f6fed;
    --good: #17824e; --good-bg: #e7f6ee; --warn: #b45309; --warn-bg: #fef3e2;
    --bad: #c0362c; --bad-bg: #fbebe9; --alert-border: #f0b877; --alert-bg: #fdf3e4; --alert-text: #7a4a12;
    --ok-border: #9fd3b6; --ok-bg: #eaf7f0; --ok-text: #1f6b46;
  }}
  * {{ box-sizing: border-box; }}
  body {{ margin: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; background: var(--bg); color: var(--text); padding: 24px 16px; }}
  .wrap {{ max-width: 1180px; margin: 0 auto; }}
  header {{ display: flex; justify-content: space-between; align-items: flex-end; margin-bottom: 20px; flex-wrap: wrap; gap: 8px; }}
  h1 {{ font-size: 22px; margin: 0 0 4px 0; }}
  .sub {{ color: var(--muted); font-size: 13px; }}
  .asof {{ font-size: 12px; color: var(--faint); text-align: right; }}
  .cards {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 12px; margin-bottom: 10px; }}
  .card {{ background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 14px 16px; }}
  .card .label {{ font-size: 12px; color: var(--muted); margin-bottom: 6px; }}
  .card .value {{ font-size: 26px; font-weight: 600; font-variant-numeric: tabular-nums; }}
  .card .value.warn {{ color: var(--warn); }}
  .card .value.good {{ color: var(--good); }}
  .card .value.bad {{ color: var(--bad); }}
  .card .note {{ font-size: 11px; color: var(--faint); margin-top: 4px; }}
  .alert {{ background: var(--alert-bg); border: 1px solid var(--alert-border); color: var(--alert-text); border-radius: 10px; padding: 12px 16px; margin-bottom: 22px; font-size: 13px; }}
  .alert b {{ color: inherit; }}
  .alert-ok {{ background: var(--ok-bg); border-color: var(--ok-border); color: var(--ok-text); }}
  .grid2 {{ display: grid; grid-template-columns: 1.3fr 1fr; gap: 16px; margin-bottom: 16px; }}
  .panel {{ background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 16px; margin-bottom: 16px; }}
  .panel h2 {{ font-size: 14px; margin: 0 0 12px 0; color: var(--text); }}
  .panel h2 .hint {{ font-weight: 400; color: var(--faint); font-size: 11.5px; }}
  .chart-box {{ position: relative; height: 240px; }}
  .chart-box.wide {{ height: 260px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 12.5px; }}
  th, td {{ text-align: left; padding: 8px 10px; border-bottom: 1px solid var(--border); white-space: nowrap; }}
  th {{ color: var(--muted); font-weight: 600; font-size: 11.5px; text-transform: uppercase; letter-spacing: .02em; }}
  tr:last-child td {{ border-bottom: none; }}
  td.num {{ font-variant-numeric: tabular-nums; }}
  footer {{ margin-top: 24px; font-size: 11.5px; color: var(--faint); text-align: center; }}
  @media (max-width: 800px) {{ .grid2 {{ grid-template-columns: 1fr; }} table {{ font-size: 11px; }} th, td {{ padding: 6px; }} }}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div>
      <h1>Customer Services Visibility Report</h1>
      <div class="sub">Desk365 helpdesk &mdash; fulfilmentplus.desk365.io</div>
    </div>
    <div class="asof">Snapshot as of<br><strong>{esc(d["snapshot_str"])}</strong></div>
  </header>
  {alert_html}
  <div class="cards">
    <div class="card"><div class="label">Total tickets</div><div class="value">{d["total"]}</div><div class="note">{esc(d["date_range_str"])}</div></div>
    <div class="card"><div class="label">Open / Pending</div><div class="value warn">{d["open_count"]}</div><div class="note">{len(d["overdue"])} breached (overdue)</div></div>
    <div class="card"><div class="label">Closed</div><div class="value good">{d["closed_count"]}</div><div class="note">{(d["closed_count"]/d["total"]*100) if d["total"] else 0:.1f}% of volume</div></div>
    <div class="card"><div class="label">SLA compliance</div><div class="value {sla_class}">{sla_pct}</div><div class="note">{d["met"]} met / {d["measurable"]} measurable</div></div>
    <div class="card"><div class="label">Median 1st response time</div><div class="value">{median_reply}</div><div class="note">n={d["reply_n"]} of {d["total"]} tickets tracked</div></div>
    <div class="card"><div class="label">Median resolution time</div><div class="value">{median_res}</div><div class="note">n={d["resolved_n"]} &middot; {d["outliers"]} outliers flagged</div></div>
  </div>
  <div class="grid2">
    <div class="panel"><h2>Ticket volume by day (created)</h2><div class="chart-box"><canvas id="volumeChart"></canvas></div></div>
    <div class="panel"><h2>SLA compliance</h2><div class="chart-box"><canvas id="slaChart"></canvas></div></div>
  </div>
  <div class="panel"><h2>Tickets by category <span class="hint">&mdash; last 12 months, weekly</span></h2><div class="chart-box wide"><canvas id="categoryChart"></canvas></div></div>
  <div class="panel">
    <h2>Tickets by customer <span class="hint">&mdash; ordered by total tickets</span></h2>
    <div style="overflow-x:auto;">
    <table>
      <thead><tr><th>Customer</th><th>Total</th><th>Open / Pending (current)</th><th>Closed this month</th><th>Avg 1st response time</th></tr></thead>
      <tbody id="companyRows">{company_rows}</tbody>
    </table>
    </div>
  </div>
  <footer>Point-in-time snapshot pulled directly from the Desk365 API by a scheduled GitHub Actions workflow. Customer table grouped by the Desk365 "Customer" custom field (cf_Customer), falling back to company name only for the handful of tickets where it isn't set yet.</footer>
</div>
<script>
const volumeLabels = {json.dumps(d["day_labels"])};
const volumeData = {json.dumps(d["day_counts"])};
new Chart(document.getElementById("volumeChart"), {{ type: "bar", data: {{ labels: volumeLabels, datasets: [{{ label: "Tickets created", data: volumeData, backgroundColor: "#2f6fed", borderRadius: 4 }}] }}, options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ display: false }} }}, scales: {{ y: {{ beginAtZero: true, ticks: {{ stepSize: 1 }} }} }} }} }});
new Chart(document.getElementById("slaChart"), {{ type: "doughnut", data: {{ labels: ["Met", "Breached (closed late)", "Breached (open, overdue)", "On track", "No SLA policy"], datasets: [{{ data: [{d["met"]}, {d["breached"]}, {d["breached_open"]}, {d["on_track"]}, {d["na"]}], backgroundColor: ["#17824e", "#c0362c", "#7c2d12", "#2f6fed", "#98a2b3"] }}] }}, options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: "bottom", labels: {{ boxWidth: 12, font: {{ size: 11 }} }} }} }} }} }});
const palette = ["#2f6fed", "#7c3aed","#0891b2","#d97706","#dc2626","#16a34a","#db2777","#4f46e5"];
const weekLabels = {json.dumps(d["week_labels"])};
const categorySeries = {json.dumps(d["category_series"])}.map((c, i) => ({{ label: c.label, data: c.data, borderColor: palette[i % palette.length], backgroundColor: palette[i % palette.length], tension: 0.25, pointRadius: 2, borderWidth: 2 }}));
new Chart(document.getElementById("categoryChart"), {{ type: "line", data: {{ labels: weekLabels, datasets: categorySeries }}, options: {{ responsive: true, maintainAspectRatio: false, interaction: {{ mode: 'index', intersect: false }}, plugins: {{ legend: {{ position: 'bottom', labels: {{ boxWidth: 10, font: {{ size: 11 }} }} }} }}, scales: {{ y: {{ beginAtZero: true, ticks: {{ stepSize: 1 }} }}, x: {{ ticks: {{ maxTicksLimit: 14, autoSkip: true }} }} }} }} }});
</script>
</body>
</html>
'''


def main():
    api_key = os.environ.get("DESK365_API_KEY")
    if not api_key:
        print("ERROR: DESK365_API_KEY environment variable is not set.", file=sys.stderr)
        sys.exit(1)

    print("Fetching tickets from Desk365...")
    tickets_raw = fetch_all_tickets(api_key)
    print(f"  fetched {len(tickets_raw)} tickets")

    report = compute_report(tickets_raw)
    print(f"Total: {report['total']}  Open: {report['open_count']}  Closed: {report['closed_count']}")
    print(f"SLA compliance: {report['sla_compliance']} ({report['met']}/{report['measurable']})")
    if report["overdue"]:
        print(f"Overdue breached-open: {len(report['overdue'])}")
        for o in report["overdue"]:
            print(f"  #{o['ticket_number']} {o['subject']!r} {o['customer']} - {o['days']:.1f}d overdue")

    html_out = render_html(report)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html_out)
    print(f"Wrote {OUTPUT_FILE} ({len(html_out)} bytes)")


if __name__ == "__main__":
    main()

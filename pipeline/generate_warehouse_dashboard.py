"""
Warehouse Operations Dashboard Generator

Fetches live data from CartonCloud and writes warehouse_dashboard.html —
a self-contained, interactive HTML file showing:

  SALES ORDERS
  - To be picked today  (all orders currently in AWAITING_PICK_AND_PACK)
  - Outstanding & late  (AWAITING_PICK_AND_PACK past their KPI dispatch deadline)
  - Completed today     (dispatched today NZ time)

  PURCHASE ORDERS
  - Expected but not arrived   (arrivalDate >= today, not yet verified)
  - In receipt                 (arrivalDate = today or yesterday, not yet verified)
  - Open & failing KPI         (arrivalDate > 48 hrs ago, not yet verified)

  12-MONTH DAILY VOLUME CHART
  - Daily orders / lines / pick units for the last 365 days

Click any KPI card to drill into the full order table.

Usage:
    cd "C:\\Users\\MikeAppleton\\Claude\\Projects\\Fulfilment Plus KPIs"
    python generate_warehouse_dashboard.py

Optionally schedule this with Task Scheduler to auto-refresh throughout the day.
"""

import csv
import json
import os
import datetime
import re
import sys
import requests
import openpyxl
from zoneinfo import ZoneInfo

import volume_data

# ── Config ────────────────────────────────────────────────────────────────────
CREDS_FILE   = "credentials.env"
WORKBOOK     = "kpi_dashboard.xlsx"
OUTPUT_FILE  = "warehouse_dashboard.html"
AWAITING_STOCK_LOG = "awaiting_stock_log.csv"
# Volume cache file/refresh-window constants now live in volume_data.py
# (shared with mtd_kpis.py / extract_kpi_data.py) - see volume_data.VOLUME_CACHE.
NZ_TZ        = ZoneInfo("Pacific/Auckland")
CUTOFF_HOUR  = 14          # 2 pm NZ – standard order cut-off
COB_HOUR     = 20          # 8 pm NZ – close-of-business dispatch deadline (matches
                           # mtd_kpis.py / extract_kpi_data.py, which use COB_HOUR
                           # for the actual dispatch deadline, not cutoff_hour)
REQUIRED_DATE_HOUR = 17    # 5 pm NZ – COB deadline for required-date-derived pick days
INBOUND_KPI_HOURS = 48     # dock-to-stock target window (hours)
TEST_ACCOUNT_NAMES = {"TEST ACCOUNT"}

# ── Rural / SI detection for required-date deadline logic ─────────────────────
# Mirrors extract_kpi_data.py / mtd_kpis.py: transit time before a customer's
# required *delivery* date is 1 business day for North Island destinations,
# 3 for South Island (postcode >= 7000), +1 more if rural. Previously this
# script used a flat 2 business days regardless of destination, which
# understated how much time the warehouse actually has for NI/urban orders.
RURAL_POSTCODE_FILE = "nz_rural_postcodes.json"
RURAL_STREET_FILE   = "NZ Street File - Zone Guide.xlsx"
_RURAL_RE = re.compile(r'\bR\.?D\.?\s*\d', re.IGNORECASE)
_STREET_ABBREV = [(' rd',' road'),(' st',' street'),(' ave',' avenue'),
                  (' cres',' crescent'),(' dr',' drive'),(' tce',' terrace')]


def _norm_street_co(s):
    s = s.lower().strip()
    for a, b in _STREET_ABBREV:
        if s.endswith(a):
            return s[:-len(a)] + b
    return s


def _extract_street_co(addr):
    s = re.sub(r'^\d+[a-zA-Z]?\s+', '', addr.strip())
    return _norm_street_co(s.split(',')[0])


def load_rural_postcode_set_co(path=RURAL_POSTCODE_FILE):
    try:
        with open(path) as f:
            return set(json.load(f)["postcodes"])
    except Exception:
        return set()


def load_rural_street_lookup_co(path=RURAL_STREET_FILE):
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.active
        lookup = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) < 11 or not row[10]:
                continue
            flag = str(row[10]).strip().lower()
            if flag not in ("rural", "rural / non-urban", "non-urban"):
                continue
            pc = str(row[3] or "").strip().zfill(4)
            st = _extract_street_co(str(row[0] or ""))
            if pc and st:
                lookup.add((pc, st))
        wb.close()
        return lookup
    except Exception:
        return set()


def is_rural_co(postcode, street, city, rural_pcs, rural_streets):
    pc = str(postcode or "").strip().zfill(4)
    if pc in rural_pcs:
        return True
    st = _extract_street_co(str(street or ""))
    if st and (pc, st) in rural_streets:
        return True
    combined = f"{street} {city}"
    if _RURAL_RE.search(combined):
        return True
    return False


def is_si_co(postcode):
    try:
        return int(str(postcode or "").strip()) >= 7000
    except Exception:
        return False
# Re-fetch today + yesterday on every run. Everything older is read from cache.
VOLUME_REFRESH_DAYS = 2

# NZ public holidays + Auckland Anniversary — treated as non-business days.
# Update each year using https://www.employment.govt.nz/leave-and-holidays/public-holidays/public-holidays-and-anniversary-dates
PUBLIC_HOLIDAYS = {
    # 2026
    datetime.date(2026,  1,  1),   # New Year's Day
    datetime.date(2026,  1,  2),   # Day after New Year's Day
    datetime.date(2026,  1, 26),   # Auckland Anniversary
    datetime.date(2026,  2,  6),   # Waitangi Day
    datetime.date(2026,  4,  3),   # Good Friday
    datetime.date(2026,  4,  6),   # Easter Monday
    datetime.date(2026,  4, 27),   # ANZAC Day (observed)
    datetime.date(2026,  6,  1),   # King's Birthday
    datetime.date(2026,  7, 10),   # Matariki
    datetime.date(2026, 10, 26),   # Labour Day
    datetime.date(2026, 12, 25),   # Christmas Day
    datetime.date(2026, 12, 28),   # Boxing Day (observed)
}

# Orders whose customer name was wrong in CartonCloud — map order ref → correct customer
# so the right cut-off exceptions are applied.
ORDER_CUSTOMER_OVERRIDES = {
    "1752": "Farmers",  # vixxen order incorrectly named; should follow Farmers cut-off
}

WEEKDAY_NAMES = {
    "monday": 0, "tuesday": 1, "wednesday": 2, "thursday": 3,
    "friday": 4, "saturday": 5, "sunday": 6,
}
CONTENT_MATCH_FIELDS = {
    "Delivery Company Name": lambda o: (
        o.get("details", {}).get("deliver", {}).get("address", {}).get("companyName", "")
    ),
    # Matches every order for the customer, regardless of content - used for
    # blanket customer-wide holds (e.g. a customs/stock embargo affecting all
    # of that customer's orders in a date window). "Contains" should be "all".
    "All Orders": lambda o: "all",
}


# ── Credentials ───────────────────────────────────────────────────────────────
def load_credentials(path=CREDS_FILE):
    creds = {}
    with open(path) as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, val = line.split("=", 1)
            creds[key.strip()] = val.strip()
    return creds


def get_cc_token(creds):
    resp = requests.post(
        "https://api.cartoncloud.com/uaa/oauth/token",
        auth=(creds["CARTONCLOUD_CLIENT_ID"], creds["CARTONCLOUD_CLIENT_SECRET"]),
        headers={"Accept-Version": "1"},
        data={"grant_type": "client_credentials"},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


# ── CartonCloud helpers ────────────────────────────────────────────────────────
def cc_headers(token):
    return {
        "Accept-Version": "1",
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


def search_all_pages(token, tenant_id, resource, condition, size=100):
    """POST /tenants/{tenant}/resource/search with automatic pagination."""
    results = []
    page = 1
    url = f"https://api.cartoncloud.com/tenants/{tenant_id}/{resource}/search"
    while True:
        resp = requests.post(
            url,
            params={"size": size, "page": page},
            headers=cc_headers(token),
            json={"condition": condition},
            timeout=60,
        )
        resp.raise_for_status()
        batch = resp.json()
        results.extend(batch)
        total_pages = int(resp.headers.get("total-pages", "1"))
        if page >= total_pages or not batch:
            break
        page += 1
    return results


def status_condition(status_value):
    return {
        "type": "AndCondition",
        "conditions": [
            {
                "type": "TextComparisonCondition",
                "field": {"type": "JsonField", "pointer": "/status"},
                "value": {"type": "ValueField", "value": status_value},
                "method": "EQUAL_TO",
            }
        ],
    }


def date_range_condition(field_pointer, start_date, end_date):
    return {
        "type": "AndCondition",
        "conditions": [
            {
                "type": "DateComparisonCondition",
                "field": {"type": "JsonField", "pointer": field_pointer},
                "value": {"type": "ValueField", "value": start_date},
                "method": "GREATER_THAN_OR_EQUAL_TO",
            },
            {
                "type": "DateComparisonCondition",
                "field": {"type": "JsonField", "pointer": field_pointer},
                "value": {"type": "ValueField", "value": end_date},
                "method": "LESS_THAN",
            },
        ],
    }


def and_condition(*conditions):
    flat = [c for c in conditions if c]
    if len(flat) == 1:
        return flat[0]
    return {"type": "AndCondition", "conditions": flat}


# ── Order field helpers ────────────────────────────────────────────────────────
def parse_iso(ts):
    return datetime.datetime.fromisoformat(ts)


def order_items(o):
    """Return the list of order items (CartonCloud uses 'items' not 'lines')."""
    return o.get("items") or []


def order_line_count(o):
    """Number of distinct SKU lines on the order."""
    return len(order_items(o))


def order_unit_count(o):
    """Total pick units across all items (measures.quantity)."""
    return sum(
        int(item.get("measures", {}).get("quantity") or 0)
        for item in order_items(o)
    )


def order_ref(o):
    return str(o.get("references", {}).get("numericId") or o.get("id", ""))


def order_customer(o):
    return o.get("customer", {}).get("name", "")


def order_status(o):
    return o.get("status", "")


def order_created_nz(o):
    ts = o.get("timestamps", {}).get("created", {}).get("time")
    return parse_iso(ts).astimezone(NZ_TZ) if ts else None


def order_modified_nz(o):
    ts = o.get("timestamps", {}).get("modified", {}).get("time")
    return parse_iso(ts).astimezone(NZ_TZ) if ts else None


def order_has_error(o):
    return bool(o.get("details", {}).get("errors"))


HOLD_GAP_HOURS = 2  # created→modified gap (with an error present) treated as "on hold"


def load_awaiting_stock_transitions(path=AWAITING_STOCK_LOG):
    """Return {order_id: transition_date} for orders that were held in
    AWAITING_STOCK/DRAFT and later transitioned to a real in-play status,
    per daily_status_snapshot.py's log. Same shared log/authoritative
    source used by mtd_kpis.py and extract_kpi_data.py for their Order
    Cut Off clock-start adjustment - kept in sync here so an order that
    was genuinely held up on stock shows the correct "available from"
    date instead of looking late from the moment it re-surfaces."""
    transitions = {}
    if not os.path.exists(path):
        return transitions
    with open(path, newline="") as f:
        for row in csv.DictReader(f):
            transition_date = (row.get("transition_date") or "").strip()
            if not transition_date:
                continue
            try:
                transitions[row["order_id"]] = datetime.date.fromisoformat(transition_date)
            except ValueError:
                continue
    return transitions


def order_kpi_clock_start(o, stock_transitions=None):
    """Datetime the pick KPI clock should start from.

    Normally this is order creation. But CartonCloud doesn't expose
    status-change history, so an order that was held up (e.g. a stock
    shortage on one line) can sit for days before it's genuinely
    available to pick — using its creation time would make it look
    "late" the moment it surfaces, even though the warehouse never had a
    chance to pick it.

    Primary source: the shared awaiting_stock_log.csv (see
    daily_status_snapshot.py) - the same authoritative transition-date
    log used by mtd_kpis.py / extract_kpi_data.py for Order Cut Off. If
    this order was logged as AWAITING_STOCK/DRAFT and has since
    transitioned to a real in-play status, its clock starts at 00:00 NZ
    on the transition date, not its original creation time.

    Fallback heuristic (for orders the daily snapshot hasn't captured a
    transition for yet): if the order currently carries a fulfilment
    error AND its last-modified time is well after its created time,
    treat `modified` (the closest proxy for "became available to pick")
    as the clock start instead. Orders without an error, or where
    modified ≈ created, are unaffected.
    """
    created_dt = order_created_nz(o)
    if not created_dt:
        return None

    if stock_transitions:
        transition_date = stock_transitions.get(o.get("id"))
        if transition_date and transition_date > created_dt.date():
            return datetime.datetime.combine(
                transition_date, datetime.time(0, 0, 0), tzinfo=NZ_TZ
            )

    if not order_has_error(o):
        return created_dt
    modified_dt = order_modified_nz(o)
    if modified_dt and (modified_dt - created_dt) > datetime.timedelta(hours=HOLD_GAP_HOURS):
        return modified_dt
    return created_dt


def order_dispatched_nz(o):
    ts = o.get("timestamps", {}).get("dispatched", {}).get("time")
    return parse_iso(ts).astimezone(NZ_TZ) if ts else None


def order_verified_nz(o):
    ts = o.get("timestamps", {}).get("verified", {}).get("time")
    return parse_iso(ts).astimezone(NZ_TZ) if ts else None


def order_arrival_date(o):
    ds = o.get("details", {}).get("arrivalDate")
    return datetime.date.fromisoformat(ds) if ds else None


def order_required_date(o):
    """Required *delivery* date — when the customer needs goods to arrive."""
    ds = o.get("details", {}).get("deliver", {}).get("requiredDate")
    return datetime.date.fromisoformat(ds) if ds else None


def order_ship_date(o):
    """Required *ship* date — when the order should leave the warehouse, as
    set (deliberately or not) in CartonCloud's details.collect.requiredDate."""
    ds = o.get("details", {}).get("collect", {}).get("requiredDate")
    return datetime.date.fromisoformat(ds) if ds else None


def fmt_dt(dt):
    """Format a datetime for display, or return ''."""
    if not dt:
        return ""
    return dt.strftime("%d %b %Y %H:%M")


def fmt_date(d):
    if not d:
        return ""
    return d.strftime("%d %b %Y")


# ── KPI deadline logic (shared with existing scripts) ─────────────────────────
def is_business_day(d):
    return d.weekday() < 5 and d not in PUBLIC_HOLIDAYS


def next_business_day(d):
    nd = d + datetime.timedelta(days=1)
    while not is_business_day(nd):
        nd += datetime.timedelta(days=1)
    return nd


def business_days_before(d, n):
    nd = d
    remaining = n
    while remaining > 0:
        nd -= datetime.timedelta(days=1)
        if is_business_day(nd):
            remaining -= 1
    return nd


def next_weekday_on_or_after(d, target_weekday):
    delta = (target_weekday - d.weekday()) % 7
    if delta == 0:
        delta = 7  # "next Monday" means the following Monday, not today, matching mtd_kpis.py
    return d + datetime.timedelta(days=delta)


def parse_time_of_day(text):
    text = text.strip().lower()
    m = re.match(r"^(\d{1,2})(?::(\d{2}))?\s*(am|pm)?$", text)
    if not m:
        return 23, 59
    hour = int(m.group(1))
    minute = int(m.group(2)) if m.group(2) else 0
    ampm = m.group(3)
    if ampm == "pm" and hour < 12:
        hour += 12
    elif ampm == "am" and hour == 12:
        hour = 0
    return hour, minute


def parse_deadline_rule(rule_str):
    rule_str = (rule_str or "Same day").strip()
    low = rule_str.lower()
    if low == "same day":
        return {"type": "same_day"}
    if low.startswith("next business day"):
        return {"type": "next_business_day_eod"}
    m = re.match(r"^fixed\s*\((.+)\)$", low)
    if m:
        try:
            when = datetime.datetime.strptime(m.group(1).strip(), "%Y-%m-%d %H:%M")
            return {"type": "fixed", "when": when}
        except ValueError:
            pass
    m = re.match(r"^next (\w+)\s*\((.+)\)$", low)
    if m:
        weekday_name, time_part = m.groups()
        weekday = WEEKDAY_NAMES.get(weekday_name)
        if weekday is not None:
            hour, minute = parse_time_of_day(time_part)
            return {"type": "next_weekday", "weekday": weekday, "hour": hour, "minute": minute}
    return {"type": "same_day"}


def deadline_for_rule(created_dt, rule):
    if rule["type"] == "fixed":
        # Absolute one-off deadline (e.g. "Fixed (2026-09-02 20:00)") - used for
        # customer-wide operational holds (stock embargo/customs hold) where every
        # order in the effective window should get the same real-world deadline,
        # regardless of which day within the window it happened to be created.
        return rule["when"].replace(tzinfo=NZ_TZ)
    if rule["type"] == "next_weekday":
        target_date = next_weekday_on_or_after(created_dt.date(), rule["weekday"])
        return datetime.datetime.combine(
            target_date, datetime.time(rule["hour"], rule["minute"]), tzinfo=NZ_TZ
        )
    if rule["type"] == "next_business_day_eod":
        nbd = next_business_day(created_dt.date())
        return datetime.datetime.combine(nbd, datetime.time(23, 59, 59), tzinfo=NZ_TZ)
    return datetime.datetime.combine(
        created_dt.date(), datetime.time(23, 59, 59), tzinfo=NZ_TZ
    )


def order_cut_off_deadline(created_dt, cutoff_hour=CUTOFF_HOUR, lenient_next_day=False,
                            after_cutoff_extra_days=0):
    """Matches mtd_kpis.py / extract_kpi_data.py exactly: the deadline TIME is
    always COB_HOUR (8pm), never cutoff_hour - cutoff_hour only controls
    whether an order counts as "before" or "after" cutoff, not what time its
    dispatch deadline falls at. after_cutoff_extra_days adds business days on
    top of the standard 1-day grace period for orders received after cutoff
    (e.g. Logwin - Sikla's override grants +1 extra day)."""
    is_weekend = not is_business_day(created_dt.date())
    if (not is_weekend) and created_dt.hour < cutoff_hour:
        if lenient_next_day:
            nbd = next_business_day(created_dt.date())
            return datetime.datetime.combine(nbd, datetime.time(COB_HOUR, 0, 0), tzinfo=NZ_TZ)
        return datetime.datetime.combine(
            created_dt.date(), datetime.time(COB_HOUR, 0, 0), tzinfo=NZ_TZ
        )
    nbd = created_dt.date()
    for _ in range(1 + after_cutoff_extra_days):
        nbd = next_business_day(nbd)
    return datetime.datetime.combine(nbd, datetime.time(COB_HOUR, 0, 0), tzinfo=NZ_TZ)


def load_exceptions(path=WORKBOOK):
    """Load customer cutoff overrides and content-based overrides from the
    Exceptions tab of kpi_dashboard.xlsx (same logic as existing scripts).

    Also parses Table 6: Public Holidays - ad-hoc dates added there (e.g. a
    one-off warehouse outage, or a regional anniversary day not in the
    PUBLIC_HOLIDAYS constant above) are merged into PUBLIC_HOLIDAYS by the
    caller (see main()) so every deadline calculation in this file picks
    them up automatically, matching mtd_kpis.py / extract_kpi_data.py."""
    result = {"cutoff_overrides": {}, "order_exceptions": [], "content_overrides": {}, "public_holidays": set()}
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except FileNotFoundError:
        return result
    if "Exceptions" not in wb.sheetnames:
        return result
    ws = wb["Exceptions"]

    section1_row = section2_row = section3_row = section6_row = section7_row = None
    for row in range(1, ws.max_row + 1):
        val = ws.cell(row=row, column=1).value
        if val == "Customer Cutoff Time Overrides":
            section1_row = row + 1
        elif val == "Order-Level Exceptions":
            section2_row = row + 1
        elif val == "Order Content-Based Overrides":
            section3_row = row + 1
        elif val == "Table 6: Public Holidays":
            section6_row = row + 1
        elif val == "Table 7: Operational Exceptions (Non-Business Days)":
            section7_row = row + 1

    if section2_row:
        r = section2_row + 1
        action_map = {
            "exclude from calculation": "exclude",
            "count as met": "count_as_met",
        }
        while r <= ws.max_row:
            customer   = ws.cell(row=r, column=2).value
            reference  = ws.cell(row=r, column=3).value
            month      = ws.cell(row=r, column=4).value
            kpi        = ws.cell(row=r, column=5).value
            action_raw = ws.cell(row=r, column=6).value
            if not customer and not reference:
                # Skip blank rows but don't stop — merged cells can leave gaps
                r += 1
                continue
            if customer and reference and kpi and action_raw:
                action = action_map.get(str(action_raw).strip().lower())
                if action:
                    date_from = ws.cell(row=r, column=8).value
                    date_to = ws.cell(row=r, column=9).value
                    if isinstance(date_from, datetime.datetime):
                        date_from = date_from.date()
                    if isinstance(date_to, datetime.datetime):
                        date_to = date_to.date()
                    result["order_exceptions"].append({
                        "customer": customer,
                        "reference": str(reference).strip(),
                        "month": str(month).strip() if month else None,
                        "kpi": kpi,
                        "action": action,
                        "date_from": date_from,
                        "date_to": date_to,
                    })
            r += 1

    if section1_row:
        r = section1_row + 1
        while r <= ws.max_row:
            customer = ws.cell(row=r, column=1).value
            if not customer:
                break
            cutoff_str  = ws.cell(row=r, column=2).value
            rule_str    = ws.cell(row=r, column=3).value or "Same day"
            eff_from    = ws.cell(row=r, column=4).value
            eff_to      = ws.cell(row=r, column=5).value
            ch, cm = CUTOFF_HOUR, 0
            if cutoff_str:
                parts = str(cutoff_str).strip().split(":")
                ch = int(parts[0])
                cm = int(parts[1]) if len(parts) > 1 else 0
            lenient = str(rule_str).strip().lower().startswith("next business day")
            if isinstance(eff_from, datetime.datetime):
                eff_from = eff_from.date()
            if isinstance(eff_to, datetime.datetime):
                eff_to = eff_to.date()
            extra_days_raw = ws.cell(row=r, column=7).value
            try:
                after_cutoff_extra_days = int(extra_days_raw)
            except (TypeError, ValueError):
                after_cutoff_extra_days = 0
            result["cutoff_overrides"].setdefault(customer, []).append({
                "cutoff_hour": ch, "cutoff_minute": cm,
                "lenient_next_day": lenient,
                "effective_from": eff_from, "effective_to": eff_to,
                "after_cutoff_extra_days": after_cutoff_extra_days,
            })
            r += 1

    if section3_row:
        r = section3_row + 1
        while r <= ws.max_row:
            customer = ws.cell(row=r, column=1).value
            if not customer:
                break
            match_field     = ws.cell(row=r, column=2).value
            contains        = ws.cell(row=r, column=3).value
            rule_str        = ws.cell(row=r, column=4).value
            eff_from        = ws.cell(row=r, column=5).value
            eff_to          = ws.cell(row=r, column=6).value
            if isinstance(eff_from, datetime.datetime):
                eff_from = eff_from.date()
            if isinstance(eff_to, datetime.datetime):
                eff_to = eff_to.date()
            result["content_overrides"].setdefault(customer, []).append({
                "match_field": match_field,
                "contains": str(contains).strip().lower() if contains else "",
                "rule": parse_deadline_rule(rule_str),
                "effective_from": eff_from, "effective_to": eff_to,
            })
            r += 1

    if section6_row:
        r = section6_row + 1  # skip column-header row
        while r <= ws.max_row:
            val = ws.cell(row=r, column=1).value
            if val is None:
                break
            if isinstance(val, datetime.datetime):
                result["public_holidays"].add(val.date())
            elif isinstance(val, datetime.date):
                result["public_holidays"].add(val)
            r += 1

    # Operational Exceptions (Non-Business Days) - not public holidays, but
    # merged into the same date set so deadline calculations treat them
    # identically (e.g. a one-off warehouse outage).
    if section7_row:
        r = section7_row + 1  # skip column-header row
        while r <= ws.max_row:
            val = ws.cell(row=r, column=1).value
            if val is None:
                break
            if isinstance(val, datetime.datetime):
                result["public_holidays"].add(val.date())
            elif isinstance(val, datetime.date):
                result["public_holidays"].add(val)
            r += 1

    return result


def get_cutoff_override(overrides, customer, order_date):
    for ov in overrides.get(customer, []):
        if ov.get("effective_from") and order_date < ov["effective_from"]:
            continue
        if ov.get("effective_to") and order_date > ov["effective_to"]:
            continue
        return ov
    return None


def find_order_exception(exceptions_list, customer, reference, order_date, month_str=None):
    """Find an applicable Order-Level Exception for this order.

    Reference "*" matches every order reference for that customer (a
    blanket exception). If Date From/Date To are set on the exception row,
    the order's actual creation date must fall within that range - this
    takes priority over (and ignores) the Month column, letting an
    exception scope to a specific week rather than a whole month. If no
    date range is set, falls back to Month-based scoping (matches
    mtd_kpis.py / extract_kpi_data.py).
    """
    for exc in exceptions_list:
        if exc["customer"] != customer:
            continue
        if exc["reference"] != reference and exc["reference"] != "*":
            continue
        date_from = exc.get("date_from")
        date_to = exc.get("date_to")
        if date_from or date_to:
            if date_from and order_date < date_from:
                continue
            if date_to and order_date > date_to:
                continue
            return exc
        if exc["month"] is None or exc["month"] == month_str:
            return exc
    return None


def required_transit_days(o, rural_pcs, rural_streets):
    """Business days needed between ship date and delivery date, based on the
    order's delivery destination. 1 for North Island, 3 for South Island
    (postcode >= 7000), +1 more if rural. Matches Order Cut Off / DIFOT
    elsewhere in this project (extract_kpi_data.py, mtd_kpis.py)."""
    deliver_addr = o.get("details", {}).get("deliver", {}).get("address", {})
    pc   = str(deliver_addr.get("postcode") or "").strip()
    st   = str(deliver_addr.get("address1") or "").strip()
    city = str(deliver_addr.get("city") or "").strip()
    si   = is_si_co(pc)
    rur  = is_rural_co(pc, st, city, rural_pcs or set(), rural_streets or set())
    return (3 if si else 1) + (1 if rur else 0)


def required_date_ship_deadline(o, rural_pcs, rural_streets):
    """Work out the ship-by date implied by required ship/delivery dates.

    Priority:
      - Required ship date (details.collect.requiredDate) is trusted as-is,
        UNLESS it's unfeasible — i.e. later than the latest date that still
        leaves enough transit time to hit the required delivery date, in
        which case the transit-time-derived date is used instead (it's
        earlier/safer than a ship date that can't make it in time).
      - If there's no ship date but there is a required delivery date, the
        ship deadline is derived the same way: delivery date minus the
        transit days needed for that destination.
      - If neither is set, there is no required-date-driven deadline and the
        standard order cut-off applies.

    Returns (ship_by_date | None, source: 'ship'|'delivery'|'ship_capped'|None).
    """
    ship_date    = order_ship_date(o)
    deliver_date = order_required_date(o)

    derived_date = None
    if deliver_date:
        days = required_transit_days(o, rural_pcs, rural_streets)
        derived_date = business_days_before(deliver_date, days)

    if ship_date and derived_date:
        if ship_date <= derived_date:
            return ship_date, "ship"
        # Ship date given is unfeasible (too close to/after the delivery
        # date for the known transit time) — fall back to the safer,
        # transit-derived date instead of trusting the entered ship date.
        return derived_date, "ship_capped"
    if ship_date:
        return ship_date, "ship"
    if derived_date:
        return derived_date, "delivery"
    return None, None


def compute_pick_deadline(o, exceptions, rural_pcs=None, rural_streets=None, stock_transitions=None):
    """Return (deadline: datetime, deadline_label: str) for an outbound order."""
    created_dt = order_kpi_clock_start(o, stock_transitions)
    if not created_dt:
        return None, ""

    # Overrides below are scoped (Effective From/To) against the order's
    # ORIGINAL creation date, not the stock-transition-adjusted created_dt.
    # Otherwise a stock-delayed order from an earlier/unrelated period can
    # accidentally drift into a later override's effective window purely
    # because its clock-start got pushed forward.
    raw_dt = order_created_nz(o)
    raw_created_date = raw_dt.date() if raw_dt else created_dt.date()

    cust       = ORDER_CUSTOMER_OVERRIDES.get(order_ref(o), order_customer(o))
    overrides  = exceptions.get("cutoff_overrides", {})
    content_ov = exceptions.get("content_overrides", {})

    deadline = order_cut_off_deadline(created_dt)

    ov = get_cutoff_override(overrides, cust, raw_created_date)
    if ov:
        alt = order_cut_off_deadline(
            created_dt,
            cutoff_hour=ov["cutoff_hour"],
            lenient_next_day=ov["lenient_next_day"],
            after_cutoff_extra_days=ov.get("after_cutoff_extra_days", 0),
        )
        deadline = max(deadline, alt)

    # Required ship / delivery date handling — see required_date_ship_deadline().
    # Only applied if more lenient than the standard deadline so far.
    ship_by_date, _source = required_date_ship_deadline(o, rural_pcs, rural_streets)
    if ship_by_date:
        alt = datetime.datetime.combine(
            ship_by_date, datetime.time(REQUIRED_DATE_HOUR, 0), tzinfo=NZ_TZ
        )
        deadline = max(deadline, alt)

    for cov in content_ov.get(cust, []):
        if cov.get("effective_from") and raw_created_date < cov["effective_from"]:
            continue
        if cov.get("effective_to") and raw_created_date > cov["effective_to"]:
            continue
        getter = CONTENT_MATCH_FIELDS.get(cov.get("match_field", ""))
        if not getter:
            continue
        if cov["contains"] and cov["contains"] in (getter(o) or "").lower():
            alt = deadline_for_rule(created_dt, cov["rule"])
            deadline = max(deadline, alt)

    return deadline, fmt_dt(deadline)


# ── Data fetching ─────────────────────────────────────────────────────────────
def fetch_outbound_pick_queue(token, tenant_id):
    """All outbound orders that are queued or actively being picked/packed.
    CartonCloud statuses: AWAITING_PICK_AND_PACK → PICKED → PACKING_IN_PROGRESS → DISPATCHED."""
    results = []
    for status in ("AWAITING_PICK_AND_PACK", "PICKED", "PACKING_IN_PROGRESS"):
        print(f"  Fetching pick queue ({status})…")
        results.extend(search_all_pages(
            token, tenant_id, "outbound-orders",
            status_condition(status),
        ))
    return results


def fetch_outbound_completed_today(token, tenant_id, today_iso, tomorrow_iso):
    """Outbound orders dispatched today."""
    print("  Fetching orders dispatched today…")
    return search_all_pages(
        token, tenant_id, "outbound-orders",
        date_range_condition("/timestamps/dispatched/time", today_iso, tomorrow_iso),
    )


# fetch_volume_for_range / load_volume_cache / save_volume_cache / refresh_volume
# used to be defined here, but now live in volume_data.py, shared with
# mtd_kpis.py and extract_kpi_data.py (Order Accuracy by Unit's denominator
# needs the same per-customer daily volume, and reusing this cache means
# no extra CartonCloud API calls beyond what this dashboard already made).
# See volume_data.py's module docstring for details.


def fetch_inbound_open(token, tenant_id, start_iso, end_iso):
    """Fetch all open inbound POs by status.

    CartonCloud inbound status meanings (per warehouse):
      NOT_YET_RECEIVED  – hasn't physically arrived at the warehouse
      RECEIVED          – goods on site, actively being received/put away
      VERIFIED          – received, put-away in progress / verification stage
      ALLOCATED         – fully received and processed (complete – exclude)
      DRAFT / REJECTED  – exclude

    Fetch by status rather than date so we don't miss overdue POs with old
    arrival dates.
    """
    results = []
    for status in ("NOT_YET_RECEIVED", "RECEIVED", "VERIFIED"):
        print(f"  Fetching inbound orders ({status})…")
        results.extend(search_all_pages(
            token, tenant_id, "inbound-orders",
            status_condition(status),
        ))
    return results


# ── Classify orders into dashboard buckets ────────────────────────────────────
def classify_pick_queue(orders, exceptions, now, rural_pcs=None, rural_streets=None, stock_transitions=None):
    """Split pick queue into due_today, on_time (future deadline) and late.

    due_today  – deadline falls today; must dispatch today to meet KPI
    on_time    – deadline is tomorrow or later (pre-loaded / future orders)
    late       – deadline already passed
    """
    due_today, on_time, late = [], [], []
    today = now.date()
    cut_off_exceptions = [
        e for e in exceptions.get("order_exceptions", []) if e["kpi"] == "Order Cut Off"
    ]
    for o in orders:
        cust = order_customer(o)
        if cust in TEST_ACCOUNT_NAMES:
            continue
        deadline, deadline_label = compute_pick_deadline(o, exceptions, rural_pcs, rural_streets, stock_transitions)
        created_dt = order_created_nz(o)
        clock_start = order_kpi_clock_start(o, stock_transitions)
        held = bool(clock_start and created_dt and clock_start != created_dt)
        row = {
            "ref":      order_ref(o),
            "customer": cust,
            "status":   order_status(o),
            "created":  fmt_dt(created_dt),
            "deadline": deadline_label,
            "ship_date": fmt_date(order_ship_date(o)),
            "req_date": fmt_date(order_required_date(o)),
            "lines":    order_line_count(o),
            "units":    order_unit_count(o),
            "held":     held,
            "available_from": fmt_dt(clock_start) if held else "",
        }

        # Order-level Cut Off exception (e.g. blanket "count as met" for a
        # customer over a specific date range) — never show these as late,
        # regardless of the computed deadline.
        exc = None
        if created_dt:
            exc = find_order_exception(cut_off_exceptions, cust, order_ref(o), created_dt.date())
        if exc and exc["action"] == "exclude":
            continue
        waived = bool(exc and exc["action"] == "count_as_met")

        if deadline and now > deadline and not waived:
            row["hours_late"] = round((now - deadline).total_seconds() / 3600, 1)
            late.append(row)
        elif deadline and deadline.date() == today:
            due_today.append(row)
        else:
            on_time.append(row)
    return due_today, on_time, late


def annotate_completed(orders, now, today):
    rows = []
    for o in orders:
        if order_customer(o) in TEST_ACCOUNT_NAMES:
            continue
        dispatched = order_dispatched_nz(o)
        if not dispatched or dispatched.date() != today:
            continue
        rows.append({
            "ref":       order_ref(o),
            "customer":  order_customer(o),
            "status":    order_status(o),
            "created":   fmt_dt(order_created_nz(o)),
            "dispatched":fmt_dt(dispatched),
            "lines":     order_line_count(o),
            "units":     order_unit_count(o),
        })
    return rows


def classify_inbound(orders, now, today):
    """
    Classify open inbound POs into three operational buckets.

    Real CartonCloud inbound statuses observed:
      DRAFT            – not yet confirmed, ignore
      ALLOCATED        – PO confirmed, awaiting arrival
      NOT_YET_RECEIVED – explicitly marked not received
      RECEIVED         – goods physically received, being put away
      REJECTED         – ignore
      (COMPLETED / fully verified orders have timestamps.verified set)

    expected_not_arrived  – open, arrivalDate > today
    in_receipt            – open, arrivalDate <= today, within 48-hr KPI window
    failing_kpi           – open, arrivalDate <= today, past 48-hr KPI window
    """
    expected, in_receipt, failing = [], [], []

    for o in orders:
        if order_customer(o) in TEST_ACCOUNT_NAMES:
            continue
        status = order_status(o)
        arrival_date = order_arrival_date(o)

        row = {
            "ref":          order_ref(o),
            "customer":     order_customer(o),
            "status":       status,
            "arrival_date": fmt_date(arrival_date),
            "lines":        order_line_count(o),
            "units":        order_unit_count(o),
        }

        if status == "NOT_YET_RECEIVED":
            # Hasn't physically arrived at the warehouse yet
            row["days_until"] = (
                (arrival_date - today).days if arrival_date else "?"
            )
            expected.append(row)

        elif status in ("RECEIVED", "VERIFIED"):
            # Goods on site — clock starts when the PO was last modified, which
            # in CartonCloud corresponds to when it was moved to RECEIVED status.
            # (timestamps.created.time = when the PO was first entered, which may
            # be weeks before arrival, so is NOT the right clock start.)
            received_dt = order_modified_nz(o)
            if received_dt:
                hours_since = (now - received_dt).total_seconds() / 3600
            else:
                hours_since = 0

            if hours_since > INBOUND_KPI_HOURS:
                row["hours_late"] = round(hours_since - INBOUND_KPI_HOURS, 1)
                row["received_at"] = fmt_dt(received_dt)
                failing.append(row)
            else:
                row["hours_in"]    = round(max(hours_since, 0), 1)
                row["received_at"] = fmt_dt(received_dt)
                in_receipt.append(row)

    return expected, in_receipt, failing


def build_daily_volume(orders, today):
    """Kept for reference — replaced by refresh_volume() with caching."""
    by_day = {}
    for o in orders:
        if order_customer(o) in TEST_ACCOUNT_NAMES:
            continue
        dispatched = order_dispatched_nz(o)
        if not dispatched:
            continue
        day = dispatched.date().isoformat()
        if day not in by_day:
            by_day[day] = {"orders": 0, "lines": 0, "units": 0}
        by_day[day]["orders"] += 1
        by_day[day]["lines"]  += order_line_count(o)
        by_day[day]["units"]  += order_unit_count(o)
    # Return sorted list of {date, orders, lines, units}
    return [{"date": d, **v} for d, v in sorted(by_day.items())]


def summarise(rows):
    """Return (count, total_lines, total_units) for a list of order rows."""
    return (
        len(rows),
        sum(r.get("lines", 0) for r in rows),
        sum(r.get("units", 0) for r in rows),
    )


# ── HTML template ─────────────────────────────────────────────────────────────
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Warehouse Operations – Fulfilment Plus</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
:root {
  --bg:         #f0f2f5;
  --card:       #ffffff;
  --text:       #1a202c;
  --muted:      #64748b;
  --border:     #e2e8f0;
  --accent:     #1d4ed8;
  --warn:       #d97706;
  --warn-bg:    #fef3c7;
  --danger:     #dc2626;
  --danger-bg:  #fee2e2;
  --good:       #16a34a;
  --good-bg:    #dcfce7;
  --info:       #0369a1;
  --info-bg:    #e0f2fe;
  --neutral-bg: #f1f5f9;
  --radius:     10px;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;
     background:var(--bg);color:var(--text);min-height:100vh}

/* ── Header ── */
header{
  background:#fff;border-bottom:1px solid var(--border);
  padding:16px 28px;display:flex;align-items:center;
  justify-content:space-between;flex-wrap:wrap;gap:12px;
  position:sticky;top:0;z-index:100;
}
header h1{font-size:1.25rem;font-weight:700}
.header-meta{font-size:.8rem;color:var(--muted)}
.header-meta strong{color:var(--text)}
.refresh-hint{font-size:.75rem;color:var(--muted);margin-top:2px}

/* ── Layout ── */
main{padding:24px 28px 60px;max-width:1280px;margin:0 auto}
.section{margin-bottom:36px}
.section-title{
  font-size:1rem;font-weight:700;margin-bottom:14px;
  display:flex;align-items:center;gap:8px;
}
.section-title span.dot{
  display:inline-block;width:10px;height:10px;border-radius:50%
}

/* ── KPI Grid ── */
.kpi-grid{
  display:grid;
  grid-template-columns:repeat(auto-fit,minmax(280px,1fr));
  gap:16px;
}

/* ── KPI Card ── */
.kpi-card{
  background:var(--card);border:1px solid var(--border);
  border-radius:var(--radius);padding:18px 20px;
  cursor:pointer;transition:box-shadow .15s,transform .1s;
  position:relative;overflow:hidden;
}
.kpi-card:hover{box-shadow:0 4px 16px rgba(0,0,0,.1);transform:translateY(-1px)}
.kpi-card:active{transform:translateY(0)}
.kpi-card .card-title{
  font-size:.85rem;font-weight:600;color:var(--muted);
  text-transform:uppercase;letter-spacing:.04em;margin-bottom:12px;
}
.kpi-card.danger{border-top:4px solid var(--danger)}
.kpi-card.warn  {border-top:4px solid var(--warn)}
.kpi-card.good  {border-top:4px solid var(--good)}
.kpi-card.info  {border-top:4px solid var(--info)}
.kpi-card.neutral{border-top:4px solid var(--border)}

.metrics-row{display:flex;gap:0;margin-bottom:10px}
.metric{flex:1;text-align:center;padding:0 8px}
.metric:not(:last-child){border-right:1px solid var(--border)}
.metric .num{font-size:2rem;font-weight:800;line-height:1;color:var(--text)}
.metric .lbl{font-size:.7rem;color:var(--muted);margin-top:3px;text-transform:uppercase;letter-spacing:.03em}

.kpi-card.danger .metric .num{color:var(--danger)}
.kpi-card.warn   .metric .num{color:var(--warn)}
.kpi-card.good   .metric .num{color:var(--good)}
.kpi-card.info   .metric .num{color:var(--info)}

.card-footer{
  font-size:.75rem;color:var(--muted);margin-top:8px;
  display:flex;align-items:center;gap:4px;
}
.card-footer svg{width:12px;height:12px;flex-shrink:0}
.click-hint{
  position:absolute;bottom:10px;right:12px;
  font-size:.65rem;color:var(--border);pointer-events:none;
}

/* ── Chart section ── */
.chart-card{
  background:var(--card);border:1px solid var(--border);
  border-radius:var(--radius);padding:20px;
}
.chart-card h3{font-size:.95rem;font-weight:600;margin-bottom:4px}
.chart-card .chart-sub{font-size:.8rem;color:var(--muted);margin-bottom:16px}
.chart-controls{display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap}
.chart-controls button{
  padding:5px 14px;border-radius:6px;border:1px solid var(--border);
  background:#fff;color:var(--muted);font-size:.8rem;cursor:pointer;
}
.chart-controls button.active{
  background:var(--accent);color:#fff;border-color:var(--accent);font-weight:600;
}
.chart-wrap{position:relative;height:280px}

/* ── Modal ── */
.modal-backdrop{
  display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);
  z-index:1000;align-items:flex-start;justify-content:center;
  padding:40px 20px;overflow-y:auto;
}
.modal-backdrop.open{display:flex}
.modal{
  background:#fff;border-radius:12px;padding:24px;
  width:100%;max-width:960px;margin:auto;
  box-shadow:0 20px 60px rgba(0,0,0,.2);
  animation:slideIn .15s ease;
}
@keyframes slideIn{from{transform:translateY(-12px);opacity:0}to{transform:translateY(0);opacity:1}}
.modal-header{
  display:flex;align-items:flex-start;justify-content:space-between;
  margin-bottom:16px;gap:12px;
}
.modal-header h2{font-size:1.1rem;font-weight:700}
.modal-header .modal-sub{font-size:.8rem;color:var(--muted);margin-top:2px}
.modal-close{
  background:none;border:none;cursor:pointer;color:var(--muted);
  font-size:1.4rem;line-height:1;padding:0;flex-shrink:0;
}
.modal-close:hover{color:var(--text)}
.modal-stats{
  display:flex;gap:20px;margin-bottom:16px;
  padding:12px 16px;background:var(--neutral-bg);border-radius:8px;
  flex-wrap:wrap;
}
.modal-stat{text-align:center;min-width:80px}
.modal-stat .n{font-size:1.4rem;font-weight:800}
.modal-stat .l{font-size:.7rem;color:var(--muted);text-transform:uppercase}

/* ── Table ── */
.tbl-wrap{overflow-x:auto;max-height:60vh;overflow-y:auto}
table{width:100%;border-collapse:collapse;font-size:.82rem}
thead th{
  background:var(--neutral-bg);color:var(--muted);font-weight:600;
  padding:9px 12px;text-align:left;border-bottom:2px solid var(--border);
  position:sticky;top:0;
}
tbody td{padding:8px 12px;border-bottom:1px solid var(--border)}
tbody tr:last-child td{border-bottom:none}
tbody tr:hover{background:var(--neutral-bg)}
.badge{
  display:inline-block;font-size:.7rem;font-weight:600;
  padding:2px 8px;border-radius:999px;
}
.badge.danger{background:var(--danger-bg);color:var(--danger)}
.badge.warn  {background:var(--warn-bg);  color:var(--warn)}
.badge.good  {background:var(--good-bg);  color:var(--good)}
.badge.info  {background:var(--info-bg);  color:var(--info)}
.badge.neutral{background:var(--neutral-bg);color:var(--muted)}
.empty{color:var(--muted);font-size:.85rem;padding:20px 0;text-align:center}
</style>
</head>
<body>

<header>
  <div>
    <a href="index.html" style="font-size:0.85rem;color:var(--muted);text-decoration:none;">&larr; All dashboards</a>
    <h1>Warehouse Operations</h1>
    <div class="header-meta">Fulfilment Plus &middot; <strong id="gen-time">__GENERATED__</strong></div>
    <div class="refresh-hint">Re-run <code>generate_warehouse_dashboard.py</code> to refresh data</div>
  </div>
  <div style="text-align:right">
    <div class="header-meta" id="today-label"></div>
  </div>
</header>

<main>

  <!-- ── Sales Orders ──────────────────────────────────── -->
  <div class="section">
    <div class="section-title">
      <span class="dot" style="background:var(--accent)"></span>
      Sales Orders
    </div>
    <div class="kpi-grid">

      <div class="kpi-card warn" onclick="openModal('due_today')">
        <div class="card-title">Due Today – Must Pick to Meet Cut-off</div>
        <div class="metrics-row">
          <div class="metric"><div class="num" id="dt-orders">–</div><div class="lbl">Orders</div></div>
          <div class="metric"><div class="num" id="dt-lines">–</div> <div class="lbl">Lines</div></div>
          <div class="metric"><div class="num" id="dt-units">–</div> <div class="lbl">Units</div></div>
        </div>
        <div class="card-footer">KPI deadline falls today — dispatch today to meet cut-off</div>
        <div class="click-hint">click to view ›</div>
      </div>

      <div class="kpi-card neutral" onclick="openModal('pick_queue')">
        <div class="card-title">Total Pick Queue</div>
        <div class="metrics-row">
          <div class="metric"><div class="num" id="pq-orders">–</div><div class="lbl">Orders</div></div>
          <div class="metric"><div class="num" id="pq-lines">–</div> <div class="lbl">Lines</div></div>
          <div class="metric"><div class="num" id="pq-units">–</div> <div class="lbl">Units</div></div>
        </div>
        <div class="card-footer">All orders currently in pick &amp; pack queue</div>
        <div class="click-hint">click to view ›</div>
      </div>

      <div class="kpi-card danger" onclick="openModal('late_pick')">
        <div class="card-title">Outstanding – Failing KPI</div>
        <div class="metrics-row">
          <div class="metric"><div class="num" id="lp-orders">–</div><div class="lbl">Orders</div></div>
          <div class="metric"><div class="num" id="lp-lines">–</div> <div class="lbl">Lines</div></div>
          <div class="metric"><div class="num" id="lp-units">–</div> <div class="lbl">Units</div></div>
        </div>
        <div class="card-footer">Past KPI dispatch deadline and not yet picked</div>
        <div class="click-hint">click to view ›</div>
      </div>

      <div class="kpi-card good" onclick="openModal('completed')">
        <div class="card-title">Completed Today</div>
        <div class="metrics-row">
          <div class="metric"><div class="num" id="ct-orders">–</div><div class="lbl">Orders</div></div>
          <div class="metric"><div class="num" id="ct-lines">–</div> <div class="lbl">Lines</div></div>
          <div class="metric"><div class="num" id="ct-units">–</div> <div class="lbl">Units</div></div>
        </div>
        <div class="card-footer">Dispatched today (NZ time)</div>
        <div class="click-hint">click to view ›</div>
      </div>

    </div>
  </div>

  <!-- ── Purchase Orders ───────────────────────────────── -->
  <div class="section">
    <div class="section-title">
      <span class="dot" style="background:var(--warn)"></span>
      Purchase Orders
    </div>
    <div class="kpi-grid">

      <div class="kpi-card info" onclick="openModal('expected')">
        <div class="card-title">Expected – Not Yet Arrived</div>
        <div class="metrics-row">
          <div class="metric"><div class="num" id="ex-orders">–</div><div class="lbl">POs</div></div>
          <div class="metric"><div class="num" id="ex-lines">–</div> <div class="lbl">Lines</div></div>
          <div class="metric"><div class="num" id="ex-units">–</div> <div class="lbl">Units</div></div>
        </div>
        <div class="card-footer">Future arrival date, not yet verified</div>
        <div class="click-hint">click to view ›</div>
      </div>

      <div class="kpi-card warn" onclick="openModal('in_receipt')">
        <div class="card-title">In Receipt</div>
        <div class="metrics-row">
          <div class="metric"><div class="num" id="ir-orders">–</div><div class="lbl">POs</div></div>
          <div class="metric"><div class="num" id="ir-lines">–</div> <div class="lbl">Lines</div></div>
          <div class="metric"><div class="num" id="ir-units">–</div> <div class="lbl">Units</div></div>
        </div>
        <div class="card-footer">RECEIVED or VERIFIED, within 48hr dock-to-stock window</div>
        <div class="click-hint">click to view ›</div>
      </div>

      <div class="kpi-card danger" onclick="openModal('failing_po')">
        <div class="card-title">Open – Failing KPI</div>
        <div class="metrics-row">
          <div class="metric"><div class="num" id="fp-orders">–</div><div class="lbl">POs</div></div>
          <div class="metric"><div class="num" id="fp-lines">–</div> <div class="lbl">Lines</div></div>
          <div class="metric"><div class="num" id="fp-units">–</div> <div class="lbl">Units</div></div>
        </div>
        <div class="card-footer">RECEIVED/VERIFIED &gt;48 hrs ago, not yet put away</div>
        <div class="click-hint">click to view ›</div>
      </div>

    </div>
  </div>

  <!-- ── 12-Month Volume Chart ──────────────────────────── -->
  <div class="section">
    <div class="section-title">
      <span class="dot" style="background:var(--good)"></span>
      12-Month Daily Volumes
    </div>
    <div class="chart-card">
      <h3>Outbound volumes – daily (last 12 months)</h3>
      <div class="chart-sub">Sales orders dispatched per day</div>
      <div class="chart-controls">
        <button class="active" onclick="setMetric('orders',this)">Orders</button>
        <button onclick="setMetric('lines',this)">Lines</button>
        <button onclick="setMetric('units',this)">Units</button>
      </div>
      <div class="chart-wrap"><canvas id="volChart"></canvas></div>
    </div>
  </div>

</main>

<!-- ── Modal ── -->
<div class="modal-backdrop" id="modal-backdrop" onclick="closeModalOnBackdrop(event)">
  <div class="modal" id="modal">
    <div class="modal-header">
      <div>
        <h2 id="modal-title">–</h2>
        <div class="modal-sub" id="modal-sub"></div>
      </div>
      <button class="modal-close" onclick="closeModal()">×</button>
    </div>
    <div class="modal-stats" id="modal-stats"></div>
    <div class="tbl-wrap">
      <table id="modal-table">
        <thead id="modal-thead"></thead>
        <tbody id="modal-tbody"></tbody>
      </table>
    </div>
  </div>
</div>

<script>
// ── Data injected by Python ────────────────────────────────────────────────
const D = __DATA_JSON__;

// ── Initialise page ────────────────────────────────────────────────────────
document.getElementById('today-label').textContent =
  new Date().toLocaleDateString('en-NZ', {weekday:'long',day:'numeric',month:'long',year:'numeric'});

function fmt(n){ return n === 0 ? '0' : n.toLocaleString(); }

function setCard(orderId, lineId, unitId, rows){
  const [orders, lines, units] = [
    rows.length,
    rows.reduce((s,r)=>s+(r.lines||0),0),
    rows.reduce((s,r)=>s+(r.units||0),0),
  ];
  document.getElementById(orderId).textContent = fmt(orders);
  document.getElementById(lineId ).textContent = fmt(lines);
  document.getElementById(unitId ).textContent = fmt(units);
}

setCard('dt-orders','dt-lines','dt-units', D.due_today);
setCard('pq-orders','pq-lines','pq-units', D.pick_queue);
setCard('lp-orders','lp-lines','lp-units', D.late_pick);
setCard('ct-orders','ct-lines','ct-units', D.completed);
setCard('ex-orders','ex-lines','ex-units', D.expected);
setCard('ir-orders','ir-lines','ir-units', D.in_receipt);
setCard('fp-orders','fp-lines','fp-units', D.failing_po);

// ── Modal ──────────────────────────────────────────────────────────────────
const MODAL_CONFIG = {
  due_today: {
    title: 'Due Today – Must Pick to Meet Cut-off',
    sub:   'KPI deadline falls today — must be dispatched today',
    cols:  ['ref','customer','created','available_from','deadline','ship_date','req_date','lines','units'],
    hdrs:  ['Order #','Customer','Created','Held Until','KPI Deadline','Req. Ship Date','Req. Delivery Date','Lines','Units'],
    badge: null,
  },
  pick_queue: {
    title: 'Total Pick Queue',
    sub:   'All orders currently in pick & pack queue (due today + future)',
    cols:  ['ref','customer','created','available_from','deadline','ship_date','req_date','lines','units'],
    hdrs:  ['Order #','Customer','Created','Held Until','KPI Deadline','Req. Ship Date','Req. Delivery Date','Lines','Units'],
    badge: null,
  },
  late_pick: {
    title: 'Outstanding – Failing KPI',
    sub:   'Past KPI dispatch deadline and not yet picked',
    cols:  ['ref','customer','created','available_from','deadline','ship_date','req_date','hours_late','lines','units'],
    hdrs:  ['Order #','Customer','Created','Held Until','KPI Deadline','Req. Ship Date','Req. Delivery Date','Hrs Late','Lines','Units'],
    badge: (r) => r.hours_late ? `<span class="badge danger">${r.hours_late}h late</span>` : '',
    badgeCol: 'hours_late',
  },
  completed: {
    title: 'Completed Today',
    sub:   'Dispatched today (NZ time)',
    cols:  ['ref','customer','created','dispatched','lines','units'],
    hdrs:  ['Order #','Customer','Created','Dispatched','Lines','Units'],
    badge: null,
  },
  expected: {
    title: 'Expected – Not Yet Arrived',
    sub:   'Future arrival date, not yet verified / put away',
    cols:  ['ref','customer','arrival_date','days_until','lines','units'],
    hdrs:  ['PO #','Customer','Expected Arrival','Days Away','Lines','Units'],
    badge: null,
  },
  in_receipt: {
    title: 'In Receipt',
    sub:   'Goods on site — within 48-hour dock-to-stock KPI window',
    cols:  ['ref','customer','arrival_date','received_at','hours_in','lines','units'],
    hdrs:  ['PO #','Customer','Arrival Date','Received At','Hrs In','Lines','Units'],
    badge: (r) => r.hours_in != null ? `<span class="badge warn">${r.hours_in}h</span>` : '',
    badgeCol: 'hours_in',
  },
  failing_po: {
    title: 'Open POs – Failing Dock-to-Stock KPI',
    sub:   'In RECEIVED or VERIFIED status for more than 48 hours',
    cols:  ['ref','customer','arrival_date','received_at','hours_late','lines','units'],
    hdrs:  ['PO #','Customer','Arrival Date','Received At','Hrs Past KPI','Lines','Units'],
    badge: (r) => r.hours_late ? `<span class="badge danger">${r.hours_late}h late</span>` : '',
    badgeCol: 'hours_late',
  },
};

const BUCKET_MAP = {
  due_today:  D.due_today,
  pick_queue: D.pick_queue,
  late_pick:  D.late_pick,
  completed:  D.completed,
  expected:   D.expected,
  in_receipt: D.in_receipt,
  failing_po: D.failing_po,
};

function openModal(key){
  const cfg  = MODAL_CONFIG[key];
  const rows = BUCKET_MAP[key] || [];

  document.getElementById('modal-title').textContent = cfg.title;
  document.getElementById('modal-sub').textContent   = cfg.sub;

  const orders = rows.length;
  const lines  = rows.reduce((s,r)=>s+(r.lines||0),0);
  const units  = rows.reduce((s,r)=>s+(r.units||0),0);
  document.getElementById('modal-stats').innerHTML = `
    <div class="modal-stat"><div class="n">${fmt(orders)}</div><div class="l">Orders / POs</div></div>
    <div class="modal-stat"><div class="n">${fmt(lines)}</div><div class="l">Lines</div></div>
    <div class="modal-stat"><div class="n">${fmt(units)}</div><div class="l">Units</div></div>
  `;

  const thead = document.getElementById('modal-thead');
  thead.innerHTML = '<tr>' + cfg.hdrs.map(h=>`<th>${h}</th>`).join('') + '</tr>';

  const tbody = document.getElementById('modal-tbody');
  if(rows.length === 0){
    tbody.innerHTML = `<tr><td colspan="${cfg.cols.length}" class="empty">No orders in this category right now.</td></tr>`;
  } else {
    tbody.innerHTML = rows.map(r=>{
      const cells = cfg.cols.map(col=>{
        let val = r[col];
        if(val == null || val === '') val = '–';
        // Render badge on the hours_late / hours_in cell
        if(cfg.badge && col === cfg.badgeCol){
          return `<td>${cfg.badge(r)}</td>`;
        }
        if(typeof val === 'number' && ['lines','units'].includes(col)){
          val = fmt(val);
        }
        return `<td>${val}</td>`;
      });
      return `<tr>${cells.join('')}</tr>`;
    }).join('');
  }

  document.getElementById('modal-backdrop').classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeModal(){
  document.getElementById('modal-backdrop').classList.remove('open');
  document.body.style.overflow = '';
}

function closeModalOnBackdrop(e){
  if(e.target === document.getElementById('modal-backdrop')) closeModal();
}

document.addEventListener('keydown', e => { if(e.key === 'Escape') closeModal(); });

// ── 12-month chart ─────────────────────────────────────────────────────────
const volData = D.daily_volume;
let currentMetric = 'orders';
let volChart = null;

const METRIC_LABELS = { orders: 'Orders', lines: 'Lines', units: 'Units' };
const METRIC_COLORS = {
  orders: { border: '#1d4ed8', bg: 'rgba(29,78,216,.12)' },
  lines:  { border: '#16a34a', bg: 'rgba(22,163,74,.12)'  },
  units:  { border: '#d97706', bg: 'rgba(217,119,6,.12)'  },
};

function buildChart(metric){
  const ctx = document.getElementById('volChart').getContext('2d');
  if(volChart) volChart.destroy();
  const col = METRIC_COLORS[metric];
  volChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels:   volData.map(d => d.date),
      datasets: [{
        label:           METRIC_LABELS[metric],
        data:            volData.map(d => d[metric]),
        backgroundColor: col.bg,
        borderColor:     col.border,
        borderWidth:     1,
        borderRadius:    2,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            title: items => {
              const d = new Date(items[0].label + 'T00:00:00');
              return d.toLocaleDateString('en-NZ',{weekday:'short',day:'numeric',month:'short',year:'numeric'});
            },
            label: item => ` ${METRIC_LABELS[metric]}: ${item.raw.toLocaleString()}`,
          }
        }
      },
      scales: {
        x: {
          ticks: {
            maxTicksLimit: 12,
            callback: (val, idx) => {
              const d = new Date(volData[idx]?.date + 'T00:00:00');
              return d.toLocaleDateString('en-NZ',{month:'short',year:'2-digit'});
            },
          },
          grid: { display: false },
        },
        y: {
          beginAtZero: true,
          ticks: { callback: v => v.toLocaleString() },
        },
      },
    },
  });
}

function setMetric(metric, btn){
  currentMetric = metric;
  document.querySelectorAll('.chart-controls button').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  buildChart(metric);
}

buildChart('orders');
</script>
</body>
</html>
"""


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    now   = datetime.datetime.now(NZ_TZ)
    today = now.date()

    today_iso    = today.isoformat()
    tomorrow_iso = (today + datetime.timedelta(days=1)).isoformat()
    # Inbound window: 14 days in the past (capture late POs) to 60 days in future
    inbound_start = (today - datetime.timedelta(days=14)).isoformat()
    inbound_end   = (today + datetime.timedelta(days=60)).isoformat()
    # volume_start kept for any future reference; actual fetch uses refresh_volume()

    creds     = load_credentials()
    token     = get_cc_token(creds)
    tenant_id = creds["CARTONCLOUD_TENANT_ID"]
    exceptions = load_exceptions()
    # Merge any ad-hoc dates from the Exceptions tab's Table 6 (e.g. a one-off
    # warehouse outage logged as an exception) into the standard NZ public
    # holiday calendar, so every is_business_day()/deadline calculation below
    # picks them up automatically - matches mtd_kpis.py / extract_kpi_data.py,
    # which read Table 6 as their only source of holidays.
    PUBLIC_HOLIDAYS.update(exceptions.get("public_holidays", set()))
    rural_pcs     = load_rural_postcode_set_co()
    rural_streets = load_rural_street_lookup_co()

    print("Fetching CartonCloud data…")

    pick_queue_raw  = fetch_outbound_pick_queue(token, tenant_id)
    completed_raw   = fetch_outbound_completed_today(token, tenant_id, today_iso, tomorrow_iso)
    inbound_raw     = fetch_inbound_open(token, tenant_id, inbound_start, inbound_end)
    daily_volume    = volume_data.refresh_volume(token, tenant_id, today)

    print("Processing…")

    stock_transitions = load_awaiting_stock_transitions()
    due_today, on_time, late = classify_pick_queue(
        pick_queue_raw, exceptions, now, rural_pcs, rural_streets, stock_transitions
    )
    completed     = annotate_completed(completed_raw, now, today)
    expected, in_receipt, failing_po = classify_inbound(inbound_raw, now, today)

    all_queue = due_today + on_time + late   # full pick queue
    pq_n, pq_l, pq_u = summarise(all_queue)
    dt_n, dt_l, dt_u = summarise(due_today)
    lp_n, lp_l, lp_u = summarise(late)
    ct_n, ct_l, ct_u = summarise(completed)
    ex_n, ex_l, ex_u = summarise(expected)
    ir_n, ir_l, ir_u = summarise(in_receipt)
    fp_n, fp_l, fp_u = summarise(failing_po)

    print(f"\nSales Orders:")
    print(f"  Due today (cut-off): {dt_n} orders, {dt_l} lines, {dt_u} units")
    print(f"  Total pick queue:    {pq_n} orders, {pq_l} lines, {pq_u} units")
    print(f"  Late/failing KPI:    {lp_n} orders")
    print(f"  Completed today:     {ct_n} orders")
    print(f"\nPurchase Orders:")
    print(f"  Expected not arrived: {ex_n}")
    print(f"  In receipt: {ir_n}")
    print(f"  Failing KPI: {fp_n}")
    print(f"\n12-month volume: {len(daily_volume)} days of data")

    # Serialise — dates must be strings for JSON
    data = {
        "generated":    now.isoformat(),
        "due_today":    due_today,
        "pick_queue":   all_queue,
        "late_pick":    late,
        "completed":    completed,
        "expected":     expected,
        "in_receipt":   in_receipt,
        "failing_po":   failing_po,
        "daily_volume": daily_volume,
    }

    html = HTML_TEMPLATE.replace(
        "__DATA_JSON__",
        json.dumps(data, default=str),
    ).replace(
        "__GENERATED__",
        now.strftime("%d %b %Y %H:%M NZT"),
    )

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\n✓ Wrote {OUTPUT_FILE}")


if __name__ == "__main__":
    main()

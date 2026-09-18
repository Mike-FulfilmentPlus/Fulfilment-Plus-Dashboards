"""
Compute "month to date" KPI figures for the current (in-progress) month and
write them to mtd_kpis.json, which generate_dashboards.py picks up to show a
"where we're sitting this month" badge alongside last month's completed
figures.

Dock to Stock and Order Cut Off are computed from live, per-order CartonCloud
data. Orders whose outcome can't yet be determined (still inside their Dock
to Stock 48hr window, or before their Order Cut Off deadline and not yet
dispatched) are excluded from the MTD denominator - they're "in progress",
not misses.

Order Accuracy by Unit is computed from the Issue Log tab (units wrong,
logged manually as issues are found) against units dispatched this month
per customer, sourced from the shared volume_data.py cache (see that
module's docstring - it's the same daily-dispatched-order cache the
warehouse dashboard's 12-month volume chart uses, so this costs no extra
CartonCloud API calls beyond what's already being made). Customers with no
recorded dispatch volume yet this month are skipped rather than shown as
0/0.

Freight DIFOT is NOT computed by this script - see mtd_difot.py, which is
run as a separate follow-up step (Starshipit tracking lookups take longer
than fits in one run) and merges its own "Freight DIFOT" key into this same
mtd_kpis.json file. Because of that, this script MERGES into any existing
mtd_kpis.json rather than overwriting it wholesale, so an ad-hoc run of just
this script never wipes out Freight DIFOT data that mtd_difot.py already
wrote. The normal daily pipeline (see the mtd-dashboard-refresh scheduled
task) runs: mtd_kpis.py, then mtd_difot.py --fetch-only (x2), then
mtd_difot.py --compute-only, then generate_dashboards.py - in that order.

Stock Accuracy and Cycle Count have no live MTD data source and are omitted
from mtd_kpis.json; the dashboard falls back to "no MTD data" for them.

Usage:
    python mtd_kpis.py
"""

import re
import csv
import os
import json
import datetime
import requests
import openpyxl
from zoneinfo import ZoneInfo

import volume_data

WORKBOOK = "kpi_dashboard.xlsx"
CREDS_FILE = "credentials.env"
OUTPUT_FILE = "mtd_kpis.json"
AWAITING_STOCK_LOG = "awaiting_stock_log.csv"

TEST_ACCOUNT_NAMES = {"TEST ACCOUNT"}
# Orders in these statuses are not real fulfilment activity yet and must
# never count in any KPI:
#   DRAFT           - order held back, often missing product
#   AWAITING_STOCK  - explicitly waiting on inventory
#   REJECTED        - stuck on an unresolved allocation/inventory error;
#                      never actually dispatched, so it can't be scored on
#                      dispatch timing (this is an ops problem to chase
#                      separately, not an Order Cut Off miss)
# An order only becomes "in play" for Order Cut Off once it reaches
# AWAITING_PICK_AND_PACK (see daily_status_snapshot.py's IN_PLAY_STATUSES
# and load_awaiting_stock_transitions() below for the clock-start logic).
NON_KPI_STATUSES = {"DRAFT", "AWAITING_STOCK", "REJECTED"}

# Inbound (purchase order) equivalent for Dock to Stock: NOT_YET_RECEIVED
# hasn't started the clock yet; DRAFT/REJECTED are never real receiving
# activity (mirrors NON_KPI_STATUSES above, and generate_warehouse_dashboard.py's
# fetch_inbound_open(), which already excludes DRAFT/REJECTED from its own
# "open POs" view for the same reason).
INBOUND_NON_KPI_STATUSES = {"NOT_YET_RECEIVED", "DRAFT", "REJECTED"}

CUTOFF_HOUR = 14  # 2pm — submission deadline
COB_HOUR = 20    # 8pm — close of business dispatch deadline

NZ_TZ = ZoneInfo("Pacific/Auckland")

WEEKDAY_NAMES = {
    "monday": 0, "tuesday": 1, "wednesday": 2, "thursday": 3,
    "friday": 4, "saturday": 5, "sunday": 6,
}

# ── Rural / SI detection for required-date deadline logic ─────────────────────
import json as _json

RURAL_POSTCODE_FILE = "nz_rural_postcodes.json"
RURAL_STREET_FILE   = "NZ Street File - Zone Guide.xlsx"
_RURAL_RE = re.compile(r'\bR\.?D\.?\s*\d', re.IGNORECASE)
_STREET_ABBREV = [(' rd',' road'),(' st',' street'),(' ave',' avenue'),
                  (' cres',' crescent'),(' dr',' drive'),(' tce',' terrace')]

def _norm_street_co(s):
    s = s.lower().strip()
    for a, b in _STREET_ABBREV:
        if s.endswith(a): return s[:-len(a)] + b
    return s

def _extract_street_co(addr):
    s = re.sub(r'^\d+[a-zA-Z]?\s+', '', addr.strip())
    return _norm_street_co(s.split(',')[0])

def load_rural_postcode_set_co(path=RURAL_POSTCODE_FILE):
    try:
        with open(path) as f:
            return set(_json.load(f)["postcodes"])
    except Exception:
        return set()

def load_rural_street_lookup_co(path=RURAL_STREET_FILE):
    try:
        import openpyxl as _ox
        wb = _ox.load_workbook(path, data_only=True, read_only=True)
        ws = wb.active
        lookup = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) < 11 or not row[10]: continue
            flag = str(row[10]).strip().lower()
            if flag not in ("rural", "rural / non-urban", "non-urban"): continue
            pc = str(row[3] or "").strip().zfill(4)
            st = _extract_street_co(str(row[0] or ""))
            if pc and st: lookup.add((pc, st))
        wb.close()
        return lookup
    except Exception:
        return set()

def is_rural_co(postcode, street, city, rural_pcs, rural_streets):
    pc = str(postcode or "").strip().zfill(4)
    if pc in rural_pcs: return True
    st = _extract_street_co(str(street or ""))
    if st and (pc, st) in rural_streets: return True
    combined = f"{street} {city}"
    if _RURAL_RE.search(combined): return True
    return False

def is_si_co(postcode):
    try: return int(str(postcode or "").strip()) >= 7000
    except: return False

CONTENT_MATCH_FIELDS = {
    "Delivery Company Name": lambda o: o.get("details", {}).get("deliver", {}).get("address", {}).get("companyName", ""),
    # Matches every order for the customer, regardless of content - used for
    # blanket customer-wide holds (e.g. a customs/stock embargo affecting all
    # of that customer's orders in a date window). "Contains" should be "all".
    "All Orders": lambda o: "all",
}


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
    from cc_auth import get_cc_token as _get_cc_token
    return _get_cc_token(creds)


def search_all_pages(token, tenant_id, resource, condition, size=100):
    results = []
    page = 1
    while True:
        resp = requests.post(
            f"https://api.cartoncloud.com/tenants/{tenant_id}/{resource}/search",
            params={"size": size, "page": page},
            headers={
                "Accept-Version": "1",
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
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


def parse_iso(ts):
    return datetime.datetime.fromisoformat(ts)


def is_business_day(d, holidays=None):
    if d.weekday() >= 5:
        return False
    if holidays and d in holidays:
        return False
    return True


def next_business_day(d, holidays=None):
    nd = d + datetime.timedelta(days=1)
    while not is_business_day(nd, holidays):
        nd += datetime.timedelta(days=1)
    return nd


def business_days_before(d, n, holidays=None):
    nd = d
    remaining = n
    while remaining > 0:
        nd -= datetime.timedelta(days=1)
        if is_business_day(nd, holidays):
            remaining -= 1
    return nd


def next_weekday_on_or_after(d, target_weekday, holidays=None):
    delta = (target_weekday - d.weekday()) % 7
    if delta == 0:
        delta = 7  # "next Tuesday" means following Tuesday, not today
    nd = d + datetime.timedelta(days=delta)
    while holidays and nd in holidays:
        nd = next_business_day(nd, holidays)
    return nd


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


def deadline_for_rule(created_dt, rule, holidays=None):
    if rule["type"] == "fixed":
        # Absolute one-off deadline (e.g. "Fixed (2026-09-02 20:00)") - used for
        # customer-wide operational holds (stock embargo/customs hold) where every
        # order in the effective window should get the same real-world deadline,
        # regardless of which day within the window it happened to be created.
        return rule["when"].replace(tzinfo=NZ_TZ)
    if rule["type"] == "next_weekday":
        target_date = next_weekday_on_or_after(created_dt.date(), rule["weekday"], holidays)
        return datetime.datetime.combine(target_date, datetime.time(rule["hour"], rule["minute"]), tzinfo=NZ_TZ)
    if rule["type"] == "next_business_day_eod":
        nbd = next_business_day(created_dt.date(), holidays)
        return datetime.datetime.combine(nbd, datetime.time(23, 59, 59), tzinfo=NZ_TZ)
    return datetime.datetime.combine(created_dt.date(), datetime.time(23, 59, 59), tzinfo=NZ_TZ)


def order_cut_off_deadline(created_dt, cutoff_hour=CUTOFF_HOUR, lenient_next_day=False,
                           holidays=None, after_cutoff_extra_days=0):
    is_weekend = created_dt.weekday() >= 5
    is_holiday = holidays and created_dt.date() in holidays
    if (not is_weekend) and (not is_holiday) and created_dt.hour < cutoff_hour:
        if lenient_next_day:
            nbd = next_business_day(created_dt.date(), holidays)
            return datetime.datetime.combine(nbd, datetime.time(COB_HOUR, 0, 0), tzinfo=NZ_TZ)
        return datetime.datetime.combine(created_dt.date(), datetime.time(COB_HOUR, 0, 0), tzinfo=NZ_TZ)
    nbd = created_dt.date()
    for _ in range(1 + after_cutoff_extra_days):
        nbd = next_business_day(nbd, holidays)
    return datetime.datetime.combine(nbd, datetime.time(COB_HOUR, 0, 0), tzinfo=NZ_TZ)


def load_exceptions(wb):
    result = {"cutoff_overrides": {}, "order_exceptions": [], "content_overrides": {}, "public_holidays": set()}
    if "Exceptions" not in wb.sheetnames:
        return result
    ws = wb["Exceptions"]

    section1_header_row = None
    section2_header_row = None
    section3_header_row = None
    section6_header_row = None
    section7_header_row = None
    for row in range(1, ws.max_row + 1):
        val = ws.cell(row=row, column=1).value
        if val == "Customer Cutoff Time Overrides":
            section1_header_row = row + 1
        elif val == "Order-Level Exceptions":
            section2_header_row = row + 1
        elif val == "Order Content-Based Overrides":
            section3_header_row = row + 1
        elif val == "Table 6: Public Holidays":
            section6_header_row = row + 1
        elif val == "Table 7: Operational Exceptions (Non-Business Days)":
            section7_header_row = row + 1

    if section1_header_row:
        r = section1_header_row + 1
        while r <= ws.max_row:
            customer = ws.cell(row=r, column=1).value
            if not customer:
                break
            cutoff_str = ws.cell(row=r, column=2).value
            deadline_rule = ws.cell(row=r, column=3).value or "Same day"
            eff_from = ws.cell(row=r, column=4).value
            eff_to = ws.cell(row=r, column=5).value

            cutoff_hour, cutoff_minute = CUTOFF_HOUR, 0
            if cutoff_str:
                parts = str(cutoff_str).strip().split(":")
                cutoff_hour = int(parts[0])
                cutoff_minute = int(parts[1]) if len(parts) > 1 else 0

            lenient_next_day = str(deadline_rule).strip().lower().startswith("next business day")

            if isinstance(eff_from, datetime.datetime):
                eff_from = eff_from.date()
            if isinstance(eff_to, datetime.datetime):
                eff_to = eff_to.date()

            extra_days_raw = ws.cell(row=r, column=7).value
            try:
                after_cutoff_extra_days = int(extra_days_raw)
            except (TypeError, ValueError):
                after_cutoff_extra_days = 0

            use_packed = str(ws.cell(row=r, column=8).value or "").strip().upper() == "Y"

            result["cutoff_overrides"].setdefault(customer, []).append({
                "cutoff_hour": cutoff_hour,
                "cutoff_minute": cutoff_minute,
                "lenient_next_day": lenient_next_day,
                "effective_from": eff_from,
                "effective_to": eff_to,
                "after_cutoff_extra_days": after_cutoff_extra_days,
                "use_packed_timestamp": use_packed,
            })
            r += 1

    if section2_header_row:
        r = section2_header_row + 1
        action_map = {
            "exclude from calculation": "exclude",
            "count as met": "count_as_met",
        }
        while r <= ws.max_row:
            customer = ws.cell(row=r, column=2).value
            reference = ws.cell(row=r, column=3).value
            month = ws.cell(row=r, column=4).value
            kpi = ws.cell(row=r, column=5).value
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

    if section3_header_row:
        r = section3_header_row + 1
        while r <= ws.max_row:
            customer = ws.cell(row=r, column=1).value
            if not customer:
                break
            match_field = ws.cell(row=r, column=2).value
            contains = ws.cell(row=r, column=3).value
            deadline_rule_str = ws.cell(row=r, column=4).value
            eff_from = ws.cell(row=r, column=5).value
            eff_to = ws.cell(row=r, column=6).value

            if isinstance(eff_from, datetime.datetime):
                eff_from = eff_from.date()
            if isinstance(eff_to, datetime.datetime):
                eff_to = eff_to.date()

            result["content_overrides"].setdefault(customer, []).append({
                "match_field": match_field,
                "contains": str(contains).strip().lower() if contains else "",
                "rule": parse_deadline_rule(deadline_rule_str),
                "effective_from": eff_from,
                "effective_to": eff_to,
            })
            r += 1

    if section6_header_row:
        r = section6_header_row + 1  # skip column-header row
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
    if section7_header_row:
        r = section7_header_row + 1  # skip column-header row
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


def load_awaiting_stock_transitions(path=AWAITING_STOCK_LOG):
    """Return {order_id: transition_date} for orders that were held in
    AWAITING_STOCK and later transitioned to something else, per
    daily_status_snapshot.py's log. Only day-level granularity is available
    (the log is written by a once-daily poll), so the Order Cut Off clock
    for these orders is treated as starting at 00:00 NZ time on
    transition_date rather than at their original created timestamp -
    this is a documented approximation, not exact to the minute."""
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


def get_cutoff_override(overrides, customer, order_date):
    for override in overrides.get(customer, []):
        eff_from = override.get("effective_from")
        eff_to = override.get("effective_to")
        if eff_from and order_date < eff_from:
            continue
        if eff_to and order_date > eff_to:
            continue
        return override
    return None


def find_order_exception(exceptions_list, customer, reference, order_date, month_str):
    """Find an applicable Order-Level Exception for this order.

    Reference "*" matches every order reference for that customer (a
    blanket exception). If Date From/Date To are set on the exception row,
    the order's actual creation date must fall within that range - this
    takes priority over (and ignores) the Month column, letting an
    exception scope to a specific week rather than a whole month. If no
    date range is set, falls back to the original Month-based scoping.
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


def compute_dock_to_stock_mtd(inbound_orders, now, exceptions=None, month_str=None):
    """Returns {customer_name: (numerator, denominator)}. Orders that
    arrived less than 48 hours ago and haven't been verified yet are
    excluded (outcome not yet determined)."""
    exceptions = exceptions or {"order_exceptions": []}
    order_exceptions = {
        (e["customer"], e["reference"]): e
        for e in exceptions.get("order_exceptions", [])
        if e.get("kpi") == "Dock to Stock" and (month_str is None or e.get("month") == month_str)
    }

    stats = {}
    for o in inbound_orders:
        cust = o["customer"]["name"]
        if cust in TEST_ACCOUNT_NAMES:
            continue
        arrival_str = o.get("details", {}).get("arrivalDate")
        if not arrival_str:
            continue

        # Skip orders not yet physically received, or held in DRAFT/REJECTED
        # (never real receiving activity). Same treatment as NON_KPI_STATUSES
        # for outbound orders.
        if o.get("status") in INBOUND_NON_KPI_STATUSES:
            continue

        # Stock-adjustment records (details.isAdjustment) add/remove
        # inventory directly rather than representing a physical dock
        # delivery - they structurally never get a "verified" timestamp,
        # so they can never pass this KPI. Exclude entirely.
        if o.get("details", {}).get("isAdjustment"):
            continue

        reference = str(o.get("references", {}).get("numericId", ""))
        exc = order_exceptions.get((cust, reference))
        if exc and exc["action"] == "exclude":
            continue

        arrival_dt = datetime.datetime.fromisoformat(arrival_str + "T00:00:00").replace(tzinfo=NZ_TZ)

        if exc and exc["action"] == "count_as_met":
            met = True
        else:
            verified = o.get("timestamps", {}).get("verified", {}).get("time")
            if verified:
                verified_dt = parse_iso(verified).astimezone(NZ_TZ)
                hours = (verified_dt - arrival_dt).total_seconds() / 3600.0
                met = 0 <= hours <= 48
            else:
                hours_since_arrival = (now - arrival_dt).total_seconds() / 3600.0
                if hours_since_arrival < 48:
                    continue  # pending - still inside the 48hr window
                met = False

        num, den = stats.get(cust, (0, 0))
        den += 1
        if met:
            num += 1
        stats[cust] = (num, den)
    return stats


def compute_order_cut_off_mtd(outbound_orders, exceptions, now, month_str):
    """Returns {customer_name: (numerator, denominator)}. Orders not yet
    dispatched and still before their deadline are excluded (outcome not
    yet determined)."""
    _rural_pcs    = load_rural_postcode_set_co()
    _rural_streets = load_rural_street_lookup_co()
    overrides = exceptions["cutoff_overrides"]
    content_overrides = exceptions.get("content_overrides", {})
    holidays = exceptions.get("public_holidays", set())
    stock_transitions = load_awaiting_stock_transitions()

    order_results = []
    for o in outbound_orders:
        cust = o["customer"]["name"]
        if cust in TEST_ACCOUNT_NAMES:
            continue
        if o.get("status") in NON_KPI_STATUSES:
            continue
        created = o.get("timestamps", {}).get("created", {}).get("time")
        if not created:
            continue
        created_dt = parse_iso(created).astimezone(NZ_TZ)
        raw_created_date = created_dt.date()

        # If this order sat in AWAITING_STOCK before becoming pickable, the
        # Cut Off clock should start when it became pickable, not when it
        # was originally created (see daily_status_snapshot.py).
        transition_date = stock_transitions.get(o.get("id"))
        if transition_date and transition_date > created_dt.date():
            created_dt = datetime.datetime.combine(
                transition_date, datetime.time(0, 0, 0), tzinfo=NZ_TZ
            )

        deadline = order_cut_off_deadline(created_dt, holidays=holidays)

        # Overrides below are scoped (Effective From/To) against the order's
        # ORIGINAL creation date, not the stock-transition-adjusted created_dt.
        # Otherwise a stock-delayed order from an earlier/unrelated period can
        # accidentally drift into a later override's effective window purely
        # because its clock-start got pushed forward - e.g. an order genuinely
        # created before a customer-wide hold period started, but whose own
        # separate stock delay happens to shift its adjusted clock-start date
        # into the hold window, would wrongly inherit that override's deadline.
        override = get_cutoff_override(overrides, cust, raw_created_date)
        if override:
            override_deadline = order_cut_off_deadline(
                created_dt,
                cutoff_hour=override["cutoff_hour"],
                lenient_next_day=override["lenient_next_day"],
                holidays=holidays,
                after_cutoff_extra_days=override.get("after_cutoff_extra_days", 0),
            )
            deadline = max(deadline, override_deadline)

        required_date_str = o.get("details", {}).get("deliver", {}).get("requiredDate")
        if required_date_str:
            required_date = datetime.date.fromisoformat(required_date_str)
            addr = o.get("details", {}).get("deliver", {}).get("address", {})
            pc   = str(addr.get("postcode") or "").strip()
            st   = str(addr.get("address1") or "").strip()
            city = str(addr.get("city") or "").strip()
            si   = is_si_co(pc)
            rur  = is_rural_co(pc, st, city, _rural_pcs, _rural_streets)
            days = (3 if si else 1) + (1 if rur else 0)
            alt_date = business_days_before(required_date, days, holidays)
            alt_deadline = datetime.datetime.combine(alt_date, datetime.time(17, 0, 0), tzinfo=NZ_TZ)
            deadline = max(deadline, alt_deadline)

        for content_override in content_overrides.get(cust, []):
            eff_from = content_override.get("effective_from")
            eff_to = content_override.get("effective_to")
            if eff_from and raw_created_date < eff_from:
                continue
            if eff_to and raw_created_date > eff_to:
                continue
            getter = CONTENT_MATCH_FIELDS.get(content_override["match_field"])
            if not getter:
                continue
            field_value = (getter(o) or "").lower()
            if content_override["contains"] and content_override["contains"] in field_value:
                content_deadline = deadline_for_rule(created_dt, content_override["rule"], holidays)
                deadline = max(deadline, content_deadline)

        use_packed = override.get("use_packed_timestamp", False) if override else False
        dispatch_time = (
            o.get("timestamps", {}).get("packed", {}).get("time")
            if use_packed else None
        ) or o.get("timestamps", {}).get("dispatched", {}).get("time")

        reference = o.get("references", {}).get("numericId", "")
        if dispatch_time:
            dispatched_dt = parse_iso(dispatch_time).astimezone(NZ_TZ)
            met = dispatched_dt <= deadline
        else:
            if now <= deadline:
                continue  # pending - deadline hasn't passed yet
            met = False

        order_results.append({"customer": cust, "reference": str(reference), "met": met, "date": raw_created_date})

    cut_off_exceptions = [
        e for e in exceptions["order_exceptions"] if e["kpi"] == "Order Cut Off"
    ]

    stats = {}
    for r in order_results:
        exc = find_order_exception(cut_off_exceptions, r["customer"], r["reference"], r["date"], month_str)
        if exc and exc["action"] == "exclude":
            continue
        met = r["met"]
        if exc and exc["action"] == "count_as_met":
            met = True
        num, den = stats.get(r["customer"], (0, 0))
        den += 1
        if met:
            num += 1
        stats[r["customer"]] = (num, den)
    return stats


def load_issue_log(wb):
    """Read the Issue Log tab. Returns a list of dicts: date_logged (date),
    customer, order_reference (str), quantity_wrong (int), issue_details.

    Rows missing a date or customer are skipped (incomplete entry, not yet
    ready to count). A non-numeric or blank Quantity Wrong is treated as 0
    rather than raising - better to under-count a malformed row than crash
    the whole KPI run over one bad entry."""
    entries = []
    if "Issue Log" not in wb.sheetnames:
        return entries
    ws = wb["Issue Log"]
    r = 5  # row 4 is the header
    while r <= ws.max_row:
        date_logged = ws.cell(row=r, column=1).value
        customer = ws.cell(row=r, column=2).value
        if not date_logged and not customer:
            r += 1
            continue
        if isinstance(date_logged, datetime.datetime):
            date_logged = date_logged.date()
        order_reference = ws.cell(row=r, column=3).value
        qty_wrong_raw = ws.cell(row=r, column=4).value
        issue_details = ws.cell(row=r, column=5).value
        try:
            quantity_wrong = int(qty_wrong_raw)
        except (TypeError, ValueError):
            quantity_wrong = 0
        if customer and date_logged:
            entries.append({
                "date_logged": date_logged,
                "customer": customer,
                "order_reference": str(order_reference).strip() if order_reference else "",
                "quantity_wrong": quantity_wrong,
                "issue_details": issue_details,
            })
        r += 1
    return entries


def compute_order_accuracy_by_unit(issue_log, month_str, volume_cache=None, active_customers=None):
    """Returns {customer_name: (numerator, denominator)} for Order Accuracy
    by Unit, for `month_str` (YYYY-MM).

    Units dispatched comes from the shared volume_data.py cache (see that
    module - same data the warehouse dashboard's volume chart already
    fetches, so no extra API calls). Units wrong comes from summing
    Quantity Wrong on every Issue Log row whose Date Logged falls in
    month_str.

    `active_customers` should be the set of customers already known to have
    had activity this month (e.g. from that same run's Dock to Stock /
    Order Cut Off results) - a customer in this set with zero issues logged
    shows as 100% even if real unit-level volume data isn't available for
    this month (e.g. any month before the volume cache existed), rather
    than being silently omitted just because a newer data source doesn't
    reach back that far. In that case the denominator shown is a nominal
    1 (there's no real unit count to show) - the 100% is what matters.

    A customer with issues logged but no real volume data is still
    surfaced using the logged wrong-unit count as a floor, rather than
    disappearing or being misrepresented as 100% - a real problem
    shouldn't get masked just because volume data is missing."""
    units_wrong = {}
    for entry in issue_log:
        if entry["date_logged"].strftime("%Y-%m") != month_str:
            continue
        cust = entry["customer"]
        units_wrong[cust] = units_wrong.get(cust, 0) + entry["quantity_wrong"]

    units_dispatched = volume_data.all_customers_units_for_month(month_str, cache=volume_cache)

    all_customers = set(units_dispatched) | set(units_wrong) | set(active_customers or [])

    stats = {}
    for cust in all_customers:
        dispatched = units_dispatched.get(cust, 0)
        wrong = units_wrong.get(cust, 0)
        if dispatched > 0:
            good = max(dispatched - wrong, 0)  # clamp - a data-entry error shouldn't produce a negative numerator
            stats[cust] = (good, dispatched)
        elif wrong > 0:
            stats[cust] = (0, wrong)
        else:
            stats[cust] = (1, 1)  # no issues, no real volume data - default to 100%
    return stats


def main():
    now = datetime.datetime.now(NZ_TZ)
    month_start = now.date().replace(day=1)
    month_str = month_start.strftime("%Y-%m")
    today_iso = now.date().isoformat()
    tomorrow_iso = (now.date() + datetime.timedelta(days=1)).isoformat()

    creds = load_credentials()
    token = get_cc_token(creds)
    tenant_id = creds["CARTONCLOUD_TENANT_ID"]

    print(f"Month to date: {month_start.isoformat()} to {today_iso} (NZ)")

    print("Fetching inbound orders (arrival this month to date)...")
    inbound = search_all_pages(
        token, tenant_id, "inbound-orders",
        date_range_condition("/details/arrivalDate", month_start.isoformat(), tomorrow_iso),
    )
    print(f"  {len(inbound)} inbound orders")

    print("Fetching outbound orders (created this month to date)...")
    outbound = search_all_pages(
        token, tenant_id, "outbound-orders",
        date_range_condition("/timestamps/created/time", month_start.isoformat(), tomorrow_iso),
    )
    print(f"  {len(outbound)} outbound orders")

    wb = openpyxl.load_workbook(WORKBOOK)
    exceptions = load_exceptions(wb)

    dock_to_stock = compute_dock_to_stock_mtd(inbound, now, exceptions=exceptions, month_str=month_str)
    order_cut_off = compute_order_cut_off_mtd(outbound, exceptions, now, month_str)

    print("\nDock to Stock (MTD):")
    for cust, (n, d) in sorted(dock_to_stock.items()):
        print(f"  {cust}: {n}/{d} = {n/d:.1%}")

    print("\nOrder Cut Off (MTD):")
    for cust, (n, d) in sorted(order_cut_off.items()):
        print(f"  {cust}: {n}/{d} = {n/d:.1%}")

    # Order Accuracy by Unit - shares the volume cache with the warehouse
    # dashboard's 12-month chart, so refreshing it here costs no extra
    # CartonCloud calls beyond whatever that dashboard already made today.
    # active_customers (from Dock to Stock / Order Cut Off, already computed
    # above) lets a customer with zero issues show 100% even before real
    # per-customer volume data exists for them this month.
    volume_data.refresh_volume(token, tenant_id, now.date())
    volume_cache = volume_data.load_volume_cache()
    issue_log = load_issue_log(wb)
    active_customers = set(dock_to_stock.keys()) | set(order_cut_off.keys())
    order_accuracy = compute_order_accuracy_by_unit(
        issue_log, month_str, volume_cache=volume_cache, active_customers=active_customers
    )

    units_wrong_customers = {
        e["customer"] for e in issue_log if e["date_logged"].strftime("%Y-%m") == month_str
    }
    dispatched_this_month = volume_data.all_customers_units_for_month(month_str, cache=volume_cache)
    # A customer whose Order Accuracy denominator would equal its own
    # logged wrong-unit total (rather than a real dispatched-units figure)
    # is being approximated, not measured precisely - worth a heads-up.
    approximated = {
        cust for cust in units_wrong_customers
        if dispatched_this_month.get(cust, 0) <= 0
    }
    if approximated:
        print(f"\nNOTE: issues logged this month for customers with no recorded dispatch "
              f"volume yet - shown using the logged wrong-unit count as an approximate "
              f"denominator rather than a real units-dispatched figure (check for typos, "
              f"or this may just be a month/customer the volume cache doesn't cover yet): "
              f"{sorted(approximated)}")

    if order_accuracy:
        print("\nOrder Accuracy by Unit (MTD):")
        for cust, (n, d) in sorted(order_accuracy.items()):
            print(f"  {cust}: {n}/{d} = {n/d:.1%}")

    # Freight DIFOT MTD is NOT computed here - it's handled by the separate
    # mtd_difot.py script (Starshipit tracking lookups take too long to fit
    # in one run, so it uses its own resumable cache and is run as a
    # follow-up step: `mtd_kpis.py` then `mtd_difot.py --fetch-only` (x2)
    # then `mtd_difot.py --compute-only`, per the mtd-dashboard-refresh
    # scheduled task). mtd_difot.py merges its "Freight DIFOT" key into
    # this same mtd_kpis.json file rather than overwriting it - so this
    # script must do the same (merge, not overwrite) or it would wipe out
    # Freight DIFOT every time it runs, even though it has nothing to do
    # with DIFOT at all.
    existing = {}
    if os.path.exists(OUTPUT_FILE):
        try:
            with open(OUTPUT_FILE) as f:
                existing = json.load(f)
        except (OSError, json.JSONDecodeError):
            existing = {}

    out = existing
    out["month"] = month_str
    out["as_of"] = now.isoformat()
    out.setdefault("kpis", {})
    out["kpis"]["Dock to Stock"] = {cust: {"num": n, "den": d} for cust, (n, d) in dock_to_stock.items()}
    out["kpis"]["Order Cut Off"] = {cust: {"num": n, "den": d} for cust, (n, d) in order_cut_off.items()}
    out["kpis"]["Order Accuracy by Unit"] = {cust: {"num": n, "den": d} for cust, (n, d) in order_accuracy.items()}
    # "Freight DIFOT" under out["kpis"], if already present from a prior
    # mtd_difot.py run, is left untouched by simply not being reassigned.

    with open(OUTPUT_FILE, "w") as f:
        json.dump(out, f, indent=2, default=str)
    print(f"Wrote {OUTPUT_FILE}")


if __name__ == "__main__":
    main()

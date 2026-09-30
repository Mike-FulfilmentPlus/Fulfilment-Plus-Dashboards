"""
Extract KPI data from CartonCloud for a given month and append it to the
KPI_Data tab of kpi_dashboard.xlsx.

Currently computes:
  - Dock to Stock          (target 0.95)  - inbound orders verified/put-away within 48hrs of arrival date
  - Order Cut Off         (target 0.95)  - outbound orders created before 2pm NZ time that dispatched same NZ day
  - Order Accuracy by Unit (target 0.995) - units dispatched minus units wrong (from the Issue Log tab),
    over units dispatched (from the shared volume_data.py cache). Only covers months from when the
    per-customer volume cache started accumulating (see volume_data.py) - earlier months will simply
    have no Order Accuracy by Unit row rather than an error.

Stock Accuracy, Cycle Count and Freight DIFOT are not yet reliably computable
from the API and remain manual entries (see Read Me tab).

Usage:
    python extract_kpi_data.py 2026-05
"""

import re
import sys
import csv
import os
import datetime
import requests
import openpyxl
from zoneinfo import ZoneInfo

import volume_data

WORKBOOK = "kpi_dashboard.xlsx"
CREDS_FILE = "credentials.env"
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

# CartonCloud timestamps are returned with a +10:00 offset, but the warehouse
# operates on NZ time. Convert to NZ local time before applying the 2pm
# cutoff / same-day checks for Order Cut Off.
NZ_TZ = ZoneInfo("Pacific/Auckland")

WEEKDAY_NAMES = {
    "monday": 0, "tuesday": 1, "wednesday": 2, "thursday": 3,
    "friday": 4, "saturday": 5, "sunday": 6,
}

# Order Content-Based Overrides: Match Field -> function extracting the
# value to test against "Contains" from an outbound order dict.
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
    resp = requests.post(
        "https://api.cartoncloud.com/uaa/oauth/token",
        auth=(creds["CARTONCLOUD_CLIENT_ID"], creds["CARTONCLOUD_CLIENT_SECRET"]),
        headers={"Accept-Version": "1"},
        data={"grant_type": "client_credentials"},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def search_all_pages(token, tenant_id, resource, condition, size=100):
    """Fetch all pages from a CartonCloud search endpoint."""
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


def month_bounds(month_str):
    """month_str = 'YYYY-MM' -> (start_date_str, end_date_str, month_first_of_month_date)"""
    year, month = (int(x) for x in month_str.split("-"))
    start = datetime.date(year, month, 1)
    if month == 12:
        end = datetime.date(year + 1, 1, 1)
    else:
        end = datetime.date(year, month + 1, 1)
    return start.isoformat(), end.isoformat(), start


def parse_iso(ts):
    return datetime.datetime.fromisoformat(ts)


def compute_dock_to_stock(inbound_orders, exceptions=None, month_str=None):
    """Returns {customer_name: (numerator, denominator)}"""
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
        # (never real receiving activity - they haven't started the clock,
        # or never will until someone resolves them). Same treatment as
        # NON_KPI_STATUSES for outbound orders.
        if o.get("status") in INBOUND_NON_KPI_STATUSES:
            continue

        # Stock-adjustment records (details.isAdjustment) add/remove
        # inventory directly rather than representing a physical dock
        # delivery. They can reach ALLOCATED but structurally never get a
        # "verified" timestamp (there's nothing to verify on arrival), so
        # they can never pass this KPI - exclude them entirely rather than
        # count every one as an automatic miss.
        if o.get("details", {}).get("isAdjustment"):
            continue

        reference = str(o.get("references", {}).get("numericId", ""))
        exc = order_exceptions.get((cust, reference))
        if exc and exc["action"] == "exclude":
            continue

        num, den = stats.get(cust, (0, 0))
        den += 1

        if exc and exc["action"] == "count_as_met":
            num += 1
        else:
            verified = o.get("timestamps", {}).get("verified", {}).get("time")
            if verified:
                verified_dt = parse_iso(verified)
                arrival_dt = datetime.datetime.fromisoformat(arrival_str + "T00:00:00").replace(
                    tzinfo=verified_dt.tzinfo
                )
                hours = (verified_dt - arrival_dt).total_seconds() / 3600.0
                if 0 <= hours <= 48:
                    num += 1
        stats[cust] = (num, den)
    return stats


def is_business_day(d, holidays=None):
    """True if d is Mon-Fri and not a public holiday."""
    if d.weekday() >= 5:
        return False
    if holidays and d in holidays:
        return False
    return True


def next_business_day(d, holidays=None):
    """Next Mon-Fri non-holiday date after date d."""
    nd = d + datetime.timedelta(days=1)
    while not is_business_day(nd, holidays):
        nd += datetime.timedelta(days=1)
    return nd


def business_days_before(d, n, holidays=None):
    """Date that is n business days (Mon-Fri, non-holiday) before date d."""
    nd = d
    remaining = n
    while remaining > 0:
        nd -= datetime.timedelta(days=1)
        if is_business_day(nd, holidays):
            remaining -= 1
    return nd


def next_weekday_on_or_after(d, target_weekday, holidays=None):
    """Date of the next occurrence of target_weekday (0=Mon..6=Sun) on or
    after date d. If that date is a public holiday, rolls forward to the
    next business day."""
    delta = (target_weekday - d.weekday()) % 7
    if delta == 0:
        delta = 7  # "next Tuesday" means the following Tuesday, not today -
                   # matches mtd_kpis.py; without this an order created on
                   # the target weekday itself gets a same-day deadline
                   # instead of a week's grace.
    nd = d + datetime.timedelta(days=delta)
    # If it lands on a holiday, roll to next business day
    while holidays and nd in holidays:
        nd = next_business_day(nd, holidays)
    return nd


def parse_time_of_day(text):
    """Parse '5pm', '5:00pm', '17:00' -> (hour, minute). Defaults to 23:59
    if unparseable."""
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
    """
    Parse a Deadline Rule string from the Exceptions tab into a dict:
      {"type": "same_day"}
      {"type": "next_business_day_eod"}
      {"type": "next_weekday", "weekday": 0-6, "hour": int, "minute": int}
    """
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
    """Given a parsed deadline rule and an order's created time (NZ local),
    return the deadline datetime. Public holidays are skipped."""
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
    """
    Return the COB (8pm NZ) deadline by which an order must be dispatched:
      - Before cutoff_hour on a business day → COB same day
        (if lenient_next_day=True → COB next business day instead)
      - At/after cutoff_hour, or weekend/holiday → COB next business day
        (+after_cutoff_extra_days additional business days for customers
        like Sikla whose SLA extends a further day when submitted late)
    """
    is_weekend = created_dt.weekday() >= 5
    is_holiday = holidays and created_dt.date() in holidays
    if (not is_weekend) and (not is_holiday) and created_dt.hour < cutoff_hour:
        if lenient_next_day:
            nbd = next_business_day(created_dt.date(), holidays)
            return datetime.datetime.combine(nbd, datetime.time(COB_HOUR, 0, 0), tzinfo=NZ_TZ)
        return datetime.datetime.combine(created_dt.date(), datetime.time(COB_HOUR, 0, 0), tzinfo=NZ_TZ)
    # After cutoff: advance 1 + extra business days
    nbd = created_dt.date()
    for _ in range(1 + after_cutoff_extra_days):
        nbd = next_business_day(nbd, holidays)
    return datetime.datetime.combine(nbd, datetime.time(COB_HOUR, 0, 0), tzinfo=NZ_TZ)


def load_exceptions(wb):
    """
    Read the Exceptions tab and return:
      {
        "cutoff_overrides": {customer_name: [override, ...]},
        "order_exceptions": [exception, ...],
        "content_overrides": {customer_name: [content_override, ...]},
      }

    Each override dict has: cutoff_hour, cutoff_minute, lenient_next_day,
    effective_from (date), effective_to (date or None).

    Each order exception dict has: customer, reference (str), month
    (YYYY-MM str), kpi, action ("exclude" or "count_as_met").

    Each content override dict has: match_field, contains (lowercased str),
    rule (parsed deadline rule dict), effective_from (date or None),
    effective_to (date or None).
    """
    result = {"cutoff_overrides": {}, "order_exceptions": [], "content_overrides": {}, "public_holidays": set()}
    if "Exceptions" not in wb.sheetnames:
        return result
    ws = wb["Exceptions"]

    # Locate the section headers by scanning column A.
    section1_header_row = None
    section2_header_row = None
    section3_header_row = None
    section6_header_row = None
    section7_header_row = None
    for row in range(1, ws.max_row + 1):
        val = ws.cell(row=row, column=1).value
        if val == "Customer Cutoff Time Overrides":
            section1_header_row = row + 1  # header row follows section title
        elif val == "Order-Level Exceptions":
            section2_header_row = row + 1
        elif val == "Order Content-Based Overrides":
            section3_header_row = row + 1
        elif val == "Table 6: Public Holidays":
            section6_header_row = row + 1
        elif val == "Table 7: Operational Exceptions (Non-Business Days)":
            section7_header_row = row + 1

    # --- Customer Cutoff Time Overrides ---
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
                cutoff_str = str(cutoff_str).strip()
                parts = cutoff_str.split(":")
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

            override = {
                "cutoff_hour": cutoff_hour,
                "cutoff_minute": cutoff_minute,
                "lenient_next_day": lenient_next_day,
                "effective_from": eff_from,
                "effective_to": eff_to,
                "after_cutoff_extra_days": after_cutoff_extra_days,
                "use_packed_timestamp": use_packed,
            }
            result["cutoff_overrides"].setdefault(customer, []).append(override)
            r += 1

    # --- Order-Level Exceptions ---
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

    # --- Order Content-Based Overrides ---
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

    # --- Public Holidays ---
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

    # --- Operational Exceptions (Non-Business Days) ---
    # Deliberately NOT logged as public holidays (they aren't) - e.g. a one-off
    # warehouse outage - but merged into the same date set so every deadline
    # calculation treats them identically to a public holiday (rolls forward
    # to the next genuine business day).
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


def get_cutoff_override(overrides, customer, order_date):
    """Return the active override dict for customer/order_date, or None."""
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


def compute_order_cut_off(outbound_orders, exceptions=None, month_str=None, now=None):
    """Returns {customer_name: (numerator, denominator)}. Orders not yet
    dispatched whose deadline hasn't passed yet are excluded (outcome not
    yet determined - "pending", not a miss). This matters even for a
    "closed" month: some orders' deadlines legitimately fall after month-end
    (e.g. a lenient content-based override), so being created in the target
    month doesn't guarantee the deadline has already arrived by the time
    this script runs."""
    if now is None:
        now = datetime.datetime.now(NZ_TZ)
    _rural_pcs    = load_rural_postcode_set_co()
    _rural_streets = load_rural_street_lookup_co()
    exceptions = exceptions or {"cutoff_overrides": {}, "order_exceptions": [], "content_overrides": {}, "public_holidays": set()}
    overrides = exceptions["cutoff_overrides"]
    content_overrides = exceptions.get("content_overrides", {})
    holidays = exceptions.get("public_holidays", set())
    stock_transitions = load_awaiting_stock_transitions()

    # First pass: compute met/unmet per order.
    order_results = []  # list of dicts: customer, reference, met
    for o in outbound_orders:
        cust = o["customer"]["name"]
        if cust in TEST_ACCOUNT_NAMES:
            continue
        if o.get("status") in NON_KPI_STATUSES:
            continue
        created = o.get("timestamps", {}).get("created", {}).get("time")
        if not created:
            continue
        # Convert from CartonCloud's +10:00 offset to NZ local time.
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

        # Standard deadline.
        deadline = order_cut_off_deadline(created_dt, holidays=holidays)

        # Overrides below are scoped (Effective From/To) against the order's
        # ORIGINAL creation date, not the stock-transition-adjusted created_dt.
        # Otherwise a stock-delayed order from an earlier/unrelated period can
        # accidentally drift into a later override's effective window purely
        # because its clock-start got pushed forward.
        # Customer-specific cutoff override (only if more lenient).
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

        # Customers loading orders ahead of a delivery date: deadline is
        # N business days before the delivery date, at 5pm NZ (only if
        # more lenient than the standard deadline).
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

        # Order content-based overrides (e.g. orders to a particular
        # delivery destination get a different deadline), only if more
        # lenient than the deadline so far.
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

        # Use packed timestamp for customers whose dispatched timestamp is
        # only set on manual invoice print (not at actual fulfilment time).
        use_packed = override.get("use_packed_timestamp", False) if override else False
        dispatch_time = (
            o.get("timestamps", {}).get("packed", {}).get("time")
            if use_packed else None
        ) or o.get("timestamps", {}).get("dispatched", {}).get("time")

        if dispatch_time:
            dispatched_dt = parse_iso(dispatch_time).astimezone(NZ_TZ)
            met = dispatched_dt <= deadline
        else:
            if now <= deadline:
                continue  # pending - deadline hasn't passed yet
            met = False

        reference = o.get("references", {}).get("numericId", "")
        order_results.append({"customer": cust, "reference": str(reference), "met": met, "date": raw_created_date})

    # Apply order-level exceptions for this KPI/month.
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
    the whole extraction over one bad entry."""
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

    Units dispatched comes from the shared volume_data.py cache. Units
    wrong comes from summing Quantity Wrong on every Issue Log row whose
    Date Logged falls in month_str.

    `active_customers` should be the set of customers already known to have
    had activity this month (e.g. that same run's Dock to Stock / Order Cut
    Off results) - a customer in this set with zero issues shows as 100%
    even without real unit-level volume data for this month (e.g. any
    month before the per-customer volume cache existed - see
    volume_data.py), rather than being omitted just because a newer data
    source doesn't reach back that far. The denominator shown in that case
    is a nominal 1 - there's no real unit count to show, the 100% is what
    matters.

    A customer with issues logged but no real volume data is still
    surfaced using the logged wrong-unit count as a floor, rather than
    disappearing or being misrepresented as 100%."""
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


def append_rows(ws, start_row, month_date, kpi_name, target, stats, carrier=None,
                 notes_below="Below target - investigate root cause and add to next month's action plan"):
    row = start_row
    for cust in sorted(stats.keys()):
        num, den = stats[cust]
        if den == 0:
            continue
        ws.cell(row=row, column=1).value = datetime.datetime(month_date.year, month_date.month, 1)
        ws.cell(row=row, column=1).number_format = "mmm\\-yyyy"
        ws.cell(row=row, column=2).value = cust
        ws.cell(row=row, column=3).value = kpi_name
        ws.cell(row=row, column=4).value = carrier
        ws.cell(row=row, column=5).value = target
        ws.cell(row=row, column=5).number_format = "0.0%"
        ws.cell(row=row, column=6).value = num
        ws.cell(row=row, column=7).value = den
        ws.cell(row=row, column=8).value = f"=F{row}/G{row}"
        ws.cell(row=row, column=8).number_format = "0.0%"
        ws.cell(row=row, column=9).value = f'=IF(H{row}>=E{row},"Yes","No")'
        actual = num / den
        ws.cell(row=row, column=10).value = notes_below if actual < target else None
        row += 1
    return row


def main():
    if len(sys.argv) != 2:
        print("Usage: python extract_kpi_data.py YYYY-MM")
        sys.exit(1)

    month_str = sys.argv[1]
    start_date, end_date, month_date = month_bounds(month_str)

    creds = load_credentials()
    token = get_cc_token(creds)
    tenant_id = creds["CARTONCLOUD_TENANT_ID"]

    print(f"Fetching inbound orders (arrival {start_date} to {end_date})...")
    inbound = search_all_pages(
        token, tenant_id, "inbound-orders",
        date_range_condition("/details/arrivalDate", start_date, end_date),
    )
    print(f"  {len(inbound)} inbound orders")

    print(f"Fetching outbound orders (created {start_date} to {end_date})...")
    outbound = search_all_pages(
        token, tenant_id, "outbound-orders",
        date_range_condition("/timestamps/created/time", start_date, end_date),
    )
    print(f"  {len(outbound)} outbound orders")

    wb = openpyxl.load_workbook(WORKBOOK)
    exceptions = load_exceptions(wb)

    dock_to_stock = compute_dock_to_stock(inbound, exceptions=exceptions, month_str=month_str)
    order_cut_off = compute_order_cut_off(outbound, exceptions=exceptions, month_str=month_str)

    print("\nDock to Stock:")
    for cust, (n, d) in sorted(dock_to_stock.items()):
        print(f"  {cust}: {n}/{d} = {n/d:.1%}")

    print("\nOrder Cut Off:")
    for cust, (n, d) in sorted(order_cut_off.items()):
        print(f"  {cust}: {n}/{d} = {n/d:.1%}")

    # Order Accuracy by Unit - shares the volume cache with the warehouse
    # dashboard's 12-month chart and mtd_kpis.py, so this costs no extra
    # CartonCloud calls beyond whatever already ran today.
    volume_data.refresh_volume(token, tenant_id, datetime.datetime.now(NZ_TZ).date())
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
    approximated = {cust for cust in units_wrong_customers if dispatched_this_month.get(cust, 0) <= 0}
    if approximated:
        print(f"\nNOTE: issues logged for {month_str} against customers with no recorded dispatch "
              f"volume for that month - shown using the logged wrong-unit count as an approximate "
              f"denominator rather than a real units-dispatched figure...: {sorted(approximated)}")

    if order_accuracy:
        print("\nOrder Accuracy by Unit:")
        for cust, (n, d) in sorted(order_accuracy.items()):
            print(f"  {cust}: {n}/{d} = {n/d:.1%}")
    else:
        print(f"\nOrder Accuracy by Unit: no data for {month_str} "
              f"(volume cache doesn't cover this month yet)")

    # Write results to KPI_Data sheet (remove existing rows for this month first)
    wb2 = openpyxl.load_workbook(WORKBOOK)
    ws_kpi = wb2["KPI_Data"]

    # Find header row and remove any existing rows for this month
    month_date_cmp = datetime.datetime(month_date.year, month_date.month, 1)
    rows_to_delete = []
    for r in range(2, ws_kpi.max_row + 1):
        cell_val = ws_kpi.cell(row=r, column=1).value
        kpi_val = ws_kpi.cell(row=r, column=3).value
        if (isinstance(cell_val, datetime.datetime) and
                cell_val.year == month_date.year and
                cell_val.month == month_date.month and
                kpi_val in ("Dock to Stock", "Order Cut Off", "Order Accuracy by Unit")):
            rows_to_delete.append(r)
    for r in reversed(rows_to_delete):
        ws_kpi.delete_rows(r)

    start_row = ws_kpi.max_row + 1
    start_row = append_rows(ws_kpi, start_row, month_date, "Dock to Stock", 0.95, dock_to_stock)
    start_row = append_rows(ws_kpi, start_row, month_date, "Order Cut Off", 0.95, order_cut_off)
    if order_accuracy:
        start_row = append_rows(ws_kpi, start_row, month_date, "Order Accuracy by Unit", 0.995, order_accuracy)

    wb2.save(WORKBOOK)
    print(f"\nWrote results to {WORKBOOK}")
    print("Note: run extract_difot.py to add Freight DIFOT for this month.")


if __name__ == "__main__":
    main()

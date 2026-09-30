"""
extract_difot.py  -  Freight DIFOT extraction from Starshipit Delivery Performance CSV

Usage:
    python extract_difot.py YYYY-MM [path/to/DeliveryPerformanceReport.xlsx]

Delivery SLAs (from dispatch / carrier pickup date):
  NZ North Island (postcode < 7000):    1 business day
  NZ South Island (postcode >= 7000):   3 business days
  Rural (+Rural flag):                  +2 business days
  DG (dangerous goods):                 +2 business days
  FedEx International Connect Plus:     3 business days (FedEx: "2-3 business days" - using max)

DC recipient exceptions (Table 4 in Exceptions tab):
  Supercheap/GPC/SRG/Repco DC orders are counted as met and attributed directly to Oricom
  via the "Attribute To" column, bypassing the CartonCloud numericId join.

International carrier SLAs (Table 5 in Exceptions tab):
  Documents the business-day targets used for non-NZ shipments by carrier service code.
"""

import sys
import os
import glob
import json
import re
import datetime
import requests
import openpyxl
from zoneinfo import ZoneInfo
from collections import defaultdict

# -- Constants ------------------------------------------------------------------
WORKBOOK      = "kpi_dashboard.xlsx"
CREDS_FILE    = "credentials.env"
CARRIER_FILE  = "difot_carriers.json"
KPI_NAME      = "Freight DIFOT"
TEST_ACCOUNT_NAMES = {"TEST ACCOUNT"}
# Customers whose real deliveries are truck freight (e.g. Cargo Plus) rather than
# Starshipit-tracked parcel carriers (NZ Couriers, Post Haste, etc). CartonCloud's
# references.numericId is only unique per-customer, not tenant-wide, so when one of
# these customers' own numericId happens to match a genuine parcel customer's order
# in the same period, the Starshipit report's TheirRef join silently misattributes
# that unrelated parcel's carrier/on-time result to the freight customer instead.
# Confirmed for Made Group NZ Aug 2026: order #1804 (26 cartons, truck freight to
# Otoki NZ Ltd, dispatched 7 Jul) collided with an unrelated NZ Couriers overnight
# parcel also numbered 1804 (delivered to Dave Gillies, Tauranga, 16 Jul) - inflating
# Made Group's June/July Freight DIFOT with carriers (NZ Couriers/Post Haste) they
# never actually use. Excluded entirely rather than scored.
FREIGHT_ONLY_CUSTOMERS = {"Made Group NZ"}
# Carrier used by freight-only customers' truck deliveries, attributed directly
# since Starshipit never sees these shipments at all.
FREIGHT_ONLY_CARRIER = "Cargo Plus"
NZ_COUNTRIES = {"New Zealand", "NEW ZEALAND", "new zealand"}
NZ_TZ = ZoneInfo("Pacific/Auckland")

# FedEx International Connect Plus: FedEx NZ publishes "typically 2-3 business days
# worldwide". Per FP policy we use the maximum of any published range.
INTL_CARRIER_DAYS = {
    "FEDEX_INTERNATIONAL_CONNECT_PLUS": 3,
}

# Starshipit column indices (0-based, Data Source sheet)
COL_OURREF         = 1
COL_THEIRREF       = 2
COL_SHIPTO         = 3
COL_STREET         = 4
COL_SUBURB         = 5
COL_CITY           = 6
COL_POSTCODE       = 8
COL_COUNTRY        = 9
COL_CARRIER        = 12   # carrier name (NZ Couriers, Post Haste, FedEx, ...)
COL_CODE           = 14   # carrier service code
COL_ORDER_SOURCE   = 21
COL_ORDER_DATE     = 23
COL_PICKUP_DATE    = 26
COL_DELIVERED_DATE = 28
COL_STATUS         = 29

RURAL_POSTCODE_FILE = "nz_rural_postcodes.json"
RURAL_STREET_FILE   = "NZ Street File - Zone Guide.xlsx"

# Rural delivery detection uses three methods in order:
#  1. NZ Post rural postcode lookup (definitive — all RD postcodes from NZP directory)
#  2. NZ Couriers zone file: (postcode, street) flagged rural in column K
#  3. RD pattern fallback: "RD"/"R.D." + digit anywhere in street/suburb/city
RURAL_RE = re.compile(r'\bR\.?D\.?\s*\d', re.IGNORECASE)

_STREET_ABBREV = [
    (' rd', ' road'), (' st', ' street'), (' ave', ' avenue'), (' av', ' avenue'),
    (' cres', ' crescent'), (' cr', ' crescent'), (' dr', ' drive'), (' pl', ' place'),
    (' tce', ' terrace'), (' hwy', ' highway'), (' ln', ' lane'),
]

def _norm_street(name):
    s = name.lower().strip()
    for abbr, full in _STREET_ABBREV:
        if s.endswith(abbr):
            return s[:-len(abbr)] + full
    return s

def _extract_street_name(addr):
    s = re.sub(r'^\d+[\w/-]*\s+', '', str(addr or '').strip())
    return s.split(',')[0].strip()

def load_rural_postcode_set(path=RURAL_POSTCODE_FILE):
    """Load the set of NZ rural delivery postcodes from the NZP directory JSON."""
    if not os.path.exists(path):
        print(f"Warning: {path} not found — rural postcode lookup disabled.")
        return set()
    with open(path) as f:
        data = json.load(f)
    return set(data["postcodes"])

def load_rural_street_lookup(path=RURAL_STREET_FILE):
    """Build a set of (postcode_4digit, normalised_street_lower) flagged rural in NZC zone file."""
    if not os.path.exists(path):
        return set()
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rural = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        code, street, rural_flag = row[0], row[2], str(row[10] or '')
        if not street or code is None:
            continue
        try:
            pc = str(int(code)).zfill(4)
        except (TypeError, ValueError):
            continue
        if rural_flag.startswith('Yes'):
            rural.add((pc, _norm_street(street)))
    wb.close()
    return rural

def is_rural_address(postcode, street, suburb, city, rural_postcodes, rural_street_lookup):
    """Return True if the address is a rural delivery."""
    pc_str = str(postcode or '').strip().zfill(4)
    # 1. NZ Post rural postcode directory (definitive)
    if pc_str and pc_str in rural_postcodes:
        return True
    # 2. NZ Couriers zone file (urban streets on rural routes)
    if rural_street_lookup and pc_str:
        street_name = _extract_street_name(street)
        if (pc_str, _norm_street(street_name)) in rural_street_lookup:
            return True
    # 3. RD pattern fallback
    if RURAL_RE.search(f"{street} {suburb} {city}"):
        return True
    return False


# -- Helpers --------------------------------------------------------------------

def load_credentials(path=CREDS_FILE):
    creds = {}
    with open(path) as f:
        for line in f:
            line = line.strip()
            if line and "=" in line and not line.startswith("#"):
                k, v = line.split("=", 1)
                creds[k.strip()] = v.strip()
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


def fetch_cc_customer_map(token, tenant_id, start_date, end_date):
    buf = datetime.timedelta(days=14)
    fetch_start = (start_date - buf).isoformat()
    fetch_end   = (end_date   + buf).isoformat()
    cond = {
        "type": "AndCondition",
        "conditions": [
            {"type": "DateComparisonCondition",
             "field": {"type": "JsonField", "pointer": "/timestamps/created/time"},
             "value": {"type": "ValueField", "value": fetch_start},
             "method": "GREATER_THAN_OR_EQUAL_TO"},
            {"type": "DateComparisonCondition",
             "field": {"type": "JsonField", "pointer": "/timestamps/created/time"},
             "value": {"type": "ValueField", "value": fetch_end},
             "method": "LESS_THAN"},
        ],
    }
    hdrs = {"Accept-Version": "1", "Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    cc_map = {}
    page = 1
    while True:
        resp = requests.post(
            f"https://api.cartoncloud.com/tenants/{tenant_id}/outbound-orders/search",
            params={"size": 200, "page": page}, headers=hdrs, json={"condition": cond}, timeout=60,
        )
        resp.raise_for_status()
        batch = resp.json()
        if not batch:
            break
        for o in batch:
            nid      = (o.get("references") or {}).get("numericId")
            customer = (o.get("customer") or {}).get("name", "")
            if nid and customer:
                cc_map[str(nid)] = customer
        if page >= int(resp.headers.get("total-pages", "1")):
            break
        page += 1
    return cc_map


def fetch_freight_only_counts(token, tenant_id, month_start, month_end, customer_names):
    """Count real dispatched orders for freight-only customers within the exact
    month window - counted directly from CartonCloud rather than joined via the
    collision-prone numericId, since these customers' truck deliveries (e.g. Cargo
    Plus) are never seen by Starshipit at all. Each dispatched order is treated as
    on-time (these customers have no live tracking to check otherwise)."""
    cond = {"type": "AndCondition", "conditions": [
        {"type": "DateComparisonCondition",
         "field": {"type": "JsonField", "pointer": "/timestamps/created/time"},
         "value": {"type": "ValueField", "value": month_start.isoformat()},
         "method": "GREATER_THAN_OR_EQUAL_TO"},
        {"type": "DateComparisonCondition",
         "field": {"type": "JsonField", "pointer": "/timestamps/created/time"},
         "value": {"type": "ValueField", "value": month_end.isoformat()},
         "method": "LESS_THAN"},
    ]}
    hdrs = {"Accept-Version": "1", "Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    counts = defaultdict(int)
    page = 1
    while True:
        resp = requests.post(
            f"https://api.cartoncloud.com/tenants/{tenant_id}/outbound-orders/search",
            params={"size": 200, "page": page}, headers=hdrs, json={"condition": cond}, timeout=60,
        )
        resp.raise_for_status()
        batch = resp.json()
        if not batch:
            break
        for o in batch:
            cust = (o.get("customer") or {}).get("name", "")
            if cust in customer_names and o.get("status") == "DISPATCHED":
                counts[cust] += 1
        if page >= int(resp.headers.get("total-pages", "1")):
            break
        page += 1
    return counts


def add_business_days(d, n):
    d = d if isinstance(d, datetime.date) else d.date()
    while n > 0:
        d += datetime.timedelta(days=1)
        if d.weekday() < 5:
            n -= 1
    return d


def load_difot_recipient_exceptions(wb):
    exceptions = []
    if "Exceptions" not in wb.sheetnames:
        return exceptions
    ws = wb["Exceptions"]
    in_table = False
    for row in ws.iter_rows(min_row=1, values_only=True):
        cell0 = str(row[0] or "")
        if "Table 4" in cell0 and "DIFOT" in cell0:
            in_table = True
            continue
        if not in_table:
            continue
        # Stop at Table 5
        if "Table 5" in cell0:
            break
        if row[0] is None and row[1] is None:
            continue
        if str(row[0] or "").strip() in ("Match Field", ""):
            continue
        match_field  = str(row[0] or "").strip()
        contains     = str(row[1] or "").strip().lower()
        action       = str(row[2] or "").strip()
        attribute_to = str(row[5] or "").strip() or None
        if attribute_to and attribute_to.lower() in ("all", "any"):
            attribute_to = None
        if not match_field or not contains or not action:
            continue
        exceptions.append({
            "match_field":  match_field,
            "contains":     contains,
            "action":       action,
            "attribute_to": attribute_to,
        })
    return exceptions


def apply_recipient_exception(row, exceptions):
    shipto = str(row[COL_SHIPTO] or "").lower()
    for exc in exceptions:
        if exc["match_field"] == "ShipTo" and exc["contains"] in shipto:
            return exc
    return None


def is_south_island(postcode):
    try:
        return int(str(postcode).strip()) >= 7000
    except (TypeError, ValueError):
        return None


def difot_target_date(pickup_dt, si, rural=False, dg=False):
    base_days = 3 if si else 1
    extra = (2 if rural else 0) + (2 if dg else 0)
    pickup_date = pickup_dt.date() if hasattr(pickup_dt, "date") else pickup_dt
    return add_business_days(pickup_date, base_days + extra)


# -- Main -----------------------------------------------------------------------

def main():
    if len(sys.argv) < 2:
        print("Usage: python extract_difot.py YYYY-MM [DeliveryPerformanceReport.xlsx]")
        sys.exit(1)

    month_str = sys.argv[1]
    year, month = int(month_str[:4]), int(month_str[5:7])
    month_start = datetime.date(year, month, 1)
    month_end   = datetime.date(year + (month // 12), (month % 12) + 1, 1)

    if len(sys.argv) >= 3:
        csv_path = sys.argv[2]
    else:
        candidates = sorted(glob.glob("DeliveryPerformanceReport*.xlsx"))
        if not candidates:
            print("No DeliveryPerformanceReport*.xlsx found in current directory.")
            sys.exit(1)
        csv_path = candidates[-1]
    print(f"Using CSV: {csv_path}")

    wb_ss = openpyxl.load_workbook(csv_path, data_only=True, read_only=True)
    ws_ss = wb_ss["Data Source"]
    all_rows = list(ws_ss.iter_rows(min_row=2, max_row=ws_ss.max_row, values_only=True))
    wb_ss.close()

    rural_postcodes   = load_rural_postcode_set()
    rural_street_lookup = load_rural_street_lookup()
    print(f"Loaded {len(rural_postcodes)} rural postcodes (NZP directory) + {len(rural_street_lookup)} rural streets (NZC zone file)")

    wb_exc = openpyxl.load_workbook(WORKBOOK)
    recipient_exceptions = load_difot_recipient_exceptions(wb_exc)
    print(f"Loaded {len(recipient_exceptions)} DIFOT recipient exception rule(s)")
    print(f"International carrier SLA codes: {list(INTL_CARRIER_DAYS.keys())}")

    seen_ourref    = set()
    candidate_rows = []  # (row, force_met, attribute_to, intl_days)
    # intl_days = None for NZ domestic (postcode-based SLA)
    # intl_days = N    for international carrier (fixed N business days)

    for r in all_rows:
        if r[COL_ORDER_SOURCE] != "CartonCloud":
            continue

        order_date = r[COL_ORDER_DATE]
        if not order_date:
            continue
        od = order_date.date() if hasattr(order_date, "date") else order_date
        if not (month_start <= od < month_end):
            continue

        country = str(r[COL_COUNTRY] or "").strip()
        code    = str(r[COL_CODE]    or "").strip()
        is_nz   = country in NZ_COUNTRIES or country.lower() == "new zealand"

        # Determine SLA type
        intl_days = None
        if not is_nz:
            intl_days = INTL_CARRIER_DAYS.get(code)
            if intl_days is None:
                continue  # Non-NZ, non-configured carrier -- skip

        ourref = r[COL_OURREF]

        # Check DC recipient exceptions (NZ only)
        if is_nz:
            exc = apply_recipient_exception(r, recipient_exceptions)
            if exc:
                if exc["action"] == "Exclude":
                    seen_ourref.add(ourref)
                    continue
                if exc["action"] == "Count as met":
                    if ourref not in seen_ourref:
                        seen_ourref.add(ourref)
                        candidate_rows.append((r, True, exc.get("attribute_to"), None))
                    continue

        # Standard row -- needs TheirRef, Delivered status, pickup + delivery dates
        if not r[COL_THEIRREF]:
            continue
        if r[COL_STATUS] != "Delivered":
            continue
        if not r[COL_PICKUP_DATE] or not r[COL_DELIVERED_DATE]:
            continue
        if ourref not in seen_ourref:
            seen_ourref.add(ourref)
            candidate_rows.append((r, False, None, intl_days))

    exc_count  = sum(1 for _, force, _, _  in candidate_rows if force)
    intl_count = sum(1 for _, force, _, d  in candidate_rows if not force and d is not None)
    std_count  = len(candidate_rows) - exc_count - intl_count
    print(f"Month {month_str}: {std_count} NZ standard + {intl_count} international FedEx + {exc_count} DC exceptions")

    if not candidate_rows:
        print("Nothing to process.")
        return

    creds = load_credentials()
    token = get_cc_token(creds)
    tenant_id = creds["CARTONCLOUD_TENANT_ID"]
    print("Fetching CartonCloud customer map...")
    cc_map = fetch_cc_customer_map(token, tenant_id, month_start, month_end)
    print(f"  {len(cc_map)} CC orders mapped")

    stats                = defaultdict(lambda: [0, 0])                        # customer -> [on_time, total]
    carrier_stats        = defaultdict(lambda: [0, 0, 0, 0])                  # carrier  -> [on_time, total, days_total, days_count]
    cust_carrier_stats   = defaultdict(lambda: defaultdict(lambda: [0, 0, 0, 0]))  # customer -> carrier -> [on_time, total, days_total, days_count]
    dc_exceptions_log    = []  # (ourref, customer, month_str) - logged to Exceptions sheet
    skipped = 0

    for r, force_met, attribute_to, intl_days in candidate_rows:
        carrier = str(r[COL_CARRIER] or "Unknown")

        if force_met:
            customer = attribute_to or cc_map.get(str(r[COL_THEIRREF] or ""))
            if not customer or customer in TEST_ACCOUNT_NAMES or customer in FREIGHT_ONLY_CUSTOMERS:
                skipped += 1
                continue
            stats[customer][1] += 1
            stats[customer][0] += 1
            carrier_stats[carrier][1] += 1
            carrier_stats[carrier][0] += 1
            cust_carrier_stats[customer][carrier][1] += 1
            cust_carrier_stats[customer][carrier][0] += 1
            # DC exceptions excluded from avg delivery days (pre-arranged bookings, no standard tracking)
            dc_exceptions_log.append((r[COL_OURREF], customer, month_str))
            continue

        their_ref = str(r[COL_THEIRREF])
        customer  = cc_map.get(their_ref)
        if not customer or customer in TEST_ACCOUNT_NAMES or customer in FREIGHT_ONLY_CUSTOMERS:
            skipped += 1
            continue

        pickup_dt    = r[COL_PICKUP_DATE]
        delivered_dt = r[COL_DELIVERED_DATE]
        pickup_date  = pickup_dt.date() if hasattr(pickup_dt, "date") else pickup_dt

        if intl_days is not None:
            target_date = add_business_days(pickup_date, intl_days)
        else:
            si    = is_south_island(r[COL_POSTCODE])
            rural = is_rural_address(r[COL_POSTCODE], r[COL_STREET] or "", r[COL_SUBURB] or "", r[COL_CITY] or "", rural_postcodes, rural_street_lookup)
            if si is None:
                skipped += 1
                continue
            target_date = difot_target_date(pickup_dt, si, rural=rural)

        delivered_date = delivered_dt.date() if hasattr(delivered_dt, "date") else delivered_dt
        on_time        = delivered_date <= target_date
        delivery_days  = (delivered_date - pickup_date).days

        stats[customer][1] += 1
        if on_time:
            stats[customer][0] += 1
        carrier_stats[carrier][1] += 1
        if on_time:
            carrier_stats[carrier][0] += 1
        carrier_stats[carrier][2] += delivery_days
        carrier_stats[carrier][3] += 1
        cust_carrier_stats[customer][carrier][1] += 1
        if on_time:
            cust_carrier_stats[customer][carrier][0] += 1
        cust_carrier_stats[customer][carrier][2] += delivery_days
        cust_carrier_stats[customer][carrier][3] += 1

    if skipped:
        print(f"  {skipped} rows skipped (no CC match, unknown postcode, or missing TheirRef)")

    print("Fetching freight-only customer order counts (Cargo Plus, counted as met)...")
    freight_counts = fetch_freight_only_counts(token, tenant_id, month_start, month_end, FREIGHT_ONLY_CUSTOMERS)
    for cust, count in freight_counts.items():
        if count <= 0:
            continue
        stats[cust][1] += count
        stats[cust][0] += count
        carrier_stats[FREIGHT_ONLY_CARRIER][1] += count
        carrier_stats[FREIGHT_ONLY_CARRIER][0] += count
        cust_carrier_stats[cust][FREIGHT_ONLY_CARRIER][1] += count
        cust_carrier_stats[cust][FREIGHT_ONLY_CARRIER][0] += count
        print(f"  {cust}: {count} dispatched orders -> counted as met via {FREIGHT_ONLY_CARRIER}")

    print(f"\n{month_str} Freight DIFOT results:")
    for cust in sorted(stats):
        n, d = stats[cust]
        print(f"  {cust}: {n}/{d} ({n/d*100:.1f}%)")

    print(f"\n{month_str} Carrier breakdown:")
    for c in sorted(carrier_stats):
        n, d, dt, dc = carrier_stats[c]
        avg = f", {dt/dc:.1f}d avg" if dc else ""
        print(f"  {c}: {n}/{d} ({n/d*100:.1f}%){avg}")

    if not stats:
        print("No results to write.")
        return

    # Write/update carrier stats JSON
    existing_carriers = {}
    if os.path.exists(CARRIER_FILE):
        with open(CARRIER_FILE) as f:
            existing_carriers = json.load(f)
    month_data = {
        "_total": {c: {"num": n, "den": d, "days_total": dt, "days_count": dc}
                   for c, (n, d, dt, dc) in carrier_stats.items()},
    }
    for cust, cstats in cust_carrier_stats.items():
        month_data[cust] = {c: {"num": n, "den": d, "days_total": dt, "days_count": dc}
                            for c, (n, d, dt, dc) in cstats.items()}
    existing_carriers[month_str] = month_data
    with open(CARRIER_FILE, "w") as f:
        json.dump(existing_carriers, f, indent=2)
    print(f"Updated {CARRIER_FILE}")

    wb = openpyxl.load_workbook(WORKBOOK)
    ws_sched  = wb["KPI Schedule"]
    difot_target = 0.95
    for row in ws_sched.iter_rows(min_row=5, max_row=ws_sched.max_row, values_only=True):
        if row[0] and "DIFOT" in str(row[0]):
            if row[1] is not None:
                difot_target = float(row[1])
            break
    print(f"\nDIFOT target: {difot_target*100:.1f}%")

    ws_data = wb["KPI_Data"]
    to_delete = []
    for row_idx in range(4, ws_data.max_row + 1):
        cm  = ws_data.cell(row=row_idx, column=1).value
        ck  = ws_data.cell(row=row_idx, column=3).value
        if ck == KPI_NAME and cm is not None:
            ms = cm.strftime("%Y-%m") if hasattr(cm, "strftime") else str(cm)[:7]
            if ms == month_str:
                to_delete.append(row_idx)
    for row_idx in reversed(to_delete):
        ws_data.delete_rows(row_idx)
    if to_delete:
        print(f"Removed {len(to_delete)} existing {month_str} DIFOT rows")

    month_date = datetime.date(year, month, 1)
    note_bad   = "Below target - investigate root cause and add to next month's action plan"
    for customer in sorted(stats):
        num, den = stats[customer]
        ri = ws_data.max_row + 1
        ws_data.cell(row=ri, column=1, value=month_date)
        ws_data.cell(row=ri, column=2, value=customer)
        ws_data.cell(row=ri, column=3, value=KPI_NAME)
        ws_data.cell(row=ri, column=4, value=None)
        ws_data.cell(row=ri, column=5, value=difot_target)
        ws_data.cell(row=ri, column=6, value=num)
        ws_data.cell(row=ri, column=7, value=den)
        ws_data.cell(row=ri, column=8, value=f"=F{ri}/G{ri}")
        ws_data.cell(row=ri, column=9, value=f'=IF(H{ri}>=E{ri},"Yes","No")')
        actual = num / den if den else 0
        ws_data.cell(row=ri, column=10, value=note_bad if actual < difot_target else None)

    # Log DC exceptions to Order-Level Exceptions section in Exceptions sheet
    if dc_exceptions_log:
        ws_exc = wb["Exceptions"]
        note_row = None
        for row_idx in range(19, ws_exc.max_row + 1):
            val = ws_exc.cell(row=row_idx, column=1).value
            if val and str(val).startswith("Action must be either"):
                note_row = row_idx
                break
        if note_row is None:
            note_row = ws_exc.max_row + 1
        rows_to_delete = []
        for row_idx in range(19, note_row):
            m_val = ws_exc.cell(row=row_idx, column=4).value
            kpi_val = ws_exc.cell(row=row_idx, column=5).value
            if m_val == month_str and kpi_val == KPI_NAME:
                rows_to_delete.append(row_idx)
        for row_idx in reversed(rows_to_delete):
            ws_exc.delete_rows(row_idx)
            note_row -= 1
        if rows_to_delete:
            print(f"Removed {len(rows_to_delete)} existing DC exception log entries for {month_str}")
        today = datetime.date.today()
        for ourref, customer, mon in dc_exceptions_log:
            ws_exc.insert_rows(note_row)
            ws_exc.cell(row=note_row, column=1, value=today)
            ws_exc.cell(row=note_row, column=2, value=customer)
            ws_exc.cell(row=note_row, column=3, value=ourref)
            ws_exc.cell(row=note_row, column=4, value=mon)
            ws_exc.cell(row=note_row, column=5, value=KPI_NAME)
            ws_exc.cell(row=note_row, column=6, value="Count as met")
            ws_exc.cell(row=note_row, column=7, value="DC recipient exception - pre-arranged booking, excluded from avg delivery days")
            note_row += 1
        print(f"Logged {len(dc_exceptions_log)} DC exception(s) to Exceptions sheet")

    wb.save(WORKBOOK)
    print(f"\nAppended {len(stats)} rows to KPI_Data in {WORKBOOK}")


if __name__ == "__main__":
    main()

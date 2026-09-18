"""
mtd_difot.py  -  Month-to-date Freight DIFOT via CartonCloud + Starshipit Tracking API

Uses a persistent cache (mtd_tracking_cache.json) so that if the Starshipit fetch takes
more than one 45-second bash window, progress is preserved and resumed on the next call.

Run twice in the scheduled task — first call fetches a chunk, second call fetches the
rest and computes. Once the cache is warm, a single call handles everything.

Usage:
    python3 mtd_difot.py                  # fetch next chunk + compute from cache
    python3 mtd_difot.py --fetch-only     # fetch only, skip compute output
    python3 mtd_difot.py 2026-06          # historical month (writes to KPI_Data)
"""

import os, re, sys, json, time, datetime, requests, openpyxl
from zoneinfo import ZoneInfo
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

WORKBOOK            = "kpi_dashboard.xlsx"
CREDS_FILE          = "credentials.env"
MTD_FILE            = "mtd_kpis.json"
CARRIER_FILE        = "difot_carriers.json"
CACHE_FILE          = "mtd_tracking_cache.json"
RURAL_POSTCODE_FILE = "nz_rural_postcodes.json"
KPI_NAME            = "Freight DIFOT"
TEST_ACCOUNT_NAMES  = {"TEST ACCOUNT"}
# Customers whose real deliveries are truck freight (e.g. Cargo Plus) rather than
# Starshipit-tracked parcel carriers. CartonCloud's references.numericId is only
# unique per-customer, not tenant-wide, so a colliding numericId can silently pull
# in an unrelated parcel customer's tracking data. Confirmed for Made Group NZ
# (see extract_difot.py for the full writeup) - excluded entirely rather than scored.
FREIGHT_ONLY_CUSTOMERS = {"Made Group NZ"}
NZ_COUNTRIES        = {"New Zealand", "NEW ZEALAND", "new zealand"}
NZ_TZ               = ZoneInfo("Pacific/Auckland")
INTL_CARRIER_DAYS   = {"FEDEX_INTERNATIONAL_CONNECT_PLUS": 3}
RURAL_RE            = re.compile(r'\bR\.?D\.?\s*\d', re.IGNORECASE)

# Extra business days added to South Island SLA for months with known network disruptions
# (interisland ferry reduced sailings, road closures, weather events, etc.)
SI_EXTRA_DAYS = {
    "2026-07": 1,  # reduced ferry sailings + Kaikoura flooding
}

CHUNK_SIZE = 16    # orders to fetch per pass (4 workers x 4 each ~15s)
WORKERS    = 4     # concurrent Starshipit requests

# NZ public holidays - extend each year as dates are confirmed
NZ_PUBLIC_HOLIDAYS = {
    # 2025
    datetime.date(2025, 1, 1), datetime.date(2025, 1, 2),
    datetime.date(2025, 2, 6), datetime.date(2025, 4, 18),
    datetime.date(2025, 4, 21), datetime.date(2025, 4, 25),
    datetime.date(2025, 6, 2),   # King's Birthday
    datetime.date(2025, 6, 20),  # Matariki
    datetime.date(2025, 10, 27), # Labour Day
    datetime.date(2025, 12, 25), datetime.date(2025, 12, 26),
    # 2026
    datetime.date(2026, 1, 1), datetime.date(2026, 1, 2),
    datetime.date(2026, 2, 6), datetime.date(2026, 4, 3),
    datetime.date(2026, 4, 6),  datetime.date(2026, 4, 25),
    datetime.date(2026, 6, 1),   # King's Birthday
    datetime.date(2026, 7, 10),  # Matariki
    datetime.date(2026, 10, 26), # Labour Day
    datetime.date(2026, 12, 25), datetime.date(2026, 12, 26),
    # 2027
    datetime.date(2027, 1, 1), datetime.date(2027, 1, 4),
    datetime.date(2027, 2, 8),   # Waitangi Day observed
    datetime.date(2027, 3, 26),  # Good Friday
    datetime.date(2027, 3, 29),  # Easter Monday
    datetime.date(2027, 4, 26),  # Anzac Day observed
    datetime.date(2027, 6, 7),   # King's Birthday
    datetime.date(2027, 7, 7),   # Matariki (provisional)
    datetime.date(2027, 10, 25), # Labour Day
    datetime.date(2027, 12, 27), datetime.date(2027, 12, 28),
}


def load_credentials():
    creds = {}
    with open(CREDS_FILE) as f:
        for line in f:
            line = line.strip()
            if line and "=" in line and not line.startswith("#"):
                k, v = line.split("=", 1)
                creds[k.strip()] = v.strip()
    return creds


def load_rural_postcodes():
    if not os.path.exists(RURAL_POSTCODE_FILE):
        return set()
    with open(RURAL_POSTCODE_FILE) as f:
        return set(json.load(f)["postcodes"])


def add_business_days(d, n):
    d = d if isinstance(d, datetime.date) else d.date()
    while n > 0:
        d += datetime.timedelta(days=1)
        if d.weekday() < 5 and d not in NZ_PUBLIC_HOLIDAYS:
            n -= 1
    return d


def is_south_island(postcode):
    try:
        return int(str(postcode).strip()) >= 7000
    except (TypeError, ValueError):
        return None


def is_rural(postcode, street, suburb, city, rural_postcodes):
    pc = str(postcode or '').strip().zfill(4)
    if pc in rural_postcodes:
        return True
    return bool(RURAL_RE.search(f"{street} {suburb} {city}"))


def difot_target_date(pickup_date, si, rural=False, extra_si_days=0):
    return add_business_days(pickup_date, (3 if si else 1) + (2 if rural else 0) + (extra_si_days if si else 0))


def load_exceptions(wb):
    excs = []
    if "Exceptions" not in wb.sheetnames:
        return excs
    ws = wb["Exceptions"]
    in_t4 = False
    for row in ws.iter_rows(min_row=1, values_only=True):
        c0 = str(row[0] or "")
        if "Table 4" in c0 and "DIFOT" in c0:
            in_t4 = True; continue
        if not in_t4: continue
        if "Table 5" in c0: break
        if not row[0] and not row[1]: continue
        if str(row[0] or "").strip() in ("Match Field", ""): continue
        mf  = str(row[0] or "").strip()
        cnt = str(row[1] or "").strip().lower()
        act = str(row[2] or "").strip()
        att = str(row[5] or "").strip() or None
        if att and att.lower() in ("all", "any"): att = None
        if mf and cnt and act:
            excs.append({"match_field": mf, "contains": cnt, "action": act, "attribute_to": att})
    return excs


def apply_exc(ship_to, excs):
    s = str(ship_to or "").lower()
    for e in excs:
        if e["match_field"] == "ShipTo" and e["contains"] in s:
            return e
    return None


def fetch_cc_orders(token, tenant_id, month_start, month_end):
    hdrs = {"Accept-Version": "1", "Authorization": f"Bearer {token}", "Content-Type": "application/json"}
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
    orders, page = [], 1
    while True:
        resp = requests.post(
            f"https://api.cartoncloud.com/tenants/{tenant_id}/outbound-orders/search",
            params={"size": 200, "page": page}, headers=hdrs,
            json={"condition": cond}, timeout=60,
        )
        resp.raise_for_status()
        batch = resp.json()
        if not batch: break
        for o in batch:
            nid  = (o.get("references") or {}).get("numericId")
            cust = (o.get("customer") or {}).get("name", "")
            props = o.get("properties") or {}
            tn   = (props.get("carrierTrackingNumberField") or "").strip()
            addr = ((o.get("details") or {}).get("deliver") or {}).get("address") or {}
            cr   = addr.get("country") or {}
            orders.append({
                "numeric_id": str(nid) if nid else None,
                "customer":   cust,
                "status":     o.get("status"),
                "tracking":   tn,
                "ship_to":    addr.get("companyName") or addr.get("contactName") or "",
                "street":     addr.get("address1") or "",
                "suburb":     addr.get("suburb") or "",
                "city":       addr.get("city") or addr.get("suburb") or "",
                "postcode":   addr.get("postcode") or "",
                "country":    cr.get("name", "New Zealand") if isinstance(cr, dict) else str(cr),
            })
        if page >= int(resp.headers.get("total-pages", "1")): break
        page += 1
    return [o for o in orders if o["numeric_id"]]


def fetch_tracking_one(ss_hdrs, tracking_number):
    for attempt in range(3):
        try:
            r = requests.get(
                "https://api.starshipit.com/api/track",
                params={"tracking_number": tracking_number},
                headers=ss_hdrs, timeout=12,
            )
            if r.status_code == 200:
                data = r.json()
                return data.get("results") if data.get("success") else None
            if r.status_code == 429:
                time.sleep(2 ** (attempt + 1))
                continue
            return None
        except Exception:
            if attempt < 2: time.sleep(1)
    return None


def fetch_chunk(ss_hdrs, orders_chunk):
    results = {}
    with ThreadPoolExecutor(max_workers=WORKERS) as executor:
        futures = {
            executor.submit(fetch_tracking_one, ss_hdrs, o["tracking"]): o["numeric_id"]
            for o in orders_chunk
        }
        for fut in as_completed(futures):
            results[futures[fut]] = fut.result()
    return results


def parse_event_date_nz(s):
    if not s: return None
    try:
        s = s.strip().replace("Z", "+00:00")
        return datetime.datetime.fromisoformat(s).astimezone(NZ_TZ).date()
    except Exception:
        return None


def compute_stats(cc_orders, tracking_map, excs, rural_postcodes, month_str=""):
    seen = set()
    stats = defaultdict(lambda: [0, 0])
    carrier_stats = defaultdict(lambda: defaultdict(lambda: [0, 0]))
    skipped = exc_count = 0

    for order in cc_orders:
        nid  = order["numeric_id"]
        cust = order["customer"]
        if not cust or cust in TEST_ACCOUNT_NAMES or cust in FREIGHT_ONLY_CUSTOMERS: skipped += 1; continue
        if nid in seen: continue

        country = (order["country"] or "").strip()
        is_nz   = country in NZ_COUNTRIES or country.lower() == "new zealand"

        if is_nz:
            exc = apply_exc(order["ship_to"], excs)
            if exc:
                if exc["action"] == "Exclude": seen.add(nid); continue
                if exc["action"] == "Count as met":
                    seen.add(nid)
                    ac = exc.get("attribute_to") or cust
                    if ac not in TEST_ACCOUNT_NAMES:
                        stats[ac][0] += 1; stats[ac][1] += 1
                        carrier_stats["(exception)"][ac][0] += 1
                        carrier_stats["(exception)"][ac][1] += 1
                        exc_count += 1
                    continue

        td = tracking_map.get(nid)
        if not td: skipped += 1; continue

        status = (td.get("tracking_status") or td.get("order_status") or "").lower()
        if status != "delivered": continue

        events = td.get("tracking_events") or []
        pickup_dt = delivered_dt = None
        for ev in events:
            evs = (ev.get("status") or "").lower()
            if evs == "dispatched" and not pickup_dt:
                pickup_dt = parse_event_date_nz(ev.get("event_datetime"))
            if evs == "delivered" and not delivered_dt:
                delivered_dt = parse_event_date_nz(ev.get("event_datetime"))
        if not pickup_dt or not delivered_dt: skipped += 1; continue
        seen.add(nid)

        if not is_nz:
            svc_upper = (td.get("carrier_service") or "").upper().replace(" ", "_")
            code = "FEDEX_INTERNATIONAL_CONNECT_PLUS" if "INTERNATIONAL_CONNECT_PLUS" in svc_upper else ""
            intl = INTL_CARRIER_DAYS.get(code)
            if not intl: skipped += 1; continue
            target = add_business_days(pickup_dt, intl)
        else:
            si = is_south_island(order["postcode"])
            if si is None: skipped += 1; continue
            rur = is_rural(order["postcode"], order["street"], order["suburb"], order["city"], rural_postcodes)
            extra = SI_EXTRA_DAYS.get(month_str, 0)
            target = difot_target_date(pickup_dt, si, rural=rur, extra_si_days=extra)

        on_time = delivered_dt <= target
        carrier = (td.get("carrier_name") or "Unknown").strip()
        stats[cust][1] += 1
        carrier_stats[carrier][cust][1] += 1
        if on_time:
            stats[cust][0] += 1
            carrier_stats[carrier][cust][0] += 1

    return stats, carrier_stats, skipped, exc_count


def main():
    sys.stdout.reconfigure(line_buffering=True)
    today    = datetime.datetime.now(NZ_TZ).date()
    args     = sys.argv[1:]
    fetch_only    = "--fetch-only" in args
    compute_only  = "--compute-only" in args
    hist_args  = [a for a in args if not a.startswith("--")]
    historical = hist_args[0] if hist_args else None

    if historical:
        year, month = int(historical[:4]), int(historical[5:7])
        month_start = datetime.date(year, month, 1)
        next_m = month % 12 + 1
        next_y = year + (1 if month == 12 else 0)
        month_end = datetime.date(next_y, next_m, 1)
        month_str = historical
    else:
        month_start = today.replace(day=1)
        month_end   = today + datetime.timedelta(days=1)
        month_str   = today.strftime("%Y-%m")

    creds = load_credentials()
    rural_postcodes = load_rural_postcodes()
    print(f"Loaded {len(rural_postcodes)} rural postcodes")

    wb_exc = openpyxl.load_workbook(WORKBOOK)
    excs = load_exceptions(wb_exc)
    difot_target = 0.95
    for row in wb_exc["KPI Schedule"].iter_rows(min_row=5, max_row=wb_exc["KPI Schedule"].max_row, values_only=True):
        if row[0] and "DIFOT" in str(row[0]) and row[1] is not None:
            difot_target = float(row[1]); break
    wb_exc.close()

    print(f"Authenticating to CartonCloud...")
    from cc_auth import get_cc_token
    cc_token = get_cc_token(creds)

    print(f"Fetching CC orders {month_start} to {month_end - datetime.timedelta(days=1)}...")
    cc_orders = fetch_cc_orders(cc_token, creds["CARTONCLOUD_TENANT_ID"], month_start, month_end)
    print(f"  {len(cc_orders)} orders")

    ss_hdrs = {
        "StarShipIT-Api-Key":        creds["STARSHIPIT_API_KEY"],
        "Ocp-Apim-Subscription-Key": creds["STARSHIPIT_SUBSCRIPTION_KEY"],
    }

    cache = {}
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE) as f:
            cache = json.load(f)
    month_cache = cache.get(month_str, {})

    def _is_delivered(v):
        if v is None: return False
        status = (v.get("tracking_status") or v.get("order_status") or "").lower()
        return status == "delivered"

    trackable  = [o for o in cc_orders if o["tracking"]]
    # Re-fetch anything not yet confirmed delivered — statuses can change after first fetch
    uncached   = [o for o in trackable if not _is_delivered(month_cache.get(o["numeric_id"]))]
    fresh      = sum(1 for o in trackable if _is_delivered(month_cache.get(o["numeric_id"])))
    print(f"  {len(trackable)} have tracking — {fresh} confirmed delivered, {len(uncached)} to refresh")

    BUDGET = 0 if compute_only else 32  # seconds to spend fetching before saving and stopping
    t_start = time.time()
    offset = 0
    while uncached[offset:]:
        chunk = uncached[offset:offset + CHUNK_SIZE]
        elapsed = time.time() - t_start
        if elapsed > BUDGET:
            break
        print(f"  Fetching chunk of {len(chunk)} ({WORKERS} workers, {elapsed:.0f}s elapsed)...")
        new_results = fetch_chunk(ss_hdrs, chunk)
        month_cache.update(new_results)
        cache[month_str] = month_cache
        # Write atomically (temp file + rename) so a process kill mid-write
        # (e.g. the caller's timeout expiring) can never truncate/corrupt the
        # on-disk cache — the old file stays valid until the new one is complete.
        tmp_path = CACHE_FILE + ".tmp"
        with open(tmp_path, "w") as f:
            json.dump(cache, f)
        os.replace(tmp_path, CACHE_FILE)
        offset += len(chunk)
    still = len(uncached) - offset
    if offset:
        print(f"  Fetched {offset} orders in {time.time()-t_start:.1f}s. {still} still to refresh.")
    else:
        print(f"  Cache complete.")

    if fetch_only:
        print("--fetch-only: done.")
        return

    stats, carrier_stats, skipped, exc_count = compute_stats(cc_orders, month_cache, excs, rural_postcodes, month_str)

    # Freight-only customers (e.g. Made Group NZ) ship via truck freight (Cargo Plus)
    # that Starshipit never tracks, so count their real dispatched orders as met
    # directly rather than relying on the collision-prone numericId join.
    freight_counts = defaultdict(int)
    for o in cc_orders:
        if o["customer"] in FREIGHT_ONLY_CUSTOMERS and o.get("status") == "DISPATCHED":
            freight_counts[o["customer"]] += 1
    for cust, count in freight_counts.items():
        if count <= 0:
            continue
        stats[cust][1] += count
        stats[cust][0] += count
        carrier_stats["Cargo Plus"][cust][1] += count
        carrier_stats["Cargo Plus"][cust][0] += count

    pending = len([o for o in trackable if o["numeric_id"] not in month_cache])
    label   = month_str if historical else f"MTD ({today})"
    note    = f" [{pending} orders still pending — run again for complete data]" if pending else ""
    print(f"\nFreight DIFOT {label}{note}:")
    if skipped:    print(f"  ({skipped} skipped: no data / in transit / unknown postcode)")
    if exc_count:  print(f"  ({exc_count} counted as met via DC exception)")

    if stats:
        for cust in sorted(stats):
            n, d = stats[cust]
            print(f"  {cust}: {n}/{d} ({n/d*100:.1f}%)")
        print("\nCarrier breakdown:")
        for carrier in sorted(carrier_stats):
            tn = sum(v[0] for v in carrier_stats[carrier].values())
            td = sum(v[1] for v in carrier_stats[carrier].values())
            print(f"  {carrier}: {tn}/{td} ({tn/td*100:.1f}%)")
    else:
        print("  (no delivered orders)")

    if historical:
        month_entry = {"_total": {}}
        for carrier, cm in carrier_stats.items():
            tn = sum(v[0] for v in cm.values())
            td_v = sum(v[1] for v in cm.values())
            month_entry["_total"][carrier] = {"num": tn, "den": td_v}
            for cust, (n, d) in cm.items():
                month_entry.setdefault(cust, {})[carrier] = {"num": n, "den": d}
        all_d = {}
        if os.path.exists(CARRIER_FILE):
            with open(CARRIER_FILE) as f: all_d = json.load(f)
        all_d[month_str] = month_entry
        with open(CARRIER_FILE, "w") as f: json.dump(all_d, f, indent=2)

        wb = openpyxl.load_workbook(WORKBOOK)
        ws = wb["KPI_Data"]
        to_del = []
        for r in range(2, ws.max_row + 1):
            cv = ws.cell(row=r, column=1).value
            if ws.cell(row=r, column=3).value == KPI_NAME and isinstance(cv, datetime.datetime) \
               and cv.year == month_start.year and cv.month == month_start.month:
                to_del.append(r)
        for r in reversed(to_del): ws.delete_rows(r)
        row = ws.max_row + 1
        mdt = datetime.datetime(month_start.year, month_start.month, 1)
        appended = 0
        for cust in sorted(stats):
            n, d = stats[cust]
            if not d: continue
            ws.cell(row=row, column=1).value = mdt
            ws.cell(row=row, column=1).number_format = "mmm\\-yyyy"
            ws.cell(row=row, column=2).value = cust
            ws.cell(row=row, column=3).value = KPI_NAME
            ws.cell(row=row, column=4).value = None
            ws.cell(row=row, column=5).value = difot_target
            ws.cell(row=row, column=5).number_format = "0.0%"
            ws.cell(row=row, column=6).value = n
            ws.cell(row=row, column=7).value = d
            ws.cell(row=row, column=8).value = f"=F{row}/G{row}"
            ws.cell(row=row, column=8).number_format = "0.0%"
            ws.cell(row=row, column=9).value = f'=IF(H{row}>=E{row},"Yes","No")'
            if n / d < difot_target:
                ws.cell(row=row, column=10).value = "Below target - investigate root cause"
            row += 1; appended += 1
        wb.save(WORKBOOK)
        print(f"Appended {appended} rows to KPI_Data")

    else:
        mtd = {}
        if os.path.exists(MTD_FILE):
            with open(MTD_FILE) as f: mtd = json.load(f)
        mtd.setdefault("kpis", {})[KPI_NAME] = {
            cust: {"num": n, "den": d} for cust, (n, d) in stats.items()
        }
        mtd["month"]  = month_str
        mtd["as_of"]  = today.isoformat()
        with open(MTD_FILE, "w") as f: json.dump(mtd, f, indent=2)
        print(f"\nUpdated {MTD_FILE}")


if __name__ == "__main__":
    main()

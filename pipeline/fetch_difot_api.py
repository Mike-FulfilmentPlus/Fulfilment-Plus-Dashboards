"""
fetch_difot_api.py - Build a Delivery Performance "Data Source" workbook from the
documented Starshipit Orders/Tracking API, as a replacement for the (now broken)
undocumented Reports API used by the old run_difot.py.

Produces an xlsx with a "Data Source" sheet whose column layout exactly matches
what extract_difot.py expects (0-based indices, see COL_* constants there), so
extract_difot.py needs NO changes.
"""
import sys, os, time, json, datetime
import requests
import openpyxl
from zoneinfo import ZoneInfo
from concurrent.futures import ThreadPoolExecutor, as_completed

CREDS_FILE = "credentials.env"
NZ_TZ = ZoneInfo("Pacific/Auckland")
UTC = ZoneInfo("UTC")

# Must match extract_difot.py's COL_* layout exactly
N_COLS = 30
COL_OURREF, COL_THEIRREF, COL_SHIPTO, COL_STREET, COL_SUBURB, COL_CITY = 1, 2, 3, 4, 5, 6
COL_POSTCODE, COL_COUNTRY, COL_CARRIER, COL_CODE = 8, 9, 12, 14
COL_ORDER_SOURCE, COL_ORDER_DATE, COL_PICKUP_DATE, COL_DELIVERED_DATE, COL_STATUS = 21, 23, 26, 28, 29


def load_credentials(path=CREDS_FILE):
    creds = {}
    with open(path) as f:
        for line in f:
            line = line.strip()
            if line and "=" in line and not line.startswith("#"):
                k, v = line.split("=", 1)
                creds[k.strip()] = v.strip()
    return creds


def ss_headers(creds):
    return {
        "StarShipIT-Api-Key":        creds["STARSHIPIT_API_KEY"],
        "Ocp-Apim-Subscription-Key": creds["STARSHIPIT_SUBSCRIPTION_KEY"],
    }


import re
_FRAC_RE = re.compile(r"(\.\d{1,6})\d*")

def _parse_dt(s):
    """Parse a Starshipit timestamp, truncating >6-digit fractional seconds
    (Python rejects them) and converting Z/UTC-marked timestamps to NZ local
    naive datetimes; naive strings are assumed already NZ-local."""
    if not s:
        return None
    s = _FRAC_RE.sub(r"\1", s)
    if s.endswith("Z"):
        dt = datetime.datetime.fromisoformat(s.replace("Z", "+00:00"))
        return dt.astimezone(NZ_TZ).replace(tzinfo=None)
    return datetime.datetime.fromisoformat(s)


def list_shipped_candidates(hdrs, month_start, month_end, buffer_days=3):
    """Bulk-list shipped orders via /api/orders/shipped, filtered to CartonCloud
    orders whose order_date falls in [month_start, month_end)."""
    since = datetime.datetime.combine(month_start - datetime.timedelta(days=buffer_days), datetime.time.min)
    since_str = since.strftime("%Y-%m-%dT%H:%M:%S.000Z")
    candidates = {}  # order_number -> row dict (first occurrence wins)
    page = 1
    while True:
        resp = requests.get(
            "https://api.starshipit.com/api/orders/shipped",
            params={"since_last_updated": since_str, "limit": 250, "page": page},
            headers=hdrs, timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        orders = data.get("orders", [])
        for o in orders:
            if o.get("integration_source_name") != "CartonCloud":
                continue
            od = o.get("order_date")
            if not od:
                continue
            od_date = _parse_dt(od).date()
            if not (month_start <= od_date < month_end):
                continue
            onum = o.get("order_number")
            if not onum or onum in candidates:
                continue
            candidates[onum] = o
        total_pages = data.get("total_pages", 1)
        print(f"  page {page}/{total_pages}: {len(orders)} orders ({len(candidates)} CartonCloud candidates so far)")
        if page >= total_pages:
            break
        page += 1
        time.sleep(0.2)
    return candidates


def enrich_order(hdrs, order_number, max_retries=4):
    """Fetch full order details (destination, dangerous_goods, events) for one order."""
    for attempt in range(max_retries):
        resp = requests.get(
            "https://api.starshipit.com/api/orders",
            params=[("order_number", order_number), ("include", "Destination"), ("include", "Events")],
            headers=hdrs, timeout=30,
        )
        if resp.status_code == 429:
            time.sleep(3 * (attempt + 1))
            continue
        resp.raise_for_status()
        data = resp.json()
        return data.get("order")
    raise RuntimeError(f"Failed to enrich order {order_number} after {max_retries} retries (rate limited)")


def find_event_time(order, category, status_update_method=None):
    """Return the NZ-local naive datetime of a carrier-confirmed tracking event
    (matched by its 'category', e.g. 'Delivered' / 'PickedUp'), preferring it over
    Starshipit's own internal status-change log entry (status_update_method),
    which lags the real carrier scan and is only used as a fallback."""
    events = order.get("events") or []
    carrier_evt = None
    status_evt = None
    for e in events:
        if e.get("category") == category:
            carrier_evt = e
        elif status_update_method and e.get("method") == status_update_method:
            status_evt = e
    chosen = carrier_evt or status_evt
    if not chosen:
        return None
    return _parse_dt(chosen.get("time"))


def build_row(bulk_order, full_order):
    row = [None] * N_COLS
    dest = full_order.get("destination") or {}
    row[COL_OURREF]         = full_order.get("order_number") or bulk_order.get("order_number")
    row[COL_THEIRREF]       = full_order.get("reference") or bulk_order.get("reference")
    row[COL_SHIPTO]         = dest.get("name", "")
    row[COL_STREET]         = dest.get("street", "")
    row[COL_SUBURB]         = dest.get("suburb", "")
    row[COL_CITY]           = dest.get("city", "")
    row[COL_POSTCODE]       = dest.get("post_code", "")
    row[COL_COUNTRY]        = dest.get("country", "")
    row[COL_CARRIER]        = full_order.get("carrier_name") or bulk_order.get("carrier_name", "")
    row[COL_CODE]           = full_order.get("carrier_service_code") or bulk_order.get("carrier_service_code", "")
    row[COL_ORDER_SOURCE]   = "CartonCloud"
    row[COL_ORDER_DATE]     = _parse_dt(full_order.get("order_date") or bulk_order.get("order_date"))
    # Pickup date: prefer the carrier's own "PickedUp" scan event over the bulk
    # listing's shipped_date (which reflects when Starshipit printed the label).
    pickup_date = find_event_time(full_order, "PickedUp", "Status Update: Dispatched")
    if pickup_date is None:
        pickup_date = _parse_dt(bulk_order.get("shipped_date"))
    row[COL_PICKUP_DATE]    = pickup_date
    delivered = find_event_time(full_order, "Delivered", "Status Update: Delivered")
    row[COL_DELIVERED_DATE] = delivered
    row[COL_STATUS]         = "Delivered" if delivered else (full_order.get("status") or "")
    return row


def build_workbook(month_str, out_path, sample_limit=None):
    year, month = int(month_str[:4]), int(month_str[5:7])
    month_start = datetime.date(year, month, 1)
    next_m = month % 12 + 1
    next_y = year + (1 if month == 12 else 0)
    month_end = datetime.date(next_y, next_m, 1)

    creds = load_credentials()
    hdrs = ss_headers(creds)

    print(f"Listing shipped CartonCloud orders for {month_str} ({month_start} to {month_end})...")
    candidates = list_shipped_candidates(hdrs, month_start, month_end)
    print(f"Total CartonCloud candidates in {month_str}: {len(candidates)}")

    order_numbers = list(candidates.keys())
    if sample_limit:
        order_numbers = order_numbers[:sample_limit]
        print(f"  (sample_limit active - only enriching {len(order_numbers)})")

    # Checkpoint cache so a run that gets cut off (timeout) can resume without
    # re-fetching orders already enriched.
    cache_path = out_path + ".cache.json"
    cache = {}
    if os.path.exists(cache_path):
        with open(cache_path) as f:
            cache = json.load(f)
        print(f"  loaded {len(cache)} cached enrichments from {cache_path}")

    to_fetch = [o for o in order_numbers if o not in cache]
    print(f"  {len(to_fetch)} orders still need enrichment ({len(order_numbers) - len(to_fetch)} cached)")

    errors = []
    done = 0
    save_every = 25
    with ThreadPoolExecutor(max_workers=6) as ex:
        futures = {ex.submit(enrich_order, hdrs, onum): onum for onum in to_fetch}
        for fut in as_completed(futures):
            onum = futures[fut]
            try:
                full = fut.result()
                if full:
                    cache[onum] = full
            except Exception as e:
                errors.append((onum, str(e)))
            done += 1
            if done % save_every == 0 or done == len(to_fetch):
                with open(cache_path, "w") as f:
                    json.dump(cache, f)
                print(f"  enriched {done}/{len(to_fetch)} this run ({len(cache)}/{len(order_numbers)} total cached, {len(errors)} errors)")

    if errors:
        print(f"WARNING: {len(errors)} orders failed enrichment: {errors[:10]}")

    rows = [build_row(candidates[onum], cache[onum]) for onum in order_numbers if onum in cache]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data Source"
    ws.append([f"col{i}" for i in range(N_COLS)])  # header placeholder row (extract_difot.py starts at row 2)
    for row in rows:
        ws.append(row)
    wb.save(out_path)
    print(f"Saved {len(rows)} rows to {out_path}")
    return out_path, rows


if __name__ == "__main__":
    month_str = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else f"DeliveryPerformanceReport_API_{month_str}.xlsx"
    sample_limit = int(sys.argv[3]) if len(sys.argv) > 3 else None
    build_workbook(month_str, out_path, sample_limit=sample_limit)

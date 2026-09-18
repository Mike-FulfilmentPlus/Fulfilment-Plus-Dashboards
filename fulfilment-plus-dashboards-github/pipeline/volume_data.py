"""
Shared daily order-volume cache.

Used by:
  - generate_warehouse_dashboard.py - the 12-month daily volume chart
    (aggregate across all customers)
  - mtd_kpis.py / extract_kpi_data.py - the Order Accuracy by Unit KPI's
    denominator (units dispatched per customer per month)

Single source of truth so only one script needs to hit the CartonCloud API
per day, regardless of how many scripts need volume data. Whichever script
runs first each day calls refresh_volume(), which only re-fetches the last
VOLUME_REFRESH_DAYS days from CartonCloud; everything older is read straight
from volume_cache.json, which accumulates a day at a time and never shrinks.
The other scripts that run later the same day just read the same
already-fresh cache - no extra API calls.

This module is intentionally self-contained (its own tiny CartonCloud
fetch helpers) rather than importing from generate_warehouse_dashboard.py /
mtd_kpis.py / extract_kpi_data.py, so it has no dependency on - and can't be
broken by - unrelated changes in those larger files.

Usage:
    from volume_data import refresh_volume, units_dispatched_for_month

    creds = load_credentials()
    token = get_cc_token(creds)
    refresh_volume(token, creds["CARTONCLOUD_TENANT_ID"], today)
    units = units_dispatched_for_month("Vixxen", "2026-07")
"""

import datetime
import json
import os
import requests

CREDS_FILE = "credentials.env"
VOLUME_CACHE = "volume_cache.json"
TEST_ACCOUNT_NAMES = {"TEST ACCOUNT"}

# Each run only fetches today + yesterday from the API; everything older is
# read from the accumulated cache. Kept small since this runs on every
# refresh of every script that needs volume data.
VOLUME_REFRESH_DAYS = 2


# ── Credentials / CartonCloud fetch (self-contained, deliberately duplicated
#    from the other scripts - see module docstring) ────────────────────────
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


def _cc_headers(token):
    return {
        "Accept-Version": "1",
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


def _search_all_pages(token, tenant_id, resource, condition, size=100):
    results = []
    page = 1
    url = f"https://api.cartoncloud.com/tenants/{tenant_id}/{resource}/search"
    while True:
        resp = requests.post(
            url,
            params={"size": size, "page": page},
            headers=_cc_headers(token),
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


def _date_range_condition(field_pointer, start_date, end_date):
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


# ── Order field helpers ─────────────────────────────────────────────────────
def _order_items(o):
    return o.get("items") or []


def _order_line_count(o):
    return len(_order_items(o))


def _order_unit_count(o):
    return sum(
        int(item.get("measures", {}).get("quantity") or 0)
        for item in _order_items(o)
    )


def _order_customer(o):
    return o.get("customer", {}).get("name", "")


# ── Fetch + cache ────────────────────────────────────────────────────────────
def fetch_volume_for_range(token, tenant_id, start_iso, end_iso):
    """Fetch dispatched orders for a date range and aggregate to daily
    totals, both overall and broken down by customer.

    Uses daily date chunks to avoid the CartonCloud API result cap (~2000
    per query). Each day's query is tiny regardless of total volume.
    """
    start = datetime.date.fromisoformat(start_iso)
    end = datetime.date.fromisoformat(end_iso)
    by_day = {}
    d = start
    while d < end:
        next_d = d + datetime.timedelta(days=1)
        orders = _search_all_pages(
            token, tenant_id, "outbound-orders",
            _date_range_condition("/timestamps/dispatched/time",
                                  d.isoformat(), next_d.isoformat()),
        )
        orders = [o for o in orders if _order_customer(o) not in TEST_ACCOUNT_NAMES]
        if orders:
            by_customer = {}
            for o in orders:
                cust = _order_customer(o)
                bc = by_customer.setdefault(cust, {"orders": 0, "lines": 0, "units": 0})
                bc["orders"] += 1
                bc["lines"] += _order_line_count(o)
                bc["units"] += _order_unit_count(o)
            by_day[d.isoformat()] = {
                "date":        d.isoformat(),
                "orders":      len(orders),
                "lines":       sum(bc["lines"] for bc in by_customer.values()),
                "units":       sum(bc["units"] for bc in by_customer.values()),
                "by_customer": by_customer,
            }
        d = next_d
    return by_day


def load_volume_cache(path=VOLUME_CACHE):
    """Load persisted daily volume data from disk."""
    if not os.path.exists(path):
        return {}
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def save_volume_cache(cache, path=VOLUME_CACHE):
    """Persist daily volume data to disk."""
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cache, f)


def refresh_volume(token, tenant_id, today, path=VOLUME_CACHE):
    """Return daily volume (with per-customer breakdown), using a growing
    cache file shared by every script that needs volume data.

    Each call only fetches today and yesterday from the API. Everything
    older is read from the cache, which accumulates a day at a time. Safe
    to call from multiple scripts on the same day - whichever runs first
    does the real fetch; later calls the same day see the already-updated
    cache and only re-confirm the last 2 days (cheap, no harm in repeating).

    IMPORTANT: we do NOT delete cache entries before re-fetching. If the API
    returns an empty result for a day (network blip, timeout, etc.) the old
    cached value is preserved rather than lost.

    Days from before this per-customer breakdown was added only have the
    old {date, orders, lines, units} shape (no "by_customer" key) - callers
    reading by_customer must handle that being absent for old dates.
    """
    cache = load_volume_cache(path)

    refresh_from = today - datetime.timedelta(days=VOLUME_REFRESH_DAYS - 1)
    fresh = fetch_volume_for_range(
        token, tenant_id,
        refresh_from.isoformat(),
        (today + datetime.timedelta(days=1)).isoformat(),
    )
    # Only overwrite cache entries that have real data from the API.
    # Days with 0 orders (e.g. weekends, holidays) are not stored so they
    # don't clutter the chart - the existing cache entry (if any) is kept.
    for key, val in fresh.items():
        if val.get("orders", 0) > 0:
            cache[key] = val
    save_volume_cache(cache, path)

    return sorted(cache.values(), key=lambda x: x["date"])


def units_dispatched_for_month(customer, month_str, cache=None, path=VOLUME_CACHE):
    """Total units dispatched by `customer` in `month_str` (YYYY-MM),
    summed from the cache. Returns 0 if there's no data (either nothing
    dispatched, or the cache doesn't cover that period yet).

    Does NOT call the CartonCloud API - call refresh_volume() first if
    today's or yesterday's figures need to be current.
    """
    cache = cache if cache is not None else load_volume_cache(path)
    total = 0
    for date_str, day in cache.items():
        if not date_str.startswith(month_str):
            continue
        total += day.get("by_customer", {}).get(customer, {}).get("units", 0)
    return total


def all_customers_units_for_month(month_str, cache=None, path=VOLUME_CACHE):
    """{customer: units_dispatched} for every customer with volume in
    `month_str` (YYYY-MM), summed from the cache."""
    cache = cache if cache is not None else load_volume_cache(path)
    totals = {}
    for date_str, day in cache.items():
        if not date_str.startswith(month_str):
            continue
        for cust, stats in day.get("by_customer", {}).items():
            totals[cust] = totals.get(cust, 0) + stats.get("units", 0)
    return totals

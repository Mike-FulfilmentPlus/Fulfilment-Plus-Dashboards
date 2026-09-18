"""
Daily snapshot of outbound orders sitting in AWAITING_STOCK or DRAFT status.

Purpose: the Order Cut Off KPI should start its "clock" when an order
becomes pickable (AWAITING_PICK_AND_PACK), not when it was originally
created, for orders that were held up waiting on stock - whether that hold
shows up as AWAITING_STOCK or as DRAFT (some customers' orders sit in DRAFT
specifically because product is missing, and only get released to a real
status once stock arrives). CartonCloud's API does not expose a
status-change history, so this script polls daily, records which orders
are currently in either held status, and detects when they later move off
it (the "transition date" = stock arrived / order became pickable).

Intended to run once per day via a scheduled task. Safe to run multiple
times per day - it only appends genuinely new orders and only updates a
transition_date once (first time it's detected).

Output: awaiting_stock_log.csv in the project folder, columns:
  order_id, customer, reference, order_created, first_seen_date,
  transition_date, transition_status

extract_kpi_data.py can later read this log: for an outbound order whose
id appears here with a transition_date inside the target month, use
transition_date (not timestamps.created) as the Order Cut Off clock start.

Usage:
    python daily_status_snapshot.py
"""

import csv
import datetime
import os
import requests
from zoneinfo import ZoneInfo

CREDS_FILE = "credentials.env"
LOG_FILE = "awaiting_stock_log.csv"
NZ_TZ = ZoneInfo("Pacific/Auckland")

FIELDNAMES = [
    "order_id", "customer", "reference", "order_created",
    "first_seen_date", "transition_date", "transition_status",
]


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


HELD_STATUSES = ("AWAITING_STOCK", "DRAFT")
# Both statuses represent an order that exists but isn't yet real fulfilment
# activity: AWAITING_STOCK is explicitly waiting on inventory, and DRAFT is
# often used the same way (e.g. an order held back because product is
# missing, only released to a real status once stock arrives). Either way,
# the Order Cut Off clock should start when the order LEAVES this state, not
# when it was first created - see extract_kpi_data.py / mtd_kpis.py's use of
# this log for the actual clock-start adjustment.

IN_PLAY_STATUSES = ("AWAITING_PICK_AND_PACK", "PICKED", "PACKING_IN_PROGRESS", "DISPATCHED")
# An order is only genuinely "in play" for the Order Cut Off clock once it
# reaches AWAITING_PICK_AND_PACK (i.e. stock is available and it's ready to
# be picked) or a later stage of the normal fulfilment flow. If a
# held order instead comes back as REJECTED (stuck on an unresolved
# allocation/inventory error) or anything else outside this list, that is
# NOT a real transition - the order is still not in play, so we deliberately
# do not record a transition_date for it yet. It keeps getting re-checked on
# every future run until it actually reaches one of these statuses (or stays
# stuck indefinitely, which is a genuine ops problem to chase separately -
# both DRAFT/AWAITING_STOCK and REJECTED are fully excluded from the KPI by
# NON_KPI_STATUSES in extract_kpi_data.py / mtd_kpis.py regardless).


def get_awaiting_stock_orders(token, tenant_id):
    """All outbound orders currently in AWAITING_STOCK or DRAFT status (all pages)."""
    headers = {
        "Accept-Version": "1",
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    condition = {
        "type": "OrCondition",
        "conditions": [
            {
                "type": "TextComparisonCondition",
                "field": {"type": "JsonField", "pointer": "/status"},
                "value": {"type": "ValueField", "value": status},
                "method": "EQUAL_TO",
            }
            for status in HELD_STATUSES
        ],
    }
    results = []
    page = 1
    while True:
        resp = requests.post(
            f"https://api.cartoncloud.com/tenants/{tenant_id}/outbound-orders/search",
            params={"size": 100, "page": page},
            headers=headers,
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


def get_order_status(token, tenant_id, order_id):
    headers = {"Accept-Version": "1", "Authorization": f"Bearer {token}"}
    resp = requests.get(
        f"https://api.cartoncloud.com/tenants/{tenant_id}/outbound-orders/{order_id}",
        headers=headers,
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json().get("status")


def load_log():
    if not os.path.exists(LOG_FILE):
        return []
    with open(LOG_FILE, newline="") as f:
        return list(csv.DictReader(f))


def save_log(rows):
    with open(LOG_FILE, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(rows)


def main():
    today = datetime.datetime.now(NZ_TZ).date().isoformat()

    creds = load_credentials()
    token = get_cc_token(creds)
    tenant_id = creds["CARTONCLOUD_TENANT_ID"]

    current = get_awaiting_stock_orders(token, tenant_id)
    current_ids = {o["id"] for o in current}
    print(f"{today}: {len(current)} order(s) currently AWAITING_STOCK")

    rows = load_log()
    known_ids = {r["order_id"] for r in rows}

    # Add newly-seen AWAITING_STOCK orders
    added = 0
    for o in current:
        if o["id"] in known_ids:
            continue
        rows.append({
            "order_id": o["id"],
            "customer": o["customer"]["name"],
            "reference": o.get("references", {}).get("numericId", ""),
            "order_created": o.get("timestamps", {}).get("created", {}).get("time", ""),
            "first_seen_date": today,
            "transition_date": "",
            "transition_status": "",
        })
        added += 1
    if added:
        print(f"  added {added} newly-seen AWAITING_STOCK order(s)")

    # Check previously-seen orders that no longer appear AWAITING_STOCK/DRAFT
    # and have not yet had a transition recorded.
    transitioned = 0
    still_stuck = 0
    for r in rows:
        if r["transition_date"]:
            continue
        if r["order_id"] in current_ids:
            continue
        # No longer AWAITING_STOCK/DRAFT - find out what it became.
        try:
            status = get_order_status(token, tenant_id, r["order_id"])
        except requests.RequestException as e:
            print(f"  WARNING: could not check order {r['order_id']}: {e}")
            continue
        if status not in IN_PLAY_STATUSES:
            # Left the held status but landed somewhere that still isn't
            # real fulfilment activity (e.g. REJECTED) - not a genuine
            # transition yet. Leave transition_date blank and keep
            # re-checking on future runs.
            still_stuck += 1
            continue
        r["transition_date"] = today
        r["transition_status"] = status
        transitioned += 1
    if transitioned:
        print(f"  {transitioned} order(s) transitioned to a real in-play status")
    if still_stuck:
        print(f"  {still_stuck} order(s) left AWAITING_STOCK/DRAFT but are still not "
              f"in play (e.g. REJECTED) - not counted as a transition yet")

    save_log(rows)
    print(f"Log now has {len(rows)} row(s) -> {LOG_FILE}")


if __name__ == "__main__":
    main()

"""
Verification for the exceptions-check workflow. Run from pipeline/ after
exceptions_edit.py. Uses the same CartonCloud queries and deadline logic as
mtd_kpis.py; prints order numbers/company names only, never credentials.

Argument: path to the ORIGINAL (pre-edit) workbook, to compare deadlines.
"""

import datetime
import os
import sys

import openpyxl

sys.path.insert(0, os.getcwd())  # run from pipeline/
import mtd_kpis as mk  # noqa: E402

NZ = mk.NZ_TZ
CHECK_DATE = datetime.date(2026, 10, 5)
RURAL_PCS = mk.load_rural_postcode_set_co()
RURAL_STREETS = mk.load_rural_street_lookup_co()


def deadline_for(o, exceptions):
    """Mirror of the deadline logic in mtd_kpis.compute_order_cut_off_mtd."""
    overrides = exceptions["cutoff_overrides"]
    content_overrides = exceptions.get("content_overrides", {})
    holidays = exceptions.get("public_holidays", set())
    transitions = mk.load_awaiting_stock_transitions()
    cust = o["customer"]["name"]
    created_dt = mk.parse_iso(o["timestamps"]["created"]["time"]).astimezone(NZ)
    raw_created_date = created_dt.date()
    t = transitions.get(o.get("id"))
    if t and t > created_dt.date():
        created_dt = datetime.datetime.combine(t, datetime.time(0, 0), tzinfo=NZ)
    deadline = mk.order_cut_off_deadline(created_dt, holidays=holidays)
    override = mk.get_cutoff_override(overrides, cust, raw_created_date)
    if override:
        deadline = max(deadline, mk.order_cut_off_deadline(
            created_dt, cutoff_hour=override["cutoff_hour"],
            lenient_next_day=override["lenient_next_day"], holidays=holidays,
            after_cutoff_extra_days=override.get("after_cutoff_extra_days", 0)))
    req = o.get("details", {}).get("deliver", {}).get("requiredDate")
    if req:
        addr = o.get("details", {}).get("deliver", {}).get("address", {})
        pc = str(addr.get("postcode") or "").strip()
        days = (3 if mk.is_si_co(pc) else 1) + (1 if mk.is_rural_co(
            pc, addr.get("address1") or "", addr.get("city") or "", RURAL_PCS, RURAL_STREETS) else 0)
        alt = mk.business_days_before(datetime.date.fromisoformat(req), days, holidays)
        deadline = max(deadline, datetime.datetime.combine(alt, datetime.time(17, 0), tzinfo=NZ))
    for co in content_overrides.get(cust, []):
        if co.get("effective_from") and raw_created_date < co["effective_from"]:
            continue
        if co.get("effective_to") and raw_created_date > co["effective_to"]:
            continue
        getter = mk.CONTENT_MATCH_FIELDS.get(co["match_field"])
        if getter and co["contains"] and co["contains"] in (getter(o) or "").lower():
            deadline = max(deadline, mk.deadline_for_rule(created_dt, co["rule"], holidays))
    return deadline


def main():
    original_wb = sys.argv[1]
    exc_old = mk.load_exceptions(openpyxl.load_workbook(original_wb))
    exc_new = mk.load_exceptions(openpyxl.load_workbook(mk.WORKBOOK))

    now = datetime.datetime.now(NZ)
    month_start = now.date().replace(day=1)
    month_str = month_start.strftime("%Y-%m")
    tomorrow = (now.date() + datetime.timedelta(days=1)).isoformat()

    creds = mk.load_credentials()
    token = mk.get_cc_token(creds)
    tenant = creds["CARTONCLOUD_TENANT_ID"]
    outbound = mk.search_all_pages(
        token, tenant, "outbound-orders",
        mk.date_range_condition("/timestamps/created/time", month_start.isoformat(), tomorrow),
    )
    print(f"Fetched {len(outbound)} outbound orders created {month_start} to {now.date()} (same query as mtd_kpis.py)")

    names = sorted({o["customer"]["name"] for o in outbound if "vix" in o["customer"]["name"].lower()})
    print("Customer names containing 'vix':", names)

    # ---- Vixxen orders created Mon 5 Oct (NZ) -----------------------------
    vix = [o for o in outbound
           if o["customer"]["name"] == "Vixxen"
           and mk.parse_iso(o["timestamps"]["created"]["time"]).astimezone(NZ).date() == CHECK_DATE]
    print(f"\n===== Vixxen orders created {CHECK_DATE} (NZ): {len(vix)} =====")
    print(f"{'Order':>7}  {'Status':<24} {'Created (NZ)':<17} {'Farmers?':<8} {'Deadline before':<17} {'Deadline after':<17} Delivery company")
    non_farmers = 0
    for o in sorted(vix, key=lambda x: x["timestamps"]["created"]["time"]):
        company = o.get("details", {}).get("deliver", {}).get("address", {}).get("companyName", "") or ""
        match = "farmers" in company.lower()
        non_farmers += 0 if match else 1
        created = mk.parse_iso(o["timestamps"]["created"]["time"]).astimezone(NZ)
        d_old = deadline_for(o, exc_old)
        d_new = deadline_for(o, exc_new)
        print(f"{o.get('references', {}).get('numericId', ''):>7}  {o.get('status', ''):<24} "
              f"{created:%Y-%m-%d %H:%M}  {'YES' if match else 'NO':<8} {d_old:%Y-%m-%d %H:%M}  "
              f"{d_new:%Y-%m-%d %H:%M}  {company}")
    print(f"NON-FARMERS VIXXEN ORDERS ON {CHECK_DATE}: {non_farmers}")
    if non_farmers:
        print("STOP: some 5 Oct Vixxen orders do not match 'farmers' - review before committing.")

    # ---- Salt Shark 4938 / 4953 -------------------------------------------
    cut_off_new = [e for e in exc_new["order_exceptions"] if e["kpi"] == "Order Cut Off"]
    print("\n===== Salt Shark 4938 / 4953 =====")
    for ref in ("4938", "4953"):
        hits = [o for o in outbound if o["customer"]["name"] == "Salt Shark"
                and str(o.get("references", {}).get("numericId", "")) == ref]
        if not hits:
            print(f"  {ref}: NOT FOUND among this month's Salt Shark outbound orders")
            continue
        for o in hits:
            created = mk.parse_iso(o["timestamps"]["created"]["time"]).astimezone(NZ)
            exc = mk.find_order_exception(cut_off_new, "Salt Shark", ref, created.date(), month_str)
            disp = o.get("timestamps", {}).get("dispatched", {}).get("time")
            print(f"  {ref}: status={o.get('status')} created={created:%Y-%m-%d %H:%M} "
                  f"dispatched={mk.parse_iso(disp).astimezone(NZ):%Y-%m-%d %H:%M}" if disp else
                  f"  {ref}: status={o.get('status')} created={created:%Y-%m-%d %H:%M} dispatched=none",
                  f"-> exception={'count_as_met (MET)' if exc and exc['action'] == 'count_as_met' else 'NONE'}")

    # ---- Diagnostics -------------------------------------------------------
    print("\n===== Vixxen 5 Oct non-Farmers orders: dispatch vs deadline =====")
    for o in vix:
        company = o.get("details", {}).get("deliver", {}).get("address", {}).get("companyName", "") or ""
        if "farmers" in company.lower():
            continue
        disp = o.get("timestamps", {}).get("dispatched", {}).get("time")
        d = deadline_for(o, exc_new)
        disp_dt = mk.parse_iso(disp).astimezone(NZ) if disp else None
        print(f"  {o.get('references', {}).get('numericId')}: deadline {d:%Y-%m-%d %H:%M}, "
              f"dispatched {disp_dt:%Y-%m-%d %H:%M} -> {'MET' if disp_dt <= d else 'MISSED'}"
              if disp_dt else f"  {o.get('references', {}).get('numericId')}: not dispatched")

    print("\n===== Salt Shark: any order whose references mention 4938 / 4953 =====")
    lookback = mk.search_all_pages(
        token, tenant, "outbound-orders",
        mk.date_range_condition("/timestamps/created/time", "2026-09-01", tomorrow),
    )
    for o in lookback:
        refs = o.get("references", {})
        blob = " ".join(str(v) for v in refs.values())
        if any(x in blob for x in ("4938", "4953")):
            created = mk.parse_iso(o["timestamps"]["created"]["time"]).astimezone(NZ)
            disp = o.get("timestamps", {}).get("dispatched", {}).get("time")
            print(f"  customer={o['customer']['name']} refs={refs} status={o.get('status')} "
                  f"created={created:%Y-%m-%d %H:%M} dispatched="
                  f"{mk.parse_iso(disp).astimezone(NZ):%Y-%m-%d %H:%M}" if disp else
                  f"  customer={o['customer']['name']} refs={refs} status={o.get('status')} "
                  f"created={created:%Y-%m-%d %H:%M} dispatched=none")

    print("\n===== Salt Shark October orders currently scored as MISSED =====")
    for o in outbound:
        if o["customer"]["name"] != "Salt Shark" or o.get("status") in mk.NON_KPI_STATUSES:
            continue
        d = deadline_for(o, exc_new)
        disp = o.get("timestamps", {}).get("dispatched", {}).get("time")
        disp_dt = mk.parse_iso(disp).astimezone(NZ) if disp else None
        missed = (disp_dt > d) if disp_dt else (now > d)
        if missed:
            created = mk.parse_iso(o["timestamps"]["created"]["time"]).astimezone(NZ)
            print(f"  refs={o.get('references', {})} status={o.get('status')} created={created:%Y-%m-%d %H:%M} "
                  f"deadline={d:%Y-%m-%d %H:%M} dispatched={disp_dt:%Y-%m-%d %H:%M}" if disp_dt else
                  f"  refs={o.get('references', {})} status={o.get('status')} created={created:%Y-%m-%d %H:%M} "
                  f"deadline={d:%Y-%m-%d %H:%M} dispatched=none")

    # ---- Order Cut Off MTD, before vs after -------------------------------
    print("\n===== Order Cut Off MTD (mtd_kpis.compute_order_cut_off_mtd) =====")
    old = mk.compute_order_cut_off_mtd(outbound, exc_old, now, month_str)
    new = mk.compute_order_cut_off_mtd(outbound, exc_new, now, month_str)
    for cust in sorted(set(old) | set(new)):
        o_, n_ = old.get(cust, (0, 0)), new.get(cust, (0, 0))
        mark = "   <-- changed" if o_ != n_ else ""
        if mark or cust in ("Vixxen", "Flo & Frankie", "Salt Shark"):
            print(f"  {cust}: {o_[0]}/{o_[1]} -> {n_[0]}/{n_[1]}{mark}")
    ff = new.get("Flo & Frankie")
    if ff:
        print(f"Flo & Frankie all met: {ff[0] == ff[1]}")
    else:
        print("Flo & Frankie: no scored Order Cut Off orders yet this month")


if __name__ == "__main__":
    main()

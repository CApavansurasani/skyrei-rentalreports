#!/usr/bin/env python3
"""publish_cf.py — Cash Flow data builder for the SKY REI suite (schema cf.v1).
Sources: pl-data.json (GL), loans-data.json (engine), Cashflow.xlsx (actual cash),
AR Balances.xlsx (Dec-25 tenant closings), Leases.csv (tenant windows).
Gates: charges/collections/principal/events ties to the cent + per-tenant ledger
identity. If ANY gate fails: the exact breaks are printed and NOTHING is written.
Usage: python3 publish_cf.py [--cashflow ...] [--ar ...] [--leases ...]
                             [--pl ...] [--loans ...] [--out cf-data.json]"""
import json, csv, re, sys, argparse
from datetime import date, timedelta
from collections import defaultdict
import openpyxl
import pandas as pd

AP = argparse.ArgumentParser()
AP.add_argument("--cashflow", default="My files/Cashflow.xlsx")
AP.add_argument("--ar",       default="My files/AR Balances.xlsx")
AP.add_argument("--leases",   default="My files/Leases.csv")
AP.add_argument("--pl",       default="pl-data.json")
AP.add_argument("--loans",    default="loans-data.json")
AP.add_argument("--out",      default="cf-data.json")
ARGS = AP.parse_args()

MONTHS = ["2026-01","2026-02","2026-03","2026-04","2026-05"]
N = len(MONTHS)
OPEX_ACCTS = ["Repairs - Contractors","Repairs - Materials","Repairs - Payroll",
              "Rent Ready - Contractors","Rent Ready - Materials","Rent Ready - Payroll",
              "Property Management Expense","Utilities","Other Misc Cost","Renter's Insurance"]
STAT_ACCTS = ["Property Tax","Insurance","HOA Fee"]
CAPEX_ACCTS = ["Capex - Appliances","Capex - HVAC"]

def z(): return [0.0]*N
def rnd(x): return round(x + 0.0, 2)
def rl(a): return [rnd(v) for v in a]

# ---------------- GL ----------------
pl = json.load(open(ARGS.pl))
assert pl["periods"] == MONTHS, "period mismatch"
door_ent, gl = {}, defaultdict(lambda: defaultdict(z))
for e,d,a,vals in pl["rows"]:
    door_ent[d] = e
    row = gl[a][d]
    for i,v in enumerate(vals):
        if v: row[i] += v
OVERHEAD = set(pl["meta"]["overheadDoors"])
dw = pl["meta"]["doorWindows"]

# ---------------- loans engine ----------------
ld = json.load(open(ARGS.loans))
door_events = defaultdict(lambda: defaultdict(float))     # door -> month -> amt
events_list = []                                          # financing events strip
for l in ld["loans"]:
    ndoors = len(l["origProps"])
    for ev in l["releases"]:
        if ev["m"] in MONTHS:
            amt = ev["paid"] if ndoors == 1 else ev["base"]
            door_events[ev["prop"]][ev["m"]] += amt
            for dprop, cut in (ev.get("cuts") or {}).items():
                door_events[dprop][ev["m"]] += cut
            events_list.append({"m": ev["m"], "type": "release", "loan": l["id"], "lender": l["lender"],
                                "prop": ev["prop"], "paid": ev["paid"], "base": ev["base"], "excess": ev["excess"]})
    for ev in l["prepays"]:
        if ev["m"] in MONTHS:
            for dprop, cut in (ev.get("cuts") or {}).items():
                door_events[dprop][ev["m"]] += cut
            events_list.append({"m": ev["m"], "type": "prepay", "loan": l["id"], "lender": l["lender"],
                                "amt": ev["amt"]})
    # payoffs: schedule events not covered by releases/prepays of that month
    rel_pre = defaultdict(float)
    for ev in l["releases"]:
        if ev["m"] in MONTHS: rel_pre[ev["m"]] += ev["paid"]
    for ev in l["prepays"]:
        if ev["m"] in MONTHS: rel_pre[ev["m"]] += ev["amt"]
    for s in l["schedule"]:
        if s["m"] in MONTHS and s.get("events"):
            resid = rnd(s["events"] - rel_pre.get(s["m"], 0.0))
            if abs(resid) > 0.005:
                shares, tot = {}, 0.0
                mi = MONTHS.index(s["m"])
                prevm = "2025-12" if mi == 0 else MONTHS[mi-1]
                for dprop in l["origProps"]:
                    dd = ld["doors"].get(dprop)
                    if not dd: continue
                    prev = dd["monthly"].get(prevm, {}).get("close", 0.0)
                    shares[dprop] = prev; tot += prev
                for dprop, sh in shares.items():
                    if tot > 0: door_events[dprop][s["m"]] += resid * sh / tot
                events_list.append({"m": s["m"], "type": "payoff", "loan": l["id"], "lender": l["lender"],
                                    "amt": resid})
# new draws in window
for l in ld["loans"]:
    if l["start"] in MONTHS:
        events_list.append({"m": l["start"], "type": "draw", "loan": l["id"], "lender": l["lender"],
                            "amt": l["orig"], "doors": len(l["origProps"]), "firstPay": l["firstPay"]})
# engine per-door interest / scheduled principal
eng_int, eng_prin_total, eng_prin_sched = defaultdict(z), defaultdict(z), defaultdict(z)
for dprop, dd in ld["doors"].items():
    for i,m in enumerate(MONTHS):
        rec = dd["monthly"].get(m)
        if not rec: continue
        eng_int[dprop][i] += rec["interest"]
        eng_prin_total[dprop][i] += rec["principal"]
        eng_prin_sched[dprop][i] += rec["principal"] - door_events[dprop].get(m, 0.0)
# loan-level scheduled principal + events (assert target)
loan_sched, loan_events = z(), z()
for l in ld["loans"]:
    for s in l["schedule"]:
        if s["m"] in MONTHS:
            i = MONTHS.index(s["m"])
            loan_sched[i] += s["principal"]; loan_events[i] += s.get("events") or 0.0
# engine escrow per month (active = schedule row exists, month >= firstPay, open > 0)
esc_eng_total = z()
esc_door = defaultdict(z)
for l in ld["loans"]:
    if not l.get("escrow"): continue
    srows = {s["m"]: s for s in l["schedule"]}
    for i,m in enumerate(MONTHS):
        s = srows.get(m)
        if not s or m < l["firstPay"] or s["open"] <= 0: continue
        esc_eng_total[i] += l["escrow"]
        prevm = "2025-12" if i == 0 else MONTHS[i-1]
        shares, tot = {}, 0.0
        for dprop in l["origProps"]:
            dd = ld["doors"].get(dprop)
            if not dd: continue
            prev = dd["monthly"].get(prevm, {}).get("close", 0.0)
            if prev > 0: shares[dprop] = prev; tot += prev
        for dprop, sh in shares.items():
            esc_door[dprop][i] += l["escrow"] * sh / tot

# ---------------- Cashflow.xlsx ----------------
wb = openpyxl.load_workbook(ARGS.cashflow, read_only=True, data_only=True)
ws = wb["Cashflow"]
coll, coll_n = defaultdict(z), defaultdict(lambda:[0]*N)
cf_int, cf_prin, cf_ti, cf_ins, cf_pt, cf_hoa = (defaultdict(z) for _ in range(6))
na_cells = defaultdict(list)   # acct -> [(door, month)]
for r in ws.iter_rows(min_row=2, max_col=15, values_only=True):
    if r[0] is None and r[1] is None: continue
    dprop, acct = r[0], r[1]
    for i in range(N):
        v = r[3+i]
        if v is None: continue
        if isinstance(v, str):
            na_cells[acct].append([dprop, MONTHS[i]]); continue
        v = float(v)
        if acct == 'Rental Collection': coll[dprop][i] += v; coll_n[dprop][i] += 1
        elif acct == 'Interest Cost': cf_int[dprop][i] += v
        elif acct == 'Principal Payment': cf_prin[dprop][i] += v
        elif acct == 'T& I Payment': cf_ti[dprop][i] += v
        elif acct == 'Insurance': cf_ins[dprop][i] += v
        elif acct == 'Property Tax': cf_pt[dprop][i] += v
        elif acct == 'HOA Fee' and r[2] is None: cf_hoa[dprop][i] += v
wb.close()

# ---------------- AR balances ----------------
ar = pd.read_excel(ARGS.ar, header=None, skiprows=2, names=["prop","tenant","bal"])
ar_open = {}          # door -> (tenant, balance|None)
for _, r in ar.iterrows():
    if pd.isna(r["prop"]): continue
    dprop = str(r["prop"]).strip()
    ten = str(r["tenant"]).strip() if pd.notna(r["tenant"]) else "—"
    b = None
    if pd.notna(r["bal"]):
        try: b = float(r["bal"])
        except Exception: b = None
    ar_open[dprop] = (ten, b)

# ---------------- Leases ----------------
def pdate(s):
    m, d, y = s.split("/")
    return date(int(y), int(m), int(d))
leases = defaultdict(list)
with open(ARGS.leases, newline='', encoding='utf-8-sig') as f:
    for row in csv.DictReader(f):
        p = row["Index - Property"].strip()
        leases[p].append({"tenant": re.sub(r"\s+", " ", row["Tenants"].strip()),
                          "start": pdate(row["Start date"]), "end": pdate(row["End date"]),
                          "status": row["Status"].strip(), "type": row["Type"].strip(),
                          "rent": float(row["Rent"] or 0)})
def month_bounds(m):
    y, mo = int(m[:4]), int(m[5:])
    first = date(y, mo, 1)
    nxt = date(y+1, 1, 1) if mo == 12 else date(y, mo+1, 1)
    return first, nxt - timedelta(days=1)
def occupant(dprop, m):
    first, last = month_bounds(m)
    cands = [L for L in leases.get(dprop, []) if L["start"] <= last and L["end"] >= first]
    if not cands: return None
    cands.sort(key=lambda L: (L["start"], L["end"]))
    return cands[-1]

# ---------------- assemble doors ----------------
def normname(s): return re.sub(r"\s+", " ", s.strip().casefold())
all_doors = sorted(set(door_ent) | set(ar_open) | set(coll) | set(leases), key=str.casefold)
doors_out = {}
flags_count = defaultdict(int)
for dprop in all_doors:
    if dprop in OVERHEAD:
        continue
    ent = door_ent.get(dprop)
    w = dw.get(dprop) or [None, None, ent]
    inGL = dprop in door_ent
    chg = z()
    for i in range(N):
        chg[i] = gl["Rental Income"][dprop][i] + gl["Vacancy Loss"][dprop][i] if inGL else 0.0
    opexA = {a: rl(gl[a][dprop]) for a in OPEX_ACCTS if inGL and any(gl[a][dprop])}
    opex = [rnd(sum(gl[a][dprop][i] for a in OPEX_ACCTS)) for i in range(N)] if inGL else z()
    statA = {a: rl(gl[a][dprop]) for a in STAT_ACCTS if inGL and any(gl[a][dprop])}
    stat = [rnd(sum(gl[a][dprop][i] for a in STAT_ACCTS)) for i in range(N)] if inGL else z()
    capex = [rnd(sum(gl[a][dprop][i] for a in CAPEX_ACCTS)) for i in range(N)] if inGL else z()
    claims = rl(gl["Insurance Claim"][dprop]) if inGL else z()
    # tenant ledger buckets
    buckets, order = {}, []
    def bucket(name, src):
        k = normname(name)
        if k not in buckets:
            buckets[k] = {"tenant": name, "src": src, "opening": 0.0, "openingBlank": False,
                          "chg": z(), "coll": z(), "occ": [False]*N, "nl": 0}
            order.append(k)
        return buckets[k]
    ar_ten, ar_bal = ar_open.get(dprop, (None, None))
    if ar_ten is not None:
        b = bucket(ar_ten, "ar")
        if ar_bal is None:
            b["openingBlank"] = True; flags_count["blank-opening"] += 1
        else:
            b["opening"] = ar_bal
    occ_names = []
    for i, m in enumerate(MONTHS):
        L = occupant(dprop, m)
        if L:
            name, src = L["tenant"], "lease"
        elif ar_ten and ar_ten not in ("Vacant", "Sold", "Held for Sale", "—"):
            # no lease covers the month: activity goes to the AR-file tenant (last known occupant)
            name, src = ar_ten, "ar"
        else:
            name, src = "Vacant", "system"
        occ_names.append(name if L else (name + " ·nl" if name != "Vacant" else name))
        b = bucket(name, src)
        b["occ"][i] = True
        b["chg"][i] += chg[i]
        b["coll"][i] += coll[dprop][i]
        if L:
            b["src"] = "lease"   # lease-covered occupancy upgrades provenance
        else:
            b["nl"] += 1         # month attributed with no lease covering it
            if abs(chg[i]) > 0.005 or abs(coll[dprop][i]) > 0.005:
                flags_count["no-lease-activity"] += 1
    tenants = []
    for k in order:
        b = buckets[k]
        run = b["opening"]; close = z()
        for i in range(N):
            run += b["chg"][i] - b["coll"][i]; close[i] = rnd(run)
        # attach the tenant's lease facts (latest lease matching this name)
        lease = None
        for L in sorted(leases.get(dprop, []), key=lambda L: L["start"]):
            if normname(L["tenant"]) == k:
                lease = {"start": L["start"].isoformat(), "end": L["end"].isoformat(),
                         "rent": rnd(L["rent"]), "status": L["status"], "type": L["type"]}
        tenants.append({"tenant": b["tenant"], "src": b["src"], "opening": rnd(b["opening"]),
                        "openingBlank": b["openingBlank"], "occ": b["occ"], "nl": b["nl"],
                        "chg": rl(b["chg"]), "coll": rl(b["coll"]), "close": close,
                        "lease": lease})
    fl = []
    if not inGL: fl.append("no-gl")
    if dprop not in ar_open and inGL: fl.append("no-ar-row")
    if dprop not in leases and inGL: fl.append("no-lease")
    if ar_ten is not None and ar_ten not in ("Vacant","Sold","Held for Sale","—"):
        if not any(normname(ar_ten) == normname(L["tenant"]) for L in leases.get(dprop, [])):
            fl.append("ar-tenant-not-in-leases")
    for f in fl: flags_count[f] += 1
    doors_out[dprop] = {
        "entity": ent or (w[2] if w and len(w) > 2 else None) or "—",
        "window": [w[0], w[1]] if w else [None, None],
        "inGL": inGL,
        "chg": rl(chg), "coll": rl(coll[dprop]), "nrec": coll_n[dprop],
        "opex": opex, "opexByAcct": opexA, "stat": stat, "statByAcct": statA,
        "capex": capex, "claims": claims,
        "int": rl(gl["Interest Cost"][dprop]) if inGL else z(),
        "prinSched": rl(eng_prin_sched[dprop]), "prinTotal": rl(eng_prin_total[dprop]),
        "escrow": rl(esc_door[dprop]),
        "fileInt": rl(cf_int[dprop]), "filePrin": rl(cf_prin[dprop]),
        "occupant": occ_names, "tenants": tenants, "flags": fl,
        "arTenant": ar_ten,
    }

# memo series for drift panel
memo = {a: rl([sum(gl[a][d][i] for d in gl[a]) for i in range(N)])
        for a in ["Other Income","Collection Loss","SD - Forfeit Income","Increase/(Decrease) in Rent"]}

out = {
 "schema": "cf.v1",
 "meta": {
   "months": MONTHS,
   "published": "mock — not published",
   "source": {"gl": "pl-data.json (pl.v2)", "engine": "loans-data.json (loans.v2)",
              "cash": "Cashflow.xlsx", "ar": "AR Balances.xlsx (Dec-25 closing)", "leases": "Leases.csv"},
   "overheadDoors": sorted(OVERHEAD),
   "assumptions": [
     "Charges = Rental Income + Vacancy Loss per door-month (rent for the occupied period). True tenant charges — SD, late fees, utility billbacks — pending owner data.",
     "Charges & collections attributed to the lease occupant of the month; two tenants in one month → the later lease (owner rule, 2026-07-16). Months no lease covers: activity goes to the AR-file tenant (last known occupant); to 'Vacant' only when none exists.",
     "Collection Loss NOT deducted from tenant balances — if it represents AR write-offs, a −write-off term is needed (owner decision pending).",
     "Scheduled principal & door escrow from the loans engine (lender-verified); prepay/release/payoff cash shown as financing events, not operating outflow.",
     "Statutory cash-out = actual T&I escrow + direct PT/Ins payments from Cashflow.xlsx; HOA cash pending ('NA') — GL accrual shown as proxy.",
     "Bank balances, contributions/distributions, sale proceeds pending — cash kept is derived, not yet tied to bank movement.",
   ],
 },
 "doors": doors_out,
 "overhead": {
   "opex": [rnd(sum(gl[a]["General Company Expenses"][i] for a in OPEX_ACCTS)) for i in range(N)],
   "opexByAcct": {a: rl(gl[a]["General Company Expenses"]) for a in OPEX_ACCTS if any(gl[a]["General Company Expenses"])},
   "stat": [rnd(sum(gl[a]["General Company Expenses"][i] for a in STAT_ACCTS)) for i in range(N)],
   "int": rl(gl["Interest Cost"]["General Company Expenses"]),
   "claims": rl(gl["Insurance Claim"]["General Company Expenses"]),
   "capex": [rnd(sum(gl[a]["General Company Expenses"][i] for a in CAPEX_ACCTS)) for i in range(N)],
   "coll": rl(coll["General Company Expenses"]),
   "chg": rl([gl["Rental Income"]["General Company Expenses"][i] + gl["Vacancy Loss"]["General Company Expenses"][i] for i in range(N)]),
   "fileTI": rl(cf_ti.get("General Company Expenses", z())),
 },
 "events": sorted(events_list, key=lambda e: (e["m"], e["type"])),
 "recon": {
   "interest": {"file": rl([sum(v[i] for v in cf_int.values()) for i in range(N)]),
                "gl":   rl([sum(gl["Interest Cost"][d][i] for d in gl["Interest Cost"]) for i in range(N)]),
                "engine": rl([sum(eng_int[d][i] for d in eng_int) for i in range(N)])},
   "principal": {"file": rl([sum(v[i] for v in cf_prin.values()) for i in range(N)]),
                 "engineSched": rl(loan_sched),
                 "naCells": na_cells.get("Principal Payment", [])},
   "escrow": {"fileTI": rl([sum(v[i] for v in cf_ti.values()) for i in range(N)]),
              "engine": rl(esc_eng_total),
              "tiRows": {k: rl(v) for k, v in cf_ti.items()}},
   "collections": {"file": rl([sum(v[i] for v in coll.values()) for i in range(N)]),
                   "receipts": int(sum(sum(v) for v in coll_n.values())),
                   "naCells": na_cells.get("Rental Collection", [])},
   "statutoryCash": {"insurance": rl([sum(v[i] for v in cf_ins.values()) for i in range(N)]),
                     "propertyTax": rl([sum(v[i] for v in cf_pt.values()) for i in range(N)]),
                     "hoa": rl([sum(v[i] for v in cf_hoa.values()) for i in range(N)]),
                     "hoaNA": len(na_cells.get("HOA Fee", []))},
   "financingEvents": {"total": rl(loan_events)},
   "memo": memo,
   "naCells": {k: v for k, v in na_cells.items()},
 },
}

# ---------------- assertions ----------------
def s2(x): return round(x, 2)
errs = []
for i in range(N):
    a = sum(d["chg"][i] for d in doors_out.values()) + out["overhead"]["chg"][i]
    b = sum(gl["Rental Income"][d][i] for d in gl["Rental Income"]) + sum(gl["Vacancy Loss"][d][i] for d in gl["Vacancy Loss"])
    if abs(a-b) > 0.02: errs.append(f"chg tie {MONTHS[i]}: {a} vs {b}")
for i in range(N):
    a = sum(d["coll"][i] for d in doors_out.values()) + out["overhead"]["coll"][i]
    b = sum(v[i] for v in coll.values())
    if abs(a-b) > 0.02: errs.append(f"coll tie {MONTHS[i]}: {a} vs {b}")
for i in range(N):
    a = sum(d["prinSched"][i] for d in doors_out.values())
    if abs(a - loan_sched[i]) > 0.02: errs.append(f"prinSched tie {MONTHS[i]}: {s2(a)} vs {s2(loan_sched[i])}")
for i in range(N):
    a = sum(door_events[d].get(MONTHS[i], 0.0) for d in door_events)
    if abs(a - loan_events[i]) > 0.02: errs.append(f"events tie {MONTHS[i]}: {s2(a)} vs {s2(loan_events[i])}")
for dprop, d in doors_out.items():
    for i in range(N):
        lhs = sum(t["close"][i] for t in d["tenants"])
        opening = sum(t["opening"] for t in d["tenants"])
        rhs = opening + sum(d["chg"][j] - d["coll"][j] for j in range(i+1))
        if abs(lhs - rhs) > 0.05:
            errs.append(f"ledger identity {dprop} {MONTHS[i]}: {s2(lhs)} vs {s2(rhs)}"); break
if errs:
    print(f"GATES: {len(errs)} FAILURES — NOTHING WRITTEN")
    for e in errs[:50]: print("  ✗", e)
    sys.exit(1)
print("GATES: ALL GREEN")
json.dump(out, open(ARGS.out, "w"))
import os
print(f"{ARGS.out} bytes:", os.path.getsize(ARGS.out))
print("doors:", len(doors_out), "| events:", len(events_list), "| flags:", dict(flags_count))
print("engine escrow/mo:", rl(esc_eng_total))
print("sched principal/mo:", rl(loan_sched))
tot_open = sum(t["opening"] for d in doors_out.values() for t in d["tenants"])
print("AR opening total:", s2(tot_open))
close = sum(t["close"][N-1] for d in doors_out.values() for t in d["tenants"])
print(f"portfolio AR close {MONTHS[N-1]}: {s2(close)}")

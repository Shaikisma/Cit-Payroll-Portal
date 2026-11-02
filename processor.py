"""
CIT Payroll processor
---------------------
Takes last month's payroll workbook + the invoice workbook + the health-insurance
carrier file and produces the NEXT payroll workbook (formulas kept, Checks sheet
kept, Change Log rebuilt, a "Review Report" sheet added).

Nothing here is "magic": every automatic change is written to the Change Log
sheet, and anything that needs a human decision is highlighted + listed in the
Review Report sheet.
"""
import calendar
import difflib
import re
from collections import Counter, defaultdict
from datetime import date, datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

YELLOW = PatternFill("solid", fgColor="FFF2CC")   # please verify
ORANGE = PatternFill("solid", fgColor="FFD9B3")   # data missing
RED = PatternFill("solid", fgColor="F8CBAD")      # probably remove / fix
HEADER = PatternFill("solid", fgColor="1F3864")

NOTES_COL = 48  # AV - "Auto-review notes"


class PortalError(Exception):
    """Friendly error that is shown to HR on screen."""


# ----------------------------------------------------------------- helpers
def norm_name(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def norm_vendor(s):
    return re.sub(r"\s+", " ", str(s or "").strip().lower())


def num(x):
    return float(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else None


def add_month(y, m, n=1):
    t = y * 12 + (m - 1) + n
    return t // 12, t % 12 + 1


def parse_period(text):
    m = re.match(r"\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*-", str(text or ""))
    return (int(m.group(3)), int(m.group(1))) if m else None


# --------------------------------------------------------- file detection
def classify(path):
    """Return 'payroll', 'invoice' or 'health' by looking at the content."""
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        raise PortalError("One of the files is not a valid .xlsx Excel file.")
    try:
        for ws in wb.worksheets:
            flat = ""
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i > 12:
                    break
                flat += " ".join(str(c) for c in row if c is not None) + " "
            if "Hrs Validation per Invoice Sheet" in flat:
                return "payroll"
            if "EmployeeID" in flat and "Invoice No" in flat:
                return "invoice"
            if "Location Detail" in flat:
                return "health"
    finally:
        wb.close()
    return None


# ------------------------------------------------------------ invoice file
def load_invoices(path, year, month):
    wb = load_workbook(path, read_only=True)
    names = wb.sheetnames
    wb.close()
    mn = calendar.month_name[month].lower()
    sheet = None
    for n in names:
        low = n.lower()
        if str(year) in low and (mn in low or mn[:3] in low):
            sheet = n
            break
    if not sheet:
        raise PortalError(
            f"The invoice file has no sheet for {calendar.month_name[month]} {year}. "
            f"Sheets found: {', '.join(names)}. Add the new month sheet and upload again."
        )
    df = pd.read_excel(path, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]
    if "EmployeeID" not in df.columns:
        raise PortalError(f"Sheet '{sheet}' has no 'EmployeeID' column.")
    df = df[df["EmployeeID"].notna()].copy()
    df["EmployeeID"] = df["EmployeeID"].astype(str).str.strip()
    for c in ["Invoice Hours", "Invoice rate", "Consultant rate", "Percentage"]:
        df[c] = pd.to_numeric(df.get(c), errors="coerce")
    df["vk"] = df["Vendor Name"].map(norm_vendor)

    lookup, by_emp = {}, defaultdict(set)
    for (eid, vk), g in df.groupby(["EmployeeID", "vk"]):
        def last_pos(col):
            s = g[col].dropna()
            s = s[s > 0]
            return float(s.iloc[-1]) if len(s) else None
        lookup[(eid, vk)] = dict(
            hours=float(g["Invoice Hours"].fillna(0).sum()),
            inv_rate=last_pos("Invoice rate"),
            cons_rate=last_pos("Consultant rate"),
            pct=last_pos("Percentage"),
            duration=str(g["Invoice Duration"].iloc[0]) if "Invoice Duration" in g else "",
        )
        by_emp[eid].add(vk)
    return sheet, lookup, by_emp


def find_invoice(lookup, by_emp, eid, vendor):
    vk = norm_vendor(vendor)
    if (eid, vk) in lookup:
        return lookup[(eid, vk)], (eid, vk)
    best, score = None, 0
    for cand in by_emp.get(eid, []):
        r = difflib.SequenceMatcher(None, vk, cand).ratio()
        if r > score:
            best, score = cand, r
    if best and score >= 0.8:
        return lookup[(eid, best)], (eid, best)
    return None, (eid, vk)


# ------------------------------------------------------------- health file
def parse_health(path):
    wb = load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    section, people, adj_total, carrier_total, billing = None, [], 0.0, None, ""
    name_re = re.compile(r"^\s*[A-Za-z'\-\. ]+,\s*[A-Za-z'\-\. ]+\s*$")
    for row in ws.iter_rows(values_only=True):
        cells = [c for c in row if c not in (None, "")]
        if not cells:
            continue
        texts = [str(c).strip() for c in cells if isinstance(c, str)]
        nums = [float(c) for c in cells if isinstance(c, (int, float))]
        joined = " ".join(texts)
        m = re.search(r"([A-Za-z]+ \d{4}) Final Invoice", joined)
        if m and not billing:
            billing = m.group(1)
        if "Current" in texts:
            section = "current"
        elif "Adjustments" in texts:
            section = "adj"
        elif "Adjusted Totals" in texts:
            section = "totals"
        if "Total Current" in texts and nums:
            carrier_total = nums[-1]
        if "Total Adjustment" in texts and nums:
            adj_total = nums[-1]
        if section == "current" and nums:
            for t in texts:
                if name_re.match(t) and not t.lower().startswith("total"):
                    tier = next((x for x in texts if re.fullmatch(r"[A-Z]{3}", x)), "")
                    people.append(dict(name=t.strip(), tier=tier, cost=round(nums[-1], 2)))
                    break
    if not people:
        raise PortalError("Could not find any employees in the health-insurance file.")
    if carrier_total is None:
        carrier_total = round(sum(p["cost"] for p in people), 2)
    return dict(people=people, adj_total=round(adj_total, 2),
                carrier_total=round(carrier_total, 2), billing=billing)


def match_person(name, people, used):
    n = norm_name(name.replace("(Expenses)", ""))
    best, score = None, 0
    for i, p in enumerate(people):
        r = difflib.SequenceMatcher(None, n, norm_name(p["name"])).ratio()
        if r > score:
            best, score = i, r
    return (best, score) if best is not None and score >= 0.85 else (None, score)


# ------------------------------------------------------------- main routine
def generate(payroll_path, invoice_path, health_path, out_path,
             pay_month=None, run_date=None, config=None):
    config = config or {}
    employer = {k: float(v) for k, v in config.get("employer_contribution", {}).items()}

    wb = load_workbook(payroll_path)
    ws = wb.worksheets[0]
    old_title = ws.title

    # ---- locate blocks --------------------------------------------------
    summary_row = next((r for r in range(2, ws.max_row + 1)
                        if str(ws.cell(r, 1).value or "").strip().startswith("CIT - Summary")), None)
    if not summary_row:
        raise PortalError("Payroll file: could not find the 'CIT - Summary' block.")
    total_row = next((r for r in range(2, summary_row)
                      if not ws.cell(r, 1).value and str(ws.cell(r, 23).value or "").startswith("=SUM")), None)
    if not total_row:
        raise PortalError("Payroll file: could not find the totals row of the Regular block.")
    rows = [r for r in range(2, summary_row)
            if ws.cell(r, 1).value in ("Regular", "MPS") and ws.cell(r, 2).value]
    regular_rows = [r for r in rows if r < total_row]

    # ---- which month are we producing? ---------------------------------
    periods = Counter(parse_period(ws.cell(r, 12).value) for r in rows)
    periods.pop(None, None)
    if not periods:
        raise PortalError("Payroll file: 'Pay Period' column (L) is empty.")
    prev_ym = periods.most_common(1)[0][0]
    if pay_month:
        y, m = map(int, pay_month.split("-"))
    else:
        y, m = add_month(*prev_ym, 1)
    ry, rm = add_month(y, m, 1)
    run = datetime.strptime(run_date, "%Y-%m-%d").date() if run_date else date(ry, rm, 15)
    last_day = calendar.monthrange(y, m)[1]
    p_start, p_end = datetime(y, m, 1), datetime(y, m, last_day)
    p_text = f"{m:02d}/01/{y}-{m:02d}/{last_day:02d}/{y}"

    inv_sheet, lookup, by_emp = load_invoices(invoice_path, y, m)
    health = parse_health(health_path)

    log, issues = [], []

    def flag(r, level, text, cols=(), fill=None):
        issues.append(dict(row=r, level=level, text=text,
                           employee=(ws.cell(r, 3).value or "") if r else "-"))
        if r:
            fill = fill or {"warn": YELLOW, "missing": ORANGE, "fix": RED}.get(level, YELLOW)
            for c in cols:
                ws[f"{c}{r}"].fill = fill
            cur = ws.cell(r, NOTES_COL).value
            ws.cell(r, NOTES_COL).value = (cur + " | " if cur else "") + text

    def setv(r, col, new, why):
        cell = ws[f"{col}{r}"]
        old = cell.value
        if isinstance(old, str) and old.startswith("="):
            return
        same = (old == new) or (num(old) is not None and num(new) is not None
                                and abs(old - new) < 1e-9)
        cell.value = new
        if not same:
            log.append((f"{col}{r}", "" if old is None else str(old), "" if new is None else str(new), why))

    # ---- classify rows & pre-compute -----------------------------------
    carry = set()
    for r in rows:
        if parse_period(ws.cell(r, 12).value) != prev_ym:
            carry.add(r)
    normal = [r for r in rows if r not in carry]

    key_count = Counter()
    row_inv = {}
    for r in normal:
        eid = str(ws.cell(r, 2).value).strip()
        inv, key = find_invoice(lookup, by_emp, eid, ws.cell(r, 16).value)
        row_inv[r] = (inv, key)
        key_count[key] += 1

    # previous ADP-hours logic (before we change anything)
    prev_standard = {}
    for r in regular_rows:
        q, rr, mh = num(ws.cell(r, 17).value), num(ws.cell(r, 18).value), num(ws.cell(r, 13).value)
        if None in (q, rr, mh):
            continue
        s = q * rr
        ae = num(ws.cell(r, 31).value) or s
        net = s * mh - (num(ws.cell(r, 20).value) or 0) - (num(ws.cell(r, 22).value) or 0)
        af = net / ae if ae else None
        ao = num(ws.cell(r, 41).value)
        prev_standard[r] = (af is not None and ao is not None and abs(ao - af) < 0.002,
                            isinstance(ws.cell(r, 31).value, (int, float)))
    old_advance_total = sum(num(ws.cell(r, 26).value) or 0 for r in regular_rows)
    old_advance_before = num(ws.cell(summary_row + 5, 2).value)

    # ---- health assignment ---------------------------------------------
    used, assigned, seen_ids = set(), {}, {}
    for r in regular_rows:
        eid = str(ws.cell(r, 2).value).strip()
        seen_ids.setdefault(eid, []).append(r)
    for eid, rlist in seen_ids.items():
        name = ws.cell(rlist[0], 3).value or ""
        idx, score = match_person(name, health["people"], used)
        if idx is None:
            continue
        used.add(idx)
        pref = next((r for r in rlist if (num(ws.cell(r, 27).value) or 0) > 0), rlist[0])
        assigned[pref] = (health["people"][idx], eid)

    # ---- update each normal row ----------------------------------------
    for r in normal:
        eid = str(ws.cell(r, 2).value).strip()
        inv, key = row_inv[r]
        freq = str(ws.cell(r, 8).value or "").lower()
        is_reg = r in regular_rows

        setv(r, "K", datetime(run.year, run.month, run.day), "Pay date set to run date.")
        setv(r, "L", p_text, "Pay period rolled to new month.")
        if ws.cell(r, 37).value is not None:
            setv(r, "AK", p_start, "Pay period start rolled forward.")
            setv(r, "AL", p_end, "Pay period end rolled forward.")

        if inv:
            old_m = num(ws.cell(r, 13).value)
            setv(r, "N", round(inv["hours"], 3), f"Hours per '{inv_sheet}' invoice sheet.")
            if key_count[key] > 1:
                flag(r, "warn", f"Invoice total {inv['hours']:g} hrs is shared by {key_count[key]} payroll rows - "
                     f"split hours manually (hours kept at {old_m:g}).", ["M", "N"])
            else:
                setv(r, "M", round(inv["hours"], 3), f"Hours per '{inv_sheet}' invoice sheet.")
                if "week" in freq or "week" in inv["duration"].lower():
                    flag(r, "warn", f"Weekly billing: invoice shows {inv['hours']:g} hrs (last month {old_m:g}). "
                         "Confirm the hours for the full month.", ["M", "N"])
            for col, k, label in (("O", "inv_rate", "invoice rate"), ("Q", "cons_rate", "consultant rate"),
                                  ("R", "pct", "consultant %")):
                if key_count[key] > 1:
                    break  # shared invoice (split / expenses row): never overwrite rates automatically
                new, old = inv[k], num(ws[f"{col}{r}"].value)
                if new is not None:
                    if old is not None and abs(new - old) > 1e-6:
                        flag(r, "warn", f"{label.capitalize()} changed {old:g} -> {new:g} (per invoice sheet).", [col])
                    setv(r, col, new, f"{label.capitalize()} per invoice sheet.")
        else:
            flag(r, "missing", f"No invoice found for {eid} / {ws.cell(r, 16).value} in '{inv_sheet}'. "
                 "Hours and rates kept from last month.", ["M", "N", "P"])

        # one-off items never carry over
        for col, label in (("T", "expense reimbursement"), ("V", "paid expenses"),
                           ("Y", "bonus"), ("Z", "deduction/advance")):
            old = ws[f"{col}{r}"].value
            if num(old):
                flag(r, "warn", f"Last month's {label} of {old:g} was cleared. Re-enter if it continues.", [col])
            if old is not None and not (isinstance(old, str) and old.startswith("=")):
                setv(r, col, None if is_reg else 0, f"One-off item ({label}) cleared for the new month.")
        if ws.cell(r, 6).value:
            setv(r, "F", None, "Last month's comment cleared.")
        if ws.cell(r, 47).value:
            setv(r, "AU", None, "Last month's ADP note cleared.")

        # ADP pay hours
        if is_reg and num(ws.cell(r, 41).value) is not None:
            std, hard_ae = prev_standard.get(r, (False, False))
            if std:
                ws.cell(r, 41).value = f"=ROUNDDOWN(ROUND(AF{r},6),3)"
                if hard_ae:
                    flag(r, "warn", "ADP rate (AE) is hard-coded - kept. Verify it is still right.", ["AE"])
            else:
                flag(r, "warn", "ADP pay hours (AO) kept from last month - this row had special handling "
                     "(split check / merged rows). Re-check.", ["AO"])

        # health insurance
        if is_reg:
            old_aa = num(ws.cell(r, 27).value) or 0
            if r in assigned:
                person, pid = assigned[r]
                cost = person["cost"]
                emp = employer.get(pid, 0.0)
                new_aa = round(cost - emp, 2)
                setv(r, "AA", new_aa, f"Health insurance per carrier file ({person['tier']} tier).")
                if emp:
                    flag(r, "warn", f"Carrier cost {cost:,.2f}, employer pays {emp:,.2f} (config.json) -> deduct {new_aa:,.2f}.", ["AA"])
                elif old_aa and abs(old_aa - new_aa) > 0.005:
                    flag(r, "warn", f"Health deduction changed {old_aa:,.2f} -> {new_aa:,.2f} per carrier file.", ["AA"])
            else:
                if old_aa:
                    flag(r, "warn", f"Previously deducted health insurance {old_aa:,.2f} but employee is not on the "
                         "carrier file - deduction set to 0.", ["AA"])
                if ws.cell(r, 27).value is not None:
                    setv(r, "AA", None, "Not on this month's carrier file.")

    # ---- carry-over rows ----------------------------------------------
    for r in carry:
        flag(r, "fix", f"Row is for pay period {ws.cell(r, 12).value} (older month). "
             "Delete or update it if it was already paid.", ["A", "B", "C", "L", "M"], RED)

    # ---- unmatched carrier people --------------------------------------
    for i, p in enumerate(health["people"]):
        if i not in used:
            flag(0, "fix", f"Carrier file lists {p['name']} ({p['tier']}, {p['cost']:,.2f}) but no matching employee "
                 "was found in the payroll sheet.")
    if health["adj_total"]:
        flag(0, "warn", f"Carrier invoice has adjustments totalling {health['adj_total']:,.2f}. They are NOT "
             "deducted automatically - decide whether to recover them.")

    # ---- summary block --------------------------------------------------
    s = summary_row
    ws.cell(s, 2).value = f"{run:%b} {run.day:02d} {run.year} Regular Payroll"
    ids = {str(ws.cell(r, 2).value).strip() for r in regular_rows}
    manual_ids = {str(ws.cell(r, 2).value).strip() for r in regular_rows
                  if str(ws.cell(r, 9).value or "").strip().lower() == "manual"}
    ws.cell(s + 1, 2).value = len(ids)
    ws.cell(s + 2, 2).value = len(manual_ids)
    ws.cell(s + 4, 2).value = datetime(run.year, run.month, run.day)
    if old_advance_before is not None:
        ws.cell(s + 5, 2).value = round(old_advance_before - old_advance_total, 2)
        flag(0, "warn", f"'Advance Before payroll' set to last month's 'after' balance "
             f"({old_advance_before - old_advance_total:,.2f}). Verify.")
    ws.cell(s + 7, 2).value = health["carrier_total"]
    employer_sum = round(sum(employer.get(pid, 0) for _, (_, pid) in assigned.items()), 2)

    # ---- sheets: rename, Checks, Change Log, Review Report -------------
    new_title = f"{run:%b}-{run.day:02d} payroll"
    ws.title = new_title
    if "Checks" in wb.sheetnames:
        ck = wb["Checks"]
        for row in ck.iter_rows():
            for c in row:
                if isinstance(c.value, str) and old_title in c.value:
                    c.value = c.value.replace(f"'{old_title}'", f"'{new_title}'")
                if isinstance(c.value, str) and c.value.startswith("Input: employer-paid"):
                    ck.cell(c.row, 2).value = employer_sum
    ws.cell(1, NOTES_COL).value = "Auto-review notes"
    ws.cell(1, NOTES_COL).font = Font(bold=True)
    ws.column_dimensions["AV"].width = 70

    if "Change Log" in wb.sheetnames:
        cl = wb["Change Log"]
        cl.delete_rows(2, cl.max_row)
        for i, rec in enumerate(log, start=2):
            for j, v in enumerate(rec, start=1):
                cl.cell(i, j).value = v
    rv = wb.create_sheet("Review Report")
    for j, h in enumerate(["Level", "Row", "Employee", "What to check"], start=1):
        c = rv.cell(1, j, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEADER
    order = {"fix": 0, "missing": 1, "warn": 2}
    for i, it in enumerate(sorted(issues, key=lambda x: (order[x["level"]], x["row"])), start=2):
        rv.cell(i, 1, {"fix": "FIX", "missing": "MISSING", "warn": "VERIFY"}[it["level"]])
        rv.cell(i, 2, it["row"] or "")
        rv.cell(i, 3, it["employee"])
        rv.cell(i, 4, it["text"]).alignment = Alignment(wrap_text=True, vertical="top")
    for col, w in zip("ABCD", (10, 7, 32, 110)):
        rv.column_dimensions[col].width = w

    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)

    # ---- numbers for the on-screen preview -----------------------------
    preview, tot = [], dict(gross=0.0, health=0.0, net=0.0)
    for r in rows:
        q, rr, mh = (num(ws.cell(r, c).value) or 0 for c in (17, 18, 13))
        t, v = num(ws.cell(r, 20).value) or 0, num(ws.cell(r, 22).value) or 0
        y_, z, aa = (num(ws.cell(r, c).value) or 0 for c in (25, 26, 27))
        pay_rate = q * rr
        gross = pay_rate * mh - t
        net = gross - v + y_ - (z + aa)
        if r in regular_rows:
            tot["gross"] += gross - v
            tot["health"] += aa
            tot["net"] += net
        notes = [i for i in issues if i["row"] == r]
        preview.append(dict(row=r, block="Regular" if r in regular_rows else "MPS",
                            id=ws.cell(r, 2).value, name=ws.cell(r, 3).value,
                            vendor=ws.cell(r, 16).value, hours=mh, rate=round(pay_rate, 2),
                            gross=round(gross - v, 2), health=round(aa, 2), net=round(net, 2),
                            level=("fix" if any(i["level"] == "fix" for i in notes) else
                                   "missing" if any(i["level"] == "missing" for i in notes) else
                                   "warn" if notes else "ok")))
    return dict(
        pay_period=p_text, run_date=run.isoformat(), invoice_sheet=inv_sheet,
        sheet_name=new_title, health_billing=health["billing"],
        employees=len(ids), rows=len(rows), changes=len(log),
        totals={k: round(v, 2) for k, v in tot.items()},
        mps_gross=round(sum(p["gross"] for p in preview if p["block"] == "MPS"), 2),
        issues=sorted(issues, key=lambda x: (order[x["level"]], x["row"])),
        preview=preview,
    )

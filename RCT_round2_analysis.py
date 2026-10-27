"""
Cash transfer RCT — Survey Round 2: data merging and analysis (Python)
=====================================================================

Reproduces, in one script, the Round 2 work done in this session and the analyses that compare
Round 2 with Round 1:

  Step 1  merge_round2()             Merge the 11 repeat sheets of Survey_round_2.xlsx into the main sheet
                                     (long -> wide), add `round` = 2, `payment` and `in_payment_list` from
                                     RCT_payment_round1.xls, compute Section 3 activity costs for Virippu 2026,
                                     costs per acre (÷ rice_acres_2026) and outlier flags, and write
                                     Survey_round_2_merged_wide.xlsx (sheets: data, codebook, data_checks,
                                     cost_summary, cost_outliers, cost_review).
  Step 2  timeline_items2()          Item-level costs with activity / payment dates on the 2026 calendar and
                                     weekly aggregates -> items2.csv, timeline_weekly_round2.json
                                     (the data behind the interactive Round 2 timeline page)
  Step 3  cost_comparison2()         Per-acre cost components and timing relative to the payments,
                                     treatment vs control -> Round2_cost_comparison_payment.xlsx
  Step 4  bubble_plot2()             Weekly cost-per-acre bubble plot with the cumulative cost per acre and
                                     the two payment weeks -> Round2_cost_per_acre_bubble_overlay.jpg / _panels.jpg
  Step 5  pesticide_fert3()          Pesticide vs 3rd fertiliser dose by payment group, Round 2
                                     -> Round2_pesticide_vs_3rd_fertiliser.jpg
  Step 6  pesticide_fert3_rounds()   The same with Round 1 as the pre-payment check
                                     -> Pesticide_3rd_fertiliser_round1_vs_round2.jpg
  Step 7  pmkisan_pies_round2()      Use of the Rs 2,000 (pmkisan_survey_use), treatment and control pies
          pmkisan_pie_round1()       Use of PM-KISAN money (pmkisan_use), Round 1, combined pie
          pmkisan_excel()            -> Round2_pmkisan_survey_use_pies.jpg, Round1_pmkisan_use_pie.jpg,
          pmkisan_excel_add_round1()    PMKISAN_use_pies_round1_round2.xlsx
  Step 8  spending_after_transfer()  Spending in the weeks right after each payment, treatment vs control,
                                     Round 1 same weeks as placebo, panel difference-in-differences
                                     -> Round2_spending_after_transfer.xlsx
  Step 9  household_expenditure()    Household consumption (Section 4) per person, treatment vs control in
          household_expenditure_excel()  both rounds, DiD and ANCOVA
                                     -> Household_expenditure_per_person_round1_round2.xlsx

Usage
-----
  1. Run RCT_round1_analysis.py first: Steps 1, 6, 7, 8 and 9 read its outputs
     (Survey_Round_1_merged_wide.xlsx and items.csv).
  2. Put Survey_round_2.xlsx, Survey_round_2_questionnarie.xlsx and RCT_payment_round1.xls in DATA_DIR and run:
         python RCT_round2_analysis.py                          # uses DATA_DIR / R1_DIR / OUT_DIR below
         python RCT_round2_analysis.py <data_dir> <round1_output_dir> <out_dir>
     With the defaults, both scripts read inputs from the current folder and write to ./outputs.

Requirements: pandas, numpy, scipy, openpyxl, matplotlib, and xlrd (to read the .xls payment file:
`pip install xlrd`). If xlrd is not available, save RCT_payment_round1.xls as .xlsx in the same folder.
The PM-KISAN workbook keeps its counts as Excel formulas; Excel calculates them when the file is opened.
"""
import json
import os
import re
import sys
import textwrap
from pathlib import Path

import numpy as np
import pandas as pd
from scipy import stats
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.lines import Line2D
from matplotlib.patches import Patch
from matplotlib.ticker import FixedLocator, FuncFormatter, NullLocator
from openpyxl import Workbook, load_workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =====================================================================================
# CONFIGURATION — edit these three paths
# =====================================================================================
DATA_DIR = Path(".")            # folder with Survey_round_2.xlsx, Survey_round_2_questionnarie.xlsx, RCT_payment_round1.xls
R1_DIR = Path("outputs")        # outputs of RCT_round1_analysis.py (Survey_Round_1_merged_wide.xlsx, items.csv)
OUT_DIR = Path("outputs")       # all Round 2 outputs are written here


def read_payment(path):
    """Read RCT_payment_round1.xls (needs xlrd); fall back to an .xlsx copy with the same name."""
    path = Path(path)
    try:
        return pd.read_excel(path)
    except (ImportError, ValueError, FileNotFoundError) as err:
        alt = path.with_suffix(".xlsx")
        if alt.exists():
            return pd.read_excel(alt)
        raise RuntimeError(f"Could not read {path} ({err}). Install xlrd (pip install xlrd) "
                           f"or save the file as {alt.name}.") from err



# =====================================================================================
# SECTION 3 COST RULES (Virippu 2026) — used by Steps 1–3
# =====================================================================================
CODES = [999, 9999]                 # "don't know" codes used by enumerators
LUMP = re.compile(r"contract|contact|lump|total", re.I)
SMALL_UNIT = re.compile(r"\bml\b|\bg\b|gm|gram", re.I)
# chemical fertilisers and lime are sold per kg / bag at well above Rs 3 per unit
PER_KG_ITEM = re.compile(r"urea|potash|\bmop\b|fact|dap|sulph|phos|npk|complex|ammon|micro|magnes|zinc|borax|lime|line|dolomite", re.I)
FLOOR_ANY = 1.0       # amount ÷ quantity below Rs 1 per unit: cannot be a total for any item → amount is a unit price
FLOOR_FERT = 3.0      # for fertiliser / lime items the floor is Rs 3 per unit

# (key, label, section, qty, unit, price)  — price is per unit in the questionnaire
FIXED = [
    ("a_seed", "Seed", "A", "a_seed_qty", "a_seed_unit", "a_seed_price"),
    ("a_land", "Land preparation", "A", "a_land_qty", "a_land_unit", "a_land_price"),
    ("b_labour", "Nursery labour", "B", "b_labour_qty", "b_labour_unit", "b_labour_amt"),
    ("b_fert", "Nursery fertiliser", "B", "b_fert_qty", "b_fert_unit", "b_fert_amt"),
    ("c_mainlab", "Main-plot preparation labour", "C", "c_labour_qty", "c_labour_unit", "c_labour_amt"),
    ("c_mach", "Main-plot preparation machinery", "C", "c_mach_qty", "c_mach_unit", "c_mach_amt"),
    ("c_translab", "Transplanting labour", "C", "c_labour_qty_2", "c_labour_unit_2", "c_labour_amt_2"),
    ("c_fert1lab", "Labour, 1st fertiliser dose", "C", "c_labour_fert1_qty", "c_labour_fert1_unit", "c_labour_fert1_amt"),
    ("c_gapfill", "Gap-filling labour", "C", "c_labour_gapfilling_qty", "c_labour_gapfilling_unit", "c_labour_gapfilling_amt"),
    ("d_weed", "Herbicide / weedicide", "D", "d_weed_qty", "d_weed_unit", "d_weed_price"),
    ("d_weedlab", "Labour, weedicide application", "D", "d_weedlab_qty", "d_weedlab_unit", "d_weedlab_price"),
    ("d_fert2lab", "Labour, 2nd fertiliser dose", "D", "d_labour_fert2_qty", "d_labour_fert2_unit", "d_labour_fert2_amt"),
    ("e_pest", "Pesticide", "E", "e_pest_qty", "e_pest_unit", "e_pest_price"),
    ("e_pestlab", "Labour, pesticide application", "E", "e_labour_pest_qty", "e_labour_pest_unit", "e_labour_pest_amt"),
    ("e_fert3lab", "Labour, 3rd fertiliser dose", "E", "e_labour_fert3_qty", "e_labour_fert3_unit", "e_labour_fert3_amt"),
    ("f_mach", "Harvesting machinery", "F", "f_mach_qty", "f_mach_unit", "f_mach_amt"),
    ("f_tractor", "Tractor, harvest transport", "F", "f_tractor_qty", "f_tractor_unit", "f_tractor_amt"),
    ("f_labour", "Labour, drying and packing", "F", "f_labour_qty", "f_labour_unit", "f_labour_amt"),
]
# repeat blocks with a per-unit price: (key, label, section, prefix for qty/unit/price, max items)
FERT = [
    ("c_fert1", "Fertiliser, 1st dose", "C", "c_fert1", 4),
    ("d_fert2", "Fertiliser, 2nd dose", "D", "d_fert2", 6),
    ("e_fert3", "Fertiliser, 3rd dose / pesticide", "E", "e_fert3", 5),
]
# repeat blocks asking "amount paid (Rs.)": (key, label, section, prefix, max items)
OTHER = [
    ("a_other", "Other items, pre-nursery", "A", "a_other", 3),
    ("b_other", "Other items, nursery", "B", "b_other", 1),
    ("c_other", "Other items, after transplanting", "C", "c_other", 2),
    ("d_other", "Other items, after 2nd dose", "D", "d_other", 2),
    ("e_other", "Other items, pesticide / 4th dose", "E", "e_other", 4),
    ("f_other", "Other items, after harvest", "F", "f_other", 2),
]
ORDER = ["a_seed", "a_land", "a_other", "b_labour", "b_fert", "b_other", "c_mainlab", "c_mach", "c_translab", "c_fert1",
         "c_fert1lab", "c_gapfill", "c_other", "d_weed", "d_weedlab", "d_fert2", "d_fert2lab", "d_other",
         "e_pest", "e_pestlab", "e_fert3", "e_fert3lab", "e_other", "f_mach", "f_tractor", "f_labour", "f_other"]
SECTIONS = {"A": "Pre-nursery preparation", "B": "Nursery preparation", "C": "Transplanting",
            "D": "Weedicide / fertiliser application", "E": "Pesticide / fertiliser application", "F": "Harvesting"}


def _num(s):
    return pd.to_numeric(s, errors="coerce")


def unit_price_cost(q, unit, p):
    """cost = quantity × unit price, row-wise. Returns (cost, status)."""
    q, p = _num(q), _num(p)
    unit = unit.astype(str).where(unit.notna(), "")
    cost = pd.Series(np.nan, index=q.index)
    status = pd.Series("", index=q.index, dtype=object)
    code = q.isin(CODES) | p.isin(CODES)
    both_blank = q.isna() & p.isna()
    zero = (q == 0) | (p == 0)
    lump = q.isna() & (p > 0) & unit.str.contains(LUMP)
    ok = q.notna() & p.notna()
    cost[ok] = q[ok] * p[ok]
    cost[lump] = p[lump]; status[lump] = "quantity blank, unit is contract → price taken as lump sum"
    cost[zero & ~code] = 0.0
    cost[both_blank] = 0.0
    cost[code] = np.nan; status[code] = "999 / 9999 code in quantity or price → missing"
    half = cost.isna() & ~code
    status[half & q.isna()] = "price recorded but quantity blank → missing"
    status[half & p.isna()] = "quantity recorded but price blank → missing"
    return cost, status


def other_cost(q, unit, a, item):
    """'Amount paid (Rs.)' items: amount is the total paid, unless it is evidently a unit price."""
    q, a = _num(q), _num(a)
    unit = unit.astype(str).where(unit.notna(), "")
    item = item.astype(str).where(item.notna(), "")
    qq = q.where(~q.isin(CODES))
    cost = a.copy().astype(float)
    status = pd.Series("", index=q.index, dtype=object)
    cost[a.isna() & qq.isna()] = 0.0
    cost[a == 0] = 0.0
    ratio = a / qq
    base = (qq > 1) & (a > 0) & ~unit.str.contains(SMALL_UNIT) & ~a.isin(CODES)
    r1 = base & (ratio < FLOOR_ANY)
    r2 = base & ~r1 & (ratio < FLOOR_FERT) & item.str.contains(PER_KG_ITEM)
    amb = base & ~r1 & ~r2 & (ratio < FLOOR_FERT)
    cost[r1 | r2] = (qq * a)[r1 | r2]
    status[r1] = f"amount ÷ quantity < Rs {FLOOR_ANY:g} per unit → amount is a unit price: cost = quantity × amount"
    status[r2] = f"fertiliser/lime item, amount ÷ quantity < Rs {FLOOR_FERT:g} per unit → amount is a unit price: cost = quantity × amount"
    status[amb] = f"ambiguous (amount ÷ quantity between Rs {FLOOR_ANY:g} and {FLOOR_FERT:g}) → kept as total paid; please check"
    code = a.isin(CODES)
    cost[code] = np.nan; status[code] = "999 / 9999 code in amount → missing"
    miss = a.isna() & (qq > 0)
    cost[miss] = np.nan; status[miss] = "quantity recorded but amount blank → missing"
    return cost, status


def compute(d):
    """Return (costs DataFrame aligned to d, review rows, labels dict)."""
    out = pd.DataFrame(index=d.index)
    incomplete = pd.DataFrame(index=d.index)
    review = []
    labels = {}
    old_form = d["__version__"] == "vhWzpqVLG8rKBa5kqM2f3d"
    acres = _num(d["rice_acres_2026"])

    def log(key, label, status, extra):
        for i in status.index[status != ""]:
            review.append({"farmers_id": d.at[i, "farmers_id"], "_index": d.at[i, "_index"], "activity": key,
                           "activity_label": label, **{k: v.at[i] if hasattr(v, "at") else v for k, v in extra.items()},
                           "decision": status.at[i]})

    for key, label, sec, q, u, p in FIXED:
        c, st = unit_price_cost(d[q], d[u], d[p])
        out[f"cost_{key}"] = c
        incomplete[key] = c.isna()
        labels[key] = (label, sec, f"{q} × {p}")
        log(key, label, st, {"item": "", "quantity": d[q], "unit": d[u], "price_or_amount": d[p], "cost_used": c})

    for key, label, sec, pre, n in FERT:
        tot = pd.Series(0.0, index=d.index)
        anymiss = pd.Series(False, index=d.index)
        for k in range(1, n + 1):
            nm = "name"
            q, u, p = f"{pre}_qty_{k}", f"{pre}_unit_{k}", f"{pre}_price_{k}"
            c, st = unit_price_cost(d[q], d[u], d[p])
            anymiss |= c.isna()
            tot = tot + c.fillna(0)
            log(key, label, st, {"item": d[f"{pre}_{nm}_{k}"], "quantity": d[q], "unit": d[u], "price_or_amount": d[p], "cost_used": c})
        out[f"cost_{key}"] = tot.where(~anymiss | (tot > 0))      # all items unusable → missing
        incomplete[key] = anymiss
        labels[key] = (label, sec, f"Σ items ({pre}_qty_k × {pre}_price_k)")
    out.loc[old_form, "cost_c_fert1"] = np.nan                    # question not in the older form version

    for key, label, sec, pre, n in OTHER:
        tot = pd.Series(0.0, index=d.index)
        anymiss = pd.Series(False, index=d.index)
        for k in range(1, n + 1):
            q, u, a = f"{pre}_qty_{k}", f"{pre}_unit_{k}", f"{pre}_amt_{k}"
            c, st = other_cost(d[q], d[u], d[a], d[f"{pre}_item_{k}"])
            anymiss |= c.isna()
            tot = tot + c.fillna(0)
            log(key, label, st, {"item": d[f"{pre}_item_{k}"], "quantity": d[q], "unit": d[u], "price_or_amount": d[a], "cost_used": c})
        out[f"cost_{key}"] = tot.where(~anymiss | (tot > 0))
        incomplete[key] = anymiss
        labels[key] = (label, sec, f"Σ items: {pre}_amt_k (total paid), or {pre}_qty_k × {pre}_amt_k where the amount is evidently a unit price")

    out = out[[f"cost_{k}" for k in ORDER]]
    incomplete.loc[old_form, "c_fert1"] = False                   # not asked: not counted as incomplete
    n_missing = incomplete.sum(axis=1)
    for s in SECTIONS:
        cols = [f"cost_{k}" for k in ORDER if labels[k][1] == s]
        out[f"cost_sec{s}"] = out[cols].sum(axis=1, min_count=1)
    out["cost_total"] = out[[f"cost_sec{s}" for s in SECTIONS]].sum(axis=1, min_count=1)
    out["cost_n_missing"] = n_missing.astype(int)
    out["cost_c_fert1_not_asked"] = old_form.astype(int)
    valid_acres = acres > 0
    for c in [f"cost_{k}" for k in ORDER] + [f"cost_sec{s}" for s in SECTIONS] + ["cost_total"]:
        out["cpa_" + c[5:]] = (out[c] / acres).where(valid_acres)
    return out, pd.DataFrame(review), labels


def modified_z_log(s):
    """Iglewicz–Hoaglin modified z-score on log(value), positive values only."""
    x = np.log(s[s > 0])
    med = x.median()
    mad = (x - med).abs().median()
    z = pd.Series(np.nan, index=s.index)
    if mad > 0:
        z[x.index] = 0.6745 * (x - med) / mad
    return z


def _fmt(x):
    return f"{x:,.0f}" if abs(x) >= 10 else f"{x:,.2f}"


def _items(d, k):
    """Item-level (qty, unit, price/amount, name) frames for an activity, stacked over repeat positions."""
    for key, label, sec, q, u, p in FIXED:
        if key == k:
            return [pd.DataFrame({"qty": _num(d[q]), "unit": d[u], "price": _num(d[p]), "name": ""})]
    for key, label, sec, pre, n in FERT:
        if key == k:
            return [pd.DataFrame({"qty": _num(d[f"{pre}_qty_{j}"]), "unit": d[f"{pre}_unit_{j}"],
                                  "price": _num(d[f"{pre}_price_{j}"]), "name": d[f"{pre}_name_{j}"]}) for j in range(1, n + 1)]
    for key, label, sec, pre, n in OTHER:
        if key == k:
            return [pd.DataFrame({"qty": _num(d[f"{pre}_qty_{j}"]), "unit": d[f"{pre}_unit_{j}"],
                                  "price": _num(d[f"{pre}_amt_{j}"]), "name": d[f"{pre}_item_{j}"]}) for j in range(1, n + 1)]
    return None


def outliers(d, cst, labels, zcut=3.5):
    acres = _num(d["rice_acres_2026"])
    keys = ORDER + [f"sec{s}" for s in SECTIONS] + ["total"]
    rows, summ = [], []
    flags = pd.DataFrame(index=d.index)
    zs = {}
    for k in keys:
        zs[k] = modified_z_log(cst[f"cpa_{k}"])
        flags[f"out_{k}"] = (zs[k].abs() > zcut).astype(int)

    def acre_note(i):
        a = acres.at[i]
        if a <= 0.5:
            return f"very small rice area ({a:g} acre) inflates the per-acre figure — check rice_acres_2026"
        if a > 20:
            return f"very large rice area ({a:g} acres) for the costs reported — check rice_acres_2026"
        return None

    def cause(i, k, z):
        hi = z > 0
        if k.startswith("sec") or k == "total":
            sec_keys = [a for a in ORDER if k == "total" or labels[a][1] == k[3:]]
            drv = [a for a in sec_keys if flags.at[i, f"out_{a}"] and np.sign(zs[a].get(i, 0)) == np.sign(z)]
            if drv:
                return "driven by: " + "; ".join(f"{labels[a][0]} (Rs {_fmt(cst.at[i, f'cpa_{a}'])}/acre)" for a in drv)
            return acre_note(i) or "no single activity flagged — several moderately unusual items add up"
        parts = _items(d, k)
        allit = pd.concat(parts)
        mq = allit.loc[allit["qty"] > 0, "qty"].median()
        mp = allit.loc[allit["price"] > 0, "price"].median()
        cand = pd.DataFrame([{"qty": f.at[i, "qty"], "unit": f.at[i, "unit"], "price": f.at[i, "price"], "name": f.at[i, "name"]}
                             for f in parts])
        cand["c"] = cand["qty"].fillna(1) * cand["price"].fillna(0)
        r = cand.loc[cand["c"].idxmax()] if hi else cand.loc[cand["price"].fillna(np.inf).idxmin()]
        nm = f"{str(r['name']).strip()}: " if str(r["name"]).strip() not in ("", "nan") else ""
        qr = r["qty"] / mq if pd.notna(r["qty"]) and mq else np.nan
        pr = r["price"] / mp if pd.notna(r["price"]) and mp else np.nan
        unit = "" if pd.isna(r["unit"]) else str(r["unit"]).strip()
        desc = f"{nm}qty {_fmt(r['qty']) if pd.notna(r['qty']) else '—'} {unit} × Rs {_fmt(r['price']) if pd.notna(r['price']) else '—'}".replace("  ", " ")
        what = "amount" if k in [o[0] for o in OTHER] else "unit price"
        if hi:
            if re.search(r"\bkg", unit, re.I) and pd.notna(r["price"]) and r["price"] > 300 and what == "unit price":
                return (f"{desc}. Rs {_fmt(r['price'])} per kg is implausible for this input — looks like the bag price or the total "
                        f"was entered as the unit price")
            if what == "amount":                     # other items: the amount paid is the cost; quantity does not enter it
                if pr >= 10:
                    return f"{desc}. Amount paid is {pr:,.0f}× the median amount (Rs {_fmt(mp)}) — check"
                a = acre_note(i)
                if a:
                    return f"{desc}. {a}"
                return f"{desc}. Amount paid is {pr:,.1f}× the median amount (Rs {_fmt(mp)}) — check"
            if SMALL_UNIT.search(unit) and qr >= 20:
                return (f"{desc}. Quantity is in grams/ml while the price looks like a price per pack, "
                        f"so quantity × price overstates the cost — check")
            if pr >= 10 and qr >= 5:
                return (f"{desc}. Quantity ({qr:,.0f}× median) and unit price ({pr:,.0f}× median) are both high — "
                        f"the unit price is probably the total amount")
            if qr >= 20:
                return f"{desc}. Quantity is {qr:,.0f}× the median ({_fmt(mq)}) — looks like the total amount was entered as quantity"
            if pr >= 10:
                return f"{desc}. Unit price is {pr:,.0f}× the median ({_fmt(mp)}) — looks like the total was entered as unit price"
            a = acre_note(i)
            if a:
                return f"{desc}. {a}"
            if (pd.isna(qr) or qr < 2) and (pd.isna(pr) or pr < 2):
                return f"{desc}. Quantity and {what} are typical; per-acre value is high relative to the {acres.at[i]:g}-acre rice area — check rice_acres_2026"
            return f"{desc}. High quantity ({qr:,.1f}× median) and/or {what} ({pr:,.1f}× median) — check"
        a = acre_note(i)
        if a and acres.at[i] > 20:
            return f"{desc}. {a}"
        if (pd.isna(qr) or qr >= 0.5) and (pd.isna(pr) or pr >= 0.5):
            return (f"{desc}. Quantity and {what} are typical; per-acre value is low for a {acres.at[i]:g}-acre rice area — "
                    f"input may cover only part of the area, or check rice_acres_2026")
        return f"{desc}. Unusually low (qty {qr:,.2f}×, {what} {pr:,.2f}× median) — check units, a partial payment or own/family labour"

    for k in keys:
        cpa = cst[f"cpa_{k}"]
        pos = cpa[cpa > 0]
        z = zs[k]
        q1, q3 = pos.quantile([0.25, 0.75]) if len(pos) else (np.nan, np.nan)
        hi, lo = z > zcut, z < -zcut
        if k in labels:
            lab, sec = labels[k][0], labels[k][1]
        elif k.startswith("sec"):
            lab, sec = f"Section {k[3:]} total: {SECTIONS[k[3:]]}", k[3:]
        else:
            lab, sec = "TOTAL cost of cultivation", "All"
        lx = np.log(pos); mad = (lx - lx.median()).abs().median()
        summ.append({"section": sec, "activity": lab, "variable": f"cpa_{k}",
                     "N with cost > 0": len(pos), "N cost = 0": int((cpa == 0).sum()), "N missing": int(cpa.isna().sum()),
                     "Mean": pos.mean(), "Median": pos.median(), "P5": pos.quantile(0.05), "P95": pos.quantile(0.95),
                     "Min": pos.min(), "Max": pos.max(),
                     "Lower bound": float(np.exp(lx.median() - zcut * mad / 0.6745)),
                     "Upper bound": float(np.exp(lx.median() + zcut * mad / 0.6745)),
                     "N high outliers": int(hi.sum()), "N low outliers": int(lo.sum()),
                     "N > Q3 + 3×IQR": int((pos > q3 + 3 * (q3 - q1)).sum())})
        for i in z.index[hi | lo]:
            rows.append({"farmers_id": d.at[i, "farmers_id"], "_index": d.at[i, "_index"], "section": sec, "activity": lab,
                         "variable": f"cpa_{k}", "direction": "high" if z.at[i] > 0 else "low",
                         "rice_acres_2026": acres.at[i], "cost (Rs.)": cst.at[i, f"cost_{k}"], "cost per acre (Rs.)": cpa.at[i],
                         "median cost per acre": pos.median(), "times the median": cpa.at[i] / pos.median(),
                         "modified z (log)": z.at[i], "likely cause": cause(i, k, z.at[i])})
    # households with no usable rice area have no per-acre value: screen their raw activity cost instead
    no_area = ~(acres > 0)
    for k in ORDER:
        zc = modified_z_log(cst[f"cost_{k}"])
        pos = cst.loc[cst[f"cost_{k}"] > 0, f"cost_{k}"]
        for i in zc.index[(zc > zcut) & no_area]:
            flags.at[i, f"out_{k}"] = 1
            lab, sec = labels[k][0], labels[k][1]
            rows.append({"farmers_id": d.at[i, "farmers_id"], "_index": d.at[i, "_index"], "section": sec, "activity": lab,
                         "variable": f"cpa_{k}", "direction": "high",
                         "rice_acres_2026": acres.at[i], "cost (Rs.)": cst.at[i, f"cost_{k}"], "cost per acre (Rs.)": np.nan,
                         "median cost per acre": np.nan, "times the median": cst.at[i, f"cost_{k}"] / pos.median(),
                         "modified z (log)": zc.at[i],
                         "likely cause": "No rice area recorded, so screened on total cost (times the median = vs median total cost). "
                                         + cause(i, k, zc.at[i]).split(". very small rice area")[0]})
    for k in ORDER:                                    # keep the summary counts in step with the added rows
        extra = sum(1 for r in rows if r["variable"] == f"cpa_{k}" and pd.isna(r["cost per acre (Rs.)"]))
        for srow in summ:
            if srow["variable"] == f"cpa_{k}":
                srow["N high outliers"] += extra
    return pd.DataFrame(summ), pd.DataFrame(rows), flags

C = sys.modules[__name__]     # lets the step functions refer to the cost rules as C.<name>


# =====================================================================================
# STEP 1 — merge sheets, payment, costs  ->  Survey_round_2_merged_wide.xlsx
# =====================================================================================

def merge_round2(IN, PAYFILE, ROUND1, OUT):
    """Merge Survey_round_2 repeat sheets (long) into the main sheet (wide); add round, payment and in_payment_list. (from merge_round2.py)"""
    SURVEY = f"{IN}/Survey_round_2.xlsx"
    QUEST = f"{IN}/Survey_round_2_questionnarie.xlsx"
    MAIN_SHEET = "Survey round2"

    # block -> (count variable it follows, sheet, position variable, section label)
    BLOCKS = [
        ("crop", "other_crops_count", "crop_rep", "crop_pos", "Sec 2: other crops cultivated"),
        ("a_other", "a_other_count", "begin_repeat_0W8hGwBOT", "a_other_pos", "Sec 3A: other items, pre-nursery"),
        ("b_other", "b_other_count", "begin_repeat_oFXmOdKGR", "b_other_pos", "Sec 3B: other items, nursery"),
        ("c_fert1", "c_fert1_count", "begin_repeat_q7UBsOtMr", "c_fert1_pos", "Sec 3C: first-dose fertilisers"),
        ("c_other", "c_other_count", "begin_repeat_kb3NZEaME", "c_other_pos", "Sec 3C: other items, transplanting"),
        ("d_fert2", "d_fert2_count", "begin_repeat_pHGxYqKsw", "d_fert2_pos", "Sec 3D: second-dose fertilisers"),
        ("d_other", "d_other_count", "begin_repeat_9aGJkPAfe", "d_other_pos", "Sec 3D: other items, after 2nd dose"),
        ("e_fert3", "e_fert3_count", "begin_repeat_y8vUlnzPY", "e_fert3_pos", "Sec 3E: pesticides / third-dose fertilisers"),
        ("e_other", "e_other_count", "begin_repeat_PS5svQ16g", "e_other_pos", "Sec 3E: other items"),
        ("f_other", "f_other_count", "begin_repeat_tXgwButw2", "f_other_pos", "Sec 3F: other items, harvesting"),
        ("welfare", "welfare_count", "welfare_rep", "scheme_pos", "Sec 5: welfare schemes"),
    ]

    xl = pd.ExcelFile(SURVEY)
    main = pd.read_excel(xl, sheet_name=MAIN_SHEET)
    n_main = len(main)
    assert main["_index"].is_unique
    assert sorted(b[2] for b in BLOCKS) == sorted(s for s in xl.sheet_names if s != MAIN_SHEET), "sheet list mismatch"

    # ---- questionnaire labels for the codebook ----
    q = pd.read_excel(QUEST, sheet_name="survey")
    qlab = {r["name"]: re.sub(r"\s+", " ", str(r["label::English (en)"])).strip()
            for _, r in q.iterrows() if pd.notna(r["name"]) and pd.notna(r["label::English (en)"])}
    qtype = dict(zip(q["name"], q["type"]))
    ch = pd.read_excel(QUEST, sheet_name="choices")
    choice_lab = {(r["list_name"], str(r["name"])): r["label::English (en)"] for _, r in ch.iterrows()}

    def label_for(var):
        if var in qlab:
            return qlab[var]
        if "/" in var:
            base, opt = var.split("/", 1)
            lab = qlab.get(base, "")
            t = str(qtype.get(base, ""))
            lst = t.split()[-1] if t.startswith("select_multiple") else None
            opt_lab = choice_lab.get((lst, opt)) if lst else None
            if lab:
                return f"{lab} — option {opt}" + (f": {opt_lab}" if opt_lab else "")
        return "(not in the attached questionnaire version)"

    # ---- reshape each block long -> wide ----
    wide_blocks, codebook_extra, checks = {}, {}, []
    NA_TXT = {"na", "n/a", "nil", "none", "no", "0", "-", "nan"}
    for name, count_var, sheet, pos_var, section in BLOCKS:
        assert count_var in main.columns, f"{count_var} not in main sheet"
        long = pd.read_excel(xl, sheet_name=sheet)
        assert long["_parent_index"].isin(main["_index"]).all(), f"{name}: orphan rows"
        assert (long[pos_var] == long.groupby("_parent_index").cumcount() + 1).all(), f"{name}: position mismatch"
        datavars = [c for c in long.columns if not c.startswith("_") and c != pos_var]
        w = long.pivot(index="_parent_index", columns=pos_var, values=datavars)
        max_item = int(long[pos_var].max())
        cols = [(v, k) for k in range(1, max_item + 1) for v in datavars]       # item-major order
        w = w.reindex(columns=pd.MultiIndex.from_tuples(cols))
        w.columns = [f"{v}_{k}" for v, k in cols]
        wide_blocks[count_var] = w
        for v, k in cols:
            codebook_extra[f"{v}_{k}"] = (sheet, section, k, v)

        def row_kind(r):
            vals = [r[v] for v in datavars if pd.notna(r[v])]
            if not vals:
                return "blank"
            if all((isinstance(x, str) and x.strip().lower() in NA_TXT) or
                   (isinstance(x, (int, float)) and not isinstance(x, bool) and x == 0) for x in vals):
                return "placeholder"
            return "substantive"
        long["__kind"] = long.apply(row_kind, axis=1)
        kinds = long["__kind"].value_counts()
        if kinds.get("blank", 0) or kinds.get("placeholder", 0):
            checks.append((f"{name}: repeat rows with no real data (kept as recorded)",
                           int(kinds.get("blank", 0) + kinds.get("placeholder", 0)),
                           f"{int(kinds.get('blank', 0))} fully blank, {int(kinds.get('placeholder', 0))} placeholder entries ('Na' / 0); "
                           f"the form opens one item by default, so these mostly sit in item 1 where {count_var} = 0 or missing"))
        n_items = long[long["__kind"] == "substantive"].groupby("_parent_index").size()
        mi = main.set_index("_index")
        cnt = pd.to_numeric(mi[count_var], errors="coerce")
        rec = n_items.reindex(cnt.index).fillna(0).astype(int)
        bad = cnt.index[cnt.fillna(0) != rec]
        if len(bad):
            det = ", ".join(f"{mi.at[i, 'farmers_id']} ({'NA' if pd.isna(cnt[i]) else int(cnt[i])} vs {rec[i]})" for i in bad)
            checks.append((f"{name}: {count_var} ≠ items with real data", len(bad), "farmers_id (count reported vs items recorded): " + det))
        print(f"{name:8s} rows={len(long):4d} parents={long['_parent_index'].nunique():3d} max_items={max_item} new_cols={len(cols)}")

    # ---- merge: each block right after its count question ----
    new_cols_all = [c for w in wide_blocks.values() for c in w.columns]
    assert not (set(new_cols_all) & set(main.columns)), "column name clash"
    assert len(new_cols_all) == len(set(new_cols_all)), "duplicate new column names"
    order = []
    for c in main.columns:
        order.append(c)
        if c in wide_blocks:
            order.extend(wide_blocks[c].columns)
    merged = main.set_index("_index", drop=False).join(pd.concat(wide_blocks.values(), axis=1), how="left")
    merged = merged[order].reset_index(drop=True).copy()
    assert len(merged) == n_main

    fid_raw = merged["farmers_id"].astype(str).str.strip()
    merged["farmers_id"] = [int(x) if re.fullmatch(r"\d+", x) else x for x in fid_raw]

    # ---- payment (Round 1 payment list) ----
    pay = read_payment(PAYFILE)
    pay.columns = [c.strip() for c in pay.columns]
    pay["farmer_id"] = pay["farmer_id"].astype(str).str.strip()
    pay["Payment"] = pay["Payment"].astype(str).str.strip()
    assert set(pay["Payment"]) <= {"Yes", "No"} and pay["farmer_id"].is_unique
    pmap = dict(zip(pay["farmer_id"], pay["Payment"]))
    in_list = fid_raw.isin(pmap.keys())
    payment = fid_raw.map(lambda x: 1 if pmap.get(x) == "Yes" else 0)
    pos = list(merged.columns).index("farmers_id") + 1
    merged.insert(pos, "round", 2)
    merged.insert(pos + 1, "payment", payment.astype(int))
    merged.insert(pos + 2, "in_payment_list", in_list.astype(int))

    # ---- Section 3: activity costs, per-acre costs (÷ rice_acres_2026), outliers ----
    cst, cost_review, cost_labels = C.compute(merged)
    cst = cst.drop(columns=["cost_c_fert1_not_asked"])          # single form version: 1st-dose fertiliser asked of everyone
    cost_summary, cost_outl, cost_flags = C.outliers(merged, cst, cost_labels)
    cst["cpa_total_outlier"] = cost_flags["out_total"]
    cst["cost_outlier_n"] = cost_flags[[f"out_{k}" for k in C.ORDER]].sum(axis=1)
    anchor = list(merged.columns).index("f_other_paydate_2") + 1
    merged = pd.concat([merged.iloc[:, :anchor], cst, merged.iloc[:, anchor:]], axis=1)
    cost_cols = list(cst.columns)

    def cost_label(c):
        if c == "cost_n_missing":
            return "Number of activities with at least one cost item that could not be computed (999 code or quantity/price half-filled); such items count as 0 in totals"
        if c == "cpa_total_outlier":
            return "1 = total cost per acre is an outlier (|modified z| > 3.5 on log scale); see sheet cost_outliers"
        if c == "cost_outlier_n":
            return "Number of activities (of 27) whose cost per acre is flagged as an outlier for this household"
        if c in ("cost_total", "cpa_total"):
            base = "TOTAL cost of rice cultivation, Virippu 2026 (sum of Sections A–F)"
        elif c[4:7] == "sec" or c[5:8] == "sec":
            base = f"Section {c[-1]} total — {C.SECTIONS[c[-1]]} (sum of its activities)"
        else:
            lab, sec, how = cost_labels[c.split("_", 1)[1]]
            base = f"{sec}. {lab} = {how}"
        if c.startswith("cost_"):
            return f"Cost (Rs.): {base}"
        return f"Cost per acre (Rs./acre): {base}, ÷ rice_acres_2026 (missing if rice area is 0)"

    # ---- data checks ----
    r1 = pd.read_excel(ROUND1, sheet_name="data", usecols=["farmers_id"])["farmers_id"].astype(str).str.strip()
    key = lambda x: (0, int(x)) if x.isdigit() else (1, x)
    dups = merged[fid_raw.duplicated(keep=False)].assign(_f=fid_raw)
    pre = [
        ("Form versions", main["__version__"].nunique(), "; ".join(f"{v}: {n} interviews" for v, n in main["__version__"].value_counts().items())),
        ("Interviews / consent given", f"{n_main} / {int((main['consent'] == 1).sum())}",
         f"interview dates {main['interview_date'].min():%d %b %Y} to {main['interview_date'].max():%d %b %Y}"),
        ("Duplicate farmers_id (separate interviews; both kept)", dups["_f"].nunique(),
         "; ".join(f"{f}: _index {', '.join(map(str, g['_index']))} (interviewed {', '.join(g['interview_date'].dt.strftime('%d %b'))}; "
                   f"enumerator {', '.join(map(str, g['enumerator_id']))})" for f, g in dups.groupby("_f"))),
        ("Non-numeric farmers_id (kept as text)", int((~fid_raw.str.fullmatch(r"\d+")).sum()), ", ".join(fid_raw[~fid_raw.str.fullmatch(r"\d+")])),
        ("Payment file used", "", "RCT_payment_round1 (no Round 2 payment file); payment = 1 if 'Yes', 0 if 'No' or not listed"),
        ("Survey rows coded payment = 1 / 0", f"{int(payment.sum())} / {int((payment == 0).sum())}",
         f"of which in payment list with 'No': {int((in_list & (payment == 0)).sum())}; not in list: {int((~in_list).sum())}"),
        ("Payment-list IDs not found in Round 2", len(set(pmap) - set(fid_raw)),
         ", ".join(f"{m} ({pmap[m]})" for m in sorted(set(pmap) - set(fid_raw), key=key))),
        ("Round 2 IDs not in Round 1 (Round 1 after its ID corrections)", len(set(fid_raw) - set(r1)),
         ", ".join(sorted(set(fid_raw) - set(r1), key=key))),
        ("Round 1 IDs not in Round 2", len(set(r1) - set(fid_raw)), ", ".join(sorted(set(r1) - set(fid_raw), key=key))),
    ]
    empty = [c for c in main.columns if main[c].isna().all()]
    acres = pd.to_numeric(merged["rice_acres_2026"], errors="coerce")
    z_ac = merged.loc[~(acres > 0), "farmers_id"].astype(str).tolist()
    big = merged.loc[acres > 20, ["farmers_id"]].assign(a=acres[acres > 20])
    cost_checks = [
        ("Section 3 costs: rice_acres_2026 is 0 or missing → per-acre costs missing", len(z_ac), ", ".join(z_ac)),
        ("Section 3 costs: rice_acres_2026 > 20 acres (check)", len(big), ", ".join(f"{r.farmers_id} ({r.a:g} acres)" for r in big.itertuples())),
        ("Section 3 costs: households with total cost per acre flagged as outlier", int(cst["cpa_total_outlier"].sum()),
         ", ".join(merged.loc[cst["cpa_total_outlier"] == 1, "farmers_id"].astype(str))),
        ("Section 3 costs: harvesting (Section F)", "", "The questionnaire asks for anticipated quantity, rate and date where harvest had not happened by the interview (17–30 Sep 2026), so some Section F costs are expected rather than paid"),
    ]
    checks = pre + checks + [("Columns from main sheet that are completely empty (kept)", len(empty), ", ".join(empty))] + cost_checks

    # ---- codebook ----
    cb = []
    for i, c in enumerate(merged.columns, 1):
        if c in codebook_extra:
            src, sec, k, base = codebook_extra[c]
            cb.append([i, c, src, sec, k, f"{label_for(base)} [item {k}]"])
        elif c in cost_cols:
            cb.append([i, c, "Computed from Section 3", "Sec 3: cost of rice cultivation, Virippu 2026", None, cost_label(c)])
        elif c == "round":
            cb.append([i, c, "Added", "Survey round", None, "Survey round (2 for all rows in this file)"])
        elif c == "payment":
            cb.append([i, c, "RCT_payment_round1.xls", "RCT payment", None, "1 = Payment 'Yes' in RCT_payment_round1; 0 = 'No' or farmer not in the payment list"])
        elif c == "in_payment_list":
            cb.append([i, c, "RCT_payment_round1.xls", "RCT payment", None, "1 = farmer appears in RCT_payment_round1 (Yes or No); 0 = not listed"])
        else:
            cb.append([i, c, MAIN_SHEET, "Main sheet", None, label_for(c)])
    codebook = pd.DataFrame(cb, columns=["col_no", "variable", "source_sheet", "section", "item_no", "label"])
    checks_df = pd.DataFrame(checks, columns=["check", "count", "details"])

    with pd.ExcelWriter(OUT, engine="openpyxl", datetime_format="yyyy-mm-dd hh:mm", date_format="yyyy-mm-dd") as xw:
        merged.to_excel(xw, sheet_name="data", index=False)
        codebook.to_excel(xw, sheet_name="codebook", index=False)
        checks_df.to_excel(xw, sheet_name="data_checks", index=False)
        cost_summary.to_excel(xw, sheet_name="cost_summary", index=False, startrow=2)
        cost_outl.sort_values(["direction", "variable", "times the median"], ascending=[True, True, False]) \
            .to_excel(xw, sheet_name="cost_outliers", index=False, startrow=2)
        cost_review.to_excel(xw, sheet_name="cost_review", index=False, startrow=2)

    # ---- formatting ----
    date_only = {c for c in merged.columns if pd.api.types.is_datetime64_any_dtype(merged[c])
                 and (merged[c].dropna().dt.normalize() == merged[c].dropna()).all()}
    wb = load_workbook(OUT)
    base_font, head_font = Font(name="Arial", size=10), Font(name="Arial", size=10, bold=True)
    head_fill, new_fill = PatternFill("solid", start_color="DDE4EE"), PatternFill("solid", start_color="E8F1E4")
    cost_fill = PatternFill("solid", start_color="FCE9D9")
    added = set(codebook_extra) | {"round", "payment", "in_payment_list"}
    COST_SHEETS = {
        "cost_summary": ("Cost per acre (Rs./acre) by activity, Virippu 2026 — distribution and outlier counts",
                         "Statistics over households with cost > 0. Outlier = |modified z-score| > 3.5 on log(cost per acre) "
                         "(Iglewicz & Hoaglin), i.e. outside the lower/upper bounds shown. The Q3 + 3×IQR count is a second, raw-scale check."),
        "cost_outliers": ("Flagged per-acre costs, one row per household × activity",
                          "‘likely cause’ compares the entry with the median quantity and price for that activity. Values are as recorded; nothing has been corrected."),
        "cost_review": ("Decisions taken while computing costs (999 codes, half-filled entries, 'amount paid' read as a unit price)",
                        "Other-item questions ask the total amount paid; where amount ÷ quantity is implausibly low the amount is read as a unit price. "
                        "Ambiguous cases are kept as totals and listed here to check."),
    }
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.font = base_font
        hdr_row = 3 if ws.title in COST_SHEETS else 1
        for cell in ws[hdr_row]:
            cell.font = head_font
            cell.fill = (cost_fill if (ws.title == "data" and cell.value in cost_cols) else
                         new_fill if (ws.title == "data" and cell.value in added) else head_fill)
            cell.alignment = Alignment(vertical="center", wrap_text=ws.title in COST_SHEETS)
        ws.freeze_panes = "C2" if ws.title == "data" else ("A4" if ws.title in COST_SHEETS else "A2")
        if ws.title in COST_SHEETS:
            ws["A1"], ws["A2"] = COST_SHEETS[ws.title]
            ws["A1"].font = Font(name="Arial", size=12, bold=True)
            ws["A2"].font = Font(name="Arial", size=10, italic=True)
            ws.auto_filter.ref = f"A3:{get_column_letter(ws.max_column)}{ws.max_row}"
            ws.row_dimensions[3].height = 30
        else:
            ws.auto_filter.ref = ws.dimensions
    for name, widths in {"cost_summary": [8, 46, 18] + [11] * 14,
                         "cost_outliers": [10, 8, 8, 40, 18, 9, 10, 14, 14, 12, 10, 10, 110],
                         "cost_review": [10, 8, 11, 32, 26, 10, 18, 14, 14, 100]}.items():
        ws = wb[name]
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = w
        for row in ws.iter_rows(min_row=4):
            for cell in row:
                if isinstance(cell.value, float):
                    cell.number_format = "0.00" if ws.cell(3, cell.column).value in ("modified z (log)", "times the median", "rice_acres_2026") else "#,##0"
    ws = wb["data"]
    for j, c in enumerate(merged.columns, 1):
        ws.column_dimensions[get_column_letter(j)].width = min(max(len(str(c)) + 2, 10), 40)
        if c in date_only:
            for (cell,) in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                cell.number_format = "yyyy-mm-dd"
        if c in cost_cols and c not in ("cost_n_missing", "cpa_total_outlier", "cost_outlier_n"):
            for (cell,) in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                cell.number_format = "#,##0"
    for col, wdt in zip("ABCDEF", [8, 30, 26, 42, 9, 100]):
        wb["codebook"].column_dimensions[col].width = wdt
    dc = wb["data_checks"]
    for col, wdt in zip("ABC", [58, 12, 120]):
        dc.column_dimensions[col].width = wdt
    for (cell,) in dc.iter_rows(min_row=2, min_col=3, max_col=3):
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(OUT)
    print("final shape:", merged.shape, "| added columns:", len(added))
    pd.set_option("display.width", 250); pd.set_option("display.max_colwidth", 150)
    print(checks_df.to_string())


# =====================================================================================
# STEP 2 — item-level costs with dates, weekly aggregates  ->  items2.csv, timeline json
# =====================================================================================

def timeline_items2(DATA, OUT):
    """Item-level Section 3 expenditure with dates → weekly aggregates for the timeline page. (from timeline_data2.py)"""
    d = pd.read_excel(DATA, sheet_name="data")
    outl = pd.read_excel(DATA, sheet_name="cost_outliers", header=2)

    DATES = {
        "a_seed": ("a_seed_actdate", "a_seed_paydate"), "a_land": ("a_land_actdate", "a_land_paydate"),
        "b_labour": ("b_labour_actdate", "b_labour_paydate"), "b_fert": ("b_fert_actdate", "b_fert_paydate"),
        "c_mainlab": ("c_labour_actdate_land", "c_labour_paydate_land"), "c_mach": ("c_mach_actdate_land", "c_mach_paydate_land"),
        "c_translab": ("c_labour_actdate_2", "c_labour_paydate_2"), "c_fert1lab": ("c_labour_fert1_actdate", "c_labour_fert1_paydate"),
        "c_gapfill": ("c_labour_gapfilling_actdate", "c_labour_gapfilling_paydate"),
        "d_weed": ("d_weed_actdate", "d_weed_paydate"), "d_weedlab": ("d_weedlab_actdate", "d_weedlab_paydate"),
        "d_fert2lab": ("d_labour_fert2_actdate", "d_labour_fert2_paydate"),
        "e_pest": ("e_pest_actdate", "e_pest_paydate"), "e_pestlab": ("e_labour_pest_actdate", "e_labour_pest_paydate"),
        "e_fert3lab": ("e_labour_fert3_actdate", "e_labour_fert3_paydate"),
        "f_mach": ("f_mach_actdate", "f_mach_paydate"), "f_tractor": ("f_tractor_actdate", "f_tractor_paydate"),
        "f_labour": ("f_labour_actdate", "f_labour_paydate"),
    }

    # high outliers (entry errors) excluded per household × activity
    bad = set(zip(outl.loc[outl["direction"] == "high", "_index"], outl.loc[outl["direction"] == "high", "variable"].str[4:]))

    recs = []
    def add(key, label, sec, cost, act, pay):
        for i in cost.index:
            recs.append({"_index": d.at[i, "_index"], "farmers_id": d.at[i, "farmers_id"], "payment": d.at[i, "payment"],
                         "key": key, "label": label, "sec": sec, "cost": cost.at[i], "act": act.at[i], "pay": pay.at[i]})

    for key, label, sec, q, u, p in C.FIXED:
        c, _ = C.unit_price_cost(d[q], d[u], d[p])
        a, pdt = DATES[key]
        add(key, label, sec, c, pd.to_datetime(d[a], errors="coerce"), pd.to_datetime(d[pdt], errors="coerce"))
    for key, label, sec, pre, n in C.FERT:
        for k in range(1, n + 1):
            c, _ = C.unit_price_cost(d[f"{pre}_qty_{k}"], d[f"{pre}_unit_{k}"], d[f"{pre}_price_{k}"])
            add(key, label, sec, c, pd.to_datetime(d[f"{pre}_actdate_{k}"], errors="coerce"), pd.to_datetime(d[f"{pre}_paydate_{k}"], errors="coerce"))
    for key, label, sec, pre, n in C.OTHER:
        for k in range(1, n + 1):
            c, _ = C.other_cost(d[f"{pre}_qty_{k}"], d[f"{pre}_unit_{k}"], d[f"{pre}_amt_{k}"], d[f"{pre}_item_{k}"])
            add(key, label, sec, c, pd.to_datetime(d[f"{pre}_actdate_{k}"], errors="coerce"), pd.to_datetime(d[f"{pre}_paydate_{k}"], errors="coerce"))

    r = pd.DataFrame(recs)
    ivw = pd.to_datetime(d.set_index("_index")["interview_date"]).dt.normalize()
    r["interview"] = r["_index"].map(ivw)
    r = r[(r["cost"] > 0)]
    n_items = len(r)
    r["bad"] = [(i, k) in bad for i, k in zip(r["_index"], r["key"])]
    n_bad = int(r["bad"].sum()); bad_amt = float(r.loc[r["bad"], "cost"].sum())
    r = r[~r["bad"]]

    # dates: put every date on the 2026 season calendar (one stray 2025 date)
    log = {}
    for col in ("act", "pay"):
        s = r[col]
        log[f"{col}_year_counts"] = {str(int(k)): int(v) for k, v in s.dt.year.value_counts().items()}
        norm = pd.to_datetime({"year": 2026, "month": s.dt.month, "day": s.dt.day}, errors="coerce").where(s.notna())
        placeholder = (s.dt.month == 1) & (s.dt.day == 1)
        out_window = norm < pd.Timestamp("2026-02-01")
        log[f"{col}_dropped_placeholder_or_january"] = int((placeholder | out_window).sum())
        r[col + "_n"] = norm.where(~(placeholder | out_window))
        r[col + "_wk"] = r[col + "_n"] - pd.to_timedelta(r[col + "_n"].dt.weekday, unit="D")   # Monday of the week

    ORDER = C.ORDER
    LAB = {k: l for k, l, *_ in C.FIXED} | {k: l for k, l, *_ in C.FERT} | {k: l for k, l, *_ in C.OTHER}
    SEC = {k: s for k, _, s, *_ in C.FIXED} | {k: s for k, _, s, *_ in C.FERT} | {k: s for k, _, s, *_ in C.OTHER}

    out = {"activities": [{"key": k, "label": LAB[k], "sec": SEC[k]} for k in ORDER],
           "sections": [{"sec": s, "label": l} for s, l in C.SECTIONS.items()],
           "cells": {}, "meta": {}}
    for col, name in (("act", "activity"), ("pay", "payment")):
        g = r.dropna(subset=[col + "_wk"]).groupby(["key", col + "_wk"])
        cells = g.agg(farmers=("_index", "nunique"), total=("cost", "sum"), items=("cost", "size")).reset_index()
        cells["week"] = cells[col + "_wk"].dt.strftime("%Y-%m-%d")
        out["cells"][name] = cells[["key", "week", "farmers", "total", "items"]].round({"total": 0}).to_dict("records")
        out["meta"][name] = {"items": int(r[col + "_wk"].notna().sum()), "amount": float(r.loc[r[col + "_wk"].notna(), "cost"].sum()),
                             "households": int(r.loc[r[col + "_wk"].notna(), "_index"].nunique()),
                             "undated_items": int(r[col + "_wk"].isna().sum()), "undated_amount": float(r.loc[r[col + "_wk"].isna(), "cost"].sum())}
        sw = r.dropna(subset=[col + "_wk"]).groupby(["sec", col + "_wk"]).agg(farmers=("_index", "nunique"), total=("cost", "sum")).reset_index()
        sw["week"] = sw[col + "_wk"].dt.strftime("%Y-%m-%d")
        out.setdefault("secweeks", {})[name] = sw[["sec", "week", "farmers", "total"]].round({"total": 0}).to_dict("records")
    r["anticipated"] = (r["act_n"] > r["interview"]).astype(int)
    out["meta"]["anticipated_share"] = float(r.loc[r["anticipated"] == 1, "cost"].sum() / r.loc[r["act_n"].notna(), "cost"].sum())
    out["meta"]["interviews"] = [str(ivw.min().date()), str(ivw.max().date())]
    r.drop(columns=["bad"]).to_csv(os.path.join(os.path.dirname(OUT), "items2.csv"), index=False)
    out["meta"].update({"items_with_cost": n_items, "excluded_outlier_items": n_bad, "excluded_outlier_amount": bad_amt,
                        "households_total": int(d.shape[0]), **log,
                        "households_all_dates_2025": int((pd.concat([pd.to_datetime(d[c], errors='coerce').dt.year for c in d.columns if c.endswith(('_actdate', '_actdate_land')) or '_actdate_' in c], axis=1)
                                                          .apply(lambda row: (row.dropna() == 2025).all() and row.notna().any(), axis=1)).sum())})
    json.dump(out, open(OUT, "w"), default=str)
    print(json.dumps(out["meta"], indent=1, default=str))
    for name in ("activity", "payment"):
        c = pd.DataFrame(out["cells"][name])
        print(name, "cells", len(c), "weeks", c["week"].nunique(), c["week"].min(), c["week"].max(), "max farmers", c["farmers"].max(), "max total", c["total"].max())


# =====================================================================================
# STEP 3 — per-acre cost components, treatment vs control  ->  Round2_cost_comparison_payment.xlsx
# =====================================================================================

def cost_comparison2(DATA, OUT, ITEMS):
    """Round 2: compare Section 3 cost components (per acre, Virippu 2026) between payment = 1 and payment = 0, incl. timing relative to the payments. (from cost_compare2.py)"""
    d = pd.read_excel(DATA, sheet_name="data").copy()
    outl = pd.read_excel(DATA, sheet_name="cost_outliers", header=2)
    T = d["payment"] == 1
    dup = d["farmers_id"].astype(str).duplicated(keep=False)        # farmer IDs with two interviews (both kept)
    old = dup
    acres = pd.to_numeric(d["rice_acres_2026"], errors="coerce")

    # one form version in Round 2: totals are complete, the *_x names are kept only for the code below
    d["cost_secC_x"] = d["cost_secC"]
    d["cost_total_x"] = d["cost_total"]

    # timing relative to the RCT payments (1st week of August, 1st week of September 2026), from dated items
    it = pd.read_csv(ITEMS, parse_dates=["act_n"])
    it = it.dropna(subset=["act_n"])
    per = pd.cut(it["act_n"], [pd.Timestamp("2000-01-01"), pd.Timestamp("2026-08-01"), pd.Timestamp("2026-09-01"), pd.Timestamp("2100-01-01")],
                 right=False, labels=["pre", "aug", "sep"])
    tim = it.assign(per=per).pivot_table(index="_index", columns="per", values="cost", aggfunc="sum", fill_value=0, observed=False)
    tim = tim.reindex(d["_index"]).fillna(0)
    for p in ("pre", "aug", "sep"):
        d[f"cost_t_{p}"] = tim[p].values
        d[f"cpa_t_{p}"] = (d[f"cost_t_{p}"] / acres).where(acres > 0)
    dated_tot = tim[["pre", "aug", "sep"]].sum(axis=1).values
    d["cpa_t_share_after"] = pd.Series((tim["aug"] + tim["sep"]).values / np.where(dated_tot > 0, dated_tot, np.nan), index=d.index).where(acres > 0)
    for c in ("secC_x", "total_x"):
        d[f"cpa_{c}"] = (d[f"cost_{c}"] / acres).where(acres > 0)

    # outlier mask: (row, variable) pairs flagged in the cost_outliers sheet; derived totals get their own check
    flag = pd.DataFrame(False, index=d.index, columns=[c for c in d.columns if c.startswith("cpa_")])
    idx_of = dict(zip(d["_index"], d.index))
    for r in outl.itertuples():
        flag.at[idx_of[r._2], r.variable] = True
    flag["cpa_secC_x"] = flag["cpa_secC"]
    flag["cpa_total_x"] = flag["cpa_total"]
    for c in ("cpa_t_pre", "cpa_t_aug", "cpa_t_sep", "cpa_t_share_after"):
        flag[c] = False                        # items flagged as entry errors are already left out of the dated items

    LAB = {k: v for k, v in [(k, l) for k, l, *_ in C.FIXED] + [(k, l) for k, l, *_ in C.FERT] + [(k, l) for k, l, *_ in C.OTHER]}
    PANELS = [(f"{s}. {C.SECTIONS[s]}", [k for k in C.ORDER if k[0] == s.lower()]) for s in C.SECTIONS]
    TOTALS = [("secA", "Section A: Pre-nursery preparation"), ("secB", "Section B: Nursery preparation"),
              ("secC", "Section C: Transplanting"),
              ("secD", "Section D: Weedicide / fertiliser application"), ("secE", "Section E: Pesticide / fertiliser application"),
              ("secF", "Section F: Harvesting"),
              ("total", "TOTAL cost of cultivation")]


    def stars(p):
        return "" if pd.isna(p) else "***" if p < 0.01 else "**" if p < 0.05 else "*" if p < 0.10 else ""


    def compare(s, incidence=True):
        t, c = s[T].dropna(), s[~T].dropna()
        tt = stats.ttest_ind(t, c, equal_var=False)
        mw = stats.mannwhitneyu(t, c, alternative="two-sided").pvalue
        sd_pool = np.sqrt((t.var(ddof=1) + c.var(ddof=1)) / 2)
        out = {"N_T": len(t), "Mean_T": t.mean(), "SD_T": t.std(ddof=1), "Median_T": t.median(),
               "N_C": len(c), "Mean_C": c.mean(), "SD_C": c.std(ddof=1), "Median_C": c.median(),
               "Diff": t.mean() - c.mean(), "p_t": tt.pvalue, "p_mw": mw,
               "NormDiff": (t.mean() - c.mean()) / sd_pool if sd_pool > 0 else np.nan}
        if incidence:
            out["Inc_T"], out["Inc_C"] = (t > 0).mean(), (c > 0).mean()
            tab = np.array([[(t > 0).sum(), (t == 0).sum()], [(c > 0).sum(), (c == 0).sum()]])
            out["p_inc"] = stats.fisher_exact(tab)[1] if tab.min() >= 0 and (tab[:, 1].sum() > 0) else np.nan
        out["Stars"] = stars(out["p_t"])
        return out


    rows = []
    for panel, keys in PANELS:
        rows.append({"Variable": panel, "_panel": True})
        for k in keys:
            s = d[f"cpa_{k}"].where(~flag[f"cpa_{k}"])
            note = " †" if k == "c_fert1" else ""
            rows.append({"Variable": LAB[k], "var": f"cpa_{k}", "_panel": False, "n_excl": int(flag[f"cpa_{k}"].sum()), **compare(s)})
    rows.append({"Variable": "Section totals and total cost", "_panel": True})
    for k, lab in TOTALS:
        s = d[f"cpa_{k}"].where(~flag[f"cpa_{k}"])
        rows.append({"Variable": lab, "var": f"cpa_{k}", "_panel": False, "n_excl": int(flag[f"cpa_{k}"].sum()), **compare(s, incidence=False)})
    rows.append({"Variable": "Timing relative to the RCT payments (dated items)", "_panel": True})
    for k, lab in [("t_pre", "Spent before the 1st payment (before 1 Aug)"), ("t_aug", "Spent after the 1st payment, in August"),
                   ("t_sep", "Spent after the 2nd payment (from 1 Sep; harvest partly anticipated)")]:
        rows.append({"Variable": lab, "var": f"cpa_{k}", "_panel": False, "n_excl": 0, **compare(d[f"cpa_{k}"], incidence=True)})
    rows.append({"Variable": "Share of season cost spent after the 1st payment", "var": "cpa_t_share_after", "_panel": False, "n_excl": 0,
                 **compare(d["cpa_t_share_after"], incidence=False)})
    rows.append({"Variable": "Context", "_panel": True})
    rows.append({"Variable": "Rice area, Virippu 2026 (acres)", "var": "rice_acres_2026", "_panel": False, "n_excl": 0,
                 **compare(acres.where(acres > 0), incidence=False)})
    s_tot = d["cost_total_x"].where(~flag["cpa_total_x"] & (acres > 0))
    rows.append({"Variable": "Total cost (Rs., not per acre)", "var": "cost_total_x", "_panel": False,
                 "n_excl": int(flag["cpa_total_x"].sum()), **compare(s_tot, incidence=False)})
    res = pd.DataFrame(rows)

    # ---------------- cost shares (composition) ----------------
    ok_share = ~flag["cpa_total_x"] & ~flag[[f"cpa_sec{s}" for s in "ABDEF"] + ["cpa_secC_x"]].any(axis=1) & (d["cost_total_x"] > 0)
    share_rows = []
    for k, lab in [("secA", "A. Pre-nursery preparation"), ("secB", "B. Nursery preparation"), ("secC_x", "C. Transplanting"),
                   ("secD", "D. Weedicide / fertiliser application"), ("secE", "E. Pesticide / fertiliser application"), ("secF", "F. Harvesting")]:
        sh = (d[f"cost_{k}"] / d["cost_total_x"]).where(ok_share)
        r = compare(sh, incidence=False)
        share_rows.append({"Section": lab, **r})
    # labour / material / machinery split
    GROUPS = {
        "Hired labour": ["b_labour", "c_mainlab", "c_translab", "c_fert1lab", "c_gapfill", "d_weedlab", "d_fert2lab", "e_pestlab", "e_fert3lab", "f_labour"],
        "Machinery & transport": ["a_land", "c_mach", "f_mach", "f_tractor"],
        "Seed, fertiliser & plant protection": ["a_seed", "b_fert", "d_weed", "d_fert2", "e_pest", "e_fert3"],
        "Other items (mixed)": ["a_other", "b_other", "c_other", "d_other", "e_other", "f_other"],
    }
    for g, ks in GROUPS.items():
        sh = (d[[f"cost_{k}" for k in ks]].sum(axis=1) / d["cost_total_x"]).where(ok_share)
        share_rows.append({"Section": g, **compare(sh, incidence=False)})
    shares = pd.DataFrame(share_rows)

    # ---------------- robustness ----------------
    rob = []
    for k, lab in [(k, LAB[k]) for k in C.ORDER] + TOTALS + [("t_pre", "Spent before 1 Aug"), ("t_aug", "Spent in August"), ("t_sep", "Spent from 1 Sep")]:
        full = d[f"cpa_{k}"]
        newf = full.where(~old & ~flag[f"cpa_{k}"])
        a = compare(full, incidence=False); c3 = compare(newf, incidence=False)
        rob.append({"Variable": lab, "var": f"cpa_{k}",
                    "All obs: Mean T": a["Mean_T"], "All obs: Mean C": a["Mean_C"], "All obs: Median T": a["Median_T"], "All obs: Median C": a["Median_C"],
                    "All obs: p (t)": a["p_t"], "All obs: p (MW)": a["p_mw"],
                    "Excl. duplicate IDs: N T/C": f"{c3['N_T']}/{c3['N_C']}", "Excl. duplicate IDs: Diff": c3["Diff"], "Excl. duplicate IDs: p (t)": c3["p_t"]})
    rob = pd.DataFrame(rob)

    # ---------------- write ----------------
    nT, nC = int(T.sum()), int((~T).sum())
    hdr = ["Cost per acre (Rs./acre)", "N", "Mean", "SD", "Median", "% incurring", "N", "Mean", "SD", "Median", "% incurring",
           "Difference (T − C)", "", "p-value (t-test)", "p-value (Mann–Whitney)", "Normalized difference", "p-value (% incurring)", "Outliers excluded"]
    out_rows = []
    for _, r in res.iterrows():
        if r["_panel"]:
            out_rows.append([r["Variable"]] + [None] * 17)
        else:
            out_rows.append([r["Variable"], r["N_T"], r["Mean_T"], r["SD_T"], r["Median_T"], r.get("Inc_T"),
                             r["N_C"], r["Mean_C"], r["SD_C"], r["Median_C"], r.get("Inc_C"),
                             r["Diff"], r["Stars"], r["p_t"], r["p_mw"], r["NormDiff"], r.get("p_inc"), r["n_excl"]])
    sh_hdr = ["Share of total cost", "N T", "Mean share T", "N C", "Mean share C", "Difference (T − C)", "", "p-value (t-test)", "p-value (Mann–Whitney)"]
    sh_rows = [[r["Section"], r["N_T"], r["Mean_T"], r["N_C"], r["Mean_C"], r["Diff"], r["Stars"], r["p_t"], r["p_mw"]] for _, r in shares.iterrows()]

    act = res[(~res["_panel"]) & res["var"].astype(str).str.match(r"cpa_[a-f]_")]
    n10 = int((act["p_t"] < 0.10).sum())
    notes = [
        ["Groups", f"Treatment: payment = 1 (N = {nT}). Control: payment = 0 (N = {nC}). Payment status from RCT_payment_round1 (no Round 2 payment file)."],
        ["Timing", "Costs refer to Virippu 2026. The RCT payments were made in the 1st week of August and the 1st week of September 2026, so most of the "
                   "season's spending (seed to 2nd fertiliser dose) came before the 1st payment. The 'Timing' rows split each household's dated costs "
                   "into before 1 Aug, August, and from 1 Sep; only the later two can respond to the payments."],
        ["Anticipated costs", "Harvest costs not yet incurred at the interview (17–30 Sep 2026) were reported as expected amounts, so the 'from 1 Sep' row is "
                              "partly anticipated spending."],
        ["Measure", "Cost per acre = activity cost ÷ rice_acres_2026, from Survey_round_2_merged_wide.xlsx. Households that did not incur a cost are "
                    "included as 0. Households with no rice area have no per-acre values."],
        ["Outliers", "Values flagged in the cost_outliers sheet (|modified z| > 3.5 on log scale) are excluded from that row only; the count is in the last column. "
                     "Timing rows use dated items with flagged entry errors already left out."],
        ["Tests", "Welch two-sample t-test on means; Mann–Whitney rank-sum test (robust to skew); normalized difference = diff ÷ √[(SD²_T + SD²_C)/2]; "
                  "Fisher's exact test for the share of households incurring the cost. *** p<0.01, ** p<0.05, * p<0.10."],
        ["Duplicate IDs", f"Farmer IDs 136, 208, 244 and 257 each have two interviews ({int(dup.sum())} rows); both are kept in the main table and dropped in the robustness sheet."],
        ["Multiple comparisons", f"27 activities are tested; at the 10% level about 3 significant differences are expected by chance. Observed: {n10}."],
        ["Shares", "Share of each section in the household's total cost, for households with no flagged total or section."],
    ]

    with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
        pd.DataFrame(out_rows, columns=hdr).to_excel(xw, sheet_name="cost_comparison", index=False, startrow=3)
        pd.DataFrame(sh_rows, columns=sh_hdr).to_excel(xw, sheet_name="cost_shares", index=False, startrow=2)
        rob.drop(columns="var").to_excel(xw, sheet_name="robustness", index=False, startrow=2)
        pd.DataFrame(notes, columns=["Item", "Note"]).to_excel(xw, sheet_name="notes", index=False)

    wb = load_workbook(OUT)
    F = lambda **k: Font(**{"name": "Arial", "size": 10, **k})
    thin = Side(style="thin", color="000000")
    hfill = PatternFill("solid", start_color="DDE4EE"); pfill = PatternFill("solid", start_color="F2F2F2")
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.font = F()

    ws = wb["cost_comparison"]
    ws["A1"] = "Table: Cost of rice cultivation per acre (Virippu 2026, Survey Round 2) by payment status"
    ws["A1"].font = F(bold=True, size=12)
    ws["A2"] = (f"Treatment = payment received (N = {nT}); Control = no payment (N = {nC}). Flagged outliers excluded. "
                "Welch t-test; *** p<0.01, ** p<0.05, * p<0.10. See 'notes'.")
    ws["A2"].font = F(italic=True)
    ws["B3"] = "Treatment (payment = 1)"; ws.merge_cells("B3:F3")
    ws["G3"] = "Control (payment = 0)"; ws.merge_cells("G3:K3")
    ws["L3"] = "Comparison"; ws.merge_cells("L3:R3")
    ws.merge_cells("L4:M4")
    for cell in list(ws[3]) + list(ws[4]):
        cell.font = F(bold=True); cell.fill = hfill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["A4"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[4].height = 42
    for c in range(1, 19):
        ws.cell(3, c).border = Border(top=thin); ws.cell(4, c).border = Border(bottom=thin)
    for i, (_, r) in enumerate(res.iterrows()):
        rr = 5 + i
        if r["_panel"]:
            ws.cell(rr, 1).font = F(bold=True)
            for c in range(1, 19):
                ws.cell(rr, c).fill = pfill
            continue
        ws.cell(rr, 1).alignment = Alignment(indent=1)
        acre_row = r["var"] == "rice_acres_2026"
        share_row = r["var"] == "cpa_t_share_after"
        for c in (3, 4, 5, 8, 9, 10, 12):
            ws.cell(rr, c).number_format = "0.0%" if share_row else "0.00" if acre_row else "#,##0"
        for c in (6, 11):
            ws.cell(rr, c).number_format = "0%"
        for c in (14, 15, 16, 17):
            ws.cell(rr, c).number_format = "0.000"
        if r["p_t"] < 0.10:
            for c in (12, 13, 14):
                ws.cell(rr, c).font = F(bold=True)
        if pd.notna(r.get("p_mw")) and r["p_mw"] < 0.10:
            ws.cell(rr, 15).font = F(bold=True)
        if pd.notna(r.get("p_inc")) and r["p_inc"] < 0.10:
            ws.cell(rr, 17).font = F(bold=True)
        if r["var"] in ("cpa_total",):
            for c in range(1, 19):
                ws.cell(rr, c).font = F(bold=True)
    last = 5 + len(res) - 1
    for c in range(1, 19):
        ws.cell(last, c).border = Border(bottom=thin)
    ws.cell(last + 2, 1, f"Activities with a significant difference in means at 10%: {n10} of 27 (about 3 expected by chance). "
                         "Bold p-values are below 0.10.").font = F(italic=True)
    for j, w in enumerate([68, 5, 10, 10, 10, 9, 5, 10, 10, 10, 9, 12, 4, 10, 12, 11, 11, 9], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "B5"
    ws.page_setup.orientation = "landscape"; ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0; ws.print_title_rows = "3:4"

    ws = wb["cost_shares"]
    ws["A1"] = "Composition of cultivation cost: mean share of total cost by section and by input type"
    ws["A1"].font = F(bold=True, size=12)
    for cell in ws[3]:
        cell.font = F(bold=True); cell.fill = hfill; cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[3].height = 32
    for row in ws.iter_rows(min_row=4):
        for c in (3, 5, 6):
            row[c - 1].number_format = "0.0%"
        for c in (8, 9):
            row[c - 1].number_format = "0.000"
    ws.insert_rows(10)
    ws.cell(10, 1, "By input type").font = F(bold=True)
    for j, w in enumerate([46, 6, 13, 6, 13, 13, 4, 11, 13], 1):
        ws.column_dimensions[get_column_letter(j)].width = w

    ws = wb["robustness"]
    ws["A1"] = "Robustness: all observations incl. outliers (compare medians and Mann–Whitney), and excluding the 8 interviews with duplicate farmer IDs (outliers excluded)"
    ws["A1"].font = F(bold=True, size=12)
    for cell in ws[3]:
        cell.font = F(bold=True); cell.fill = hfill; cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[3].height = 45
    for row in ws.iter_rows(min_row=4):
        for cell in row[1:]:
            if isinstance(cell.value, float):
                cell.number_format = "0.000" if "p (" in str(ws.cell(3, cell.column).value) else "#,##0"
    ws.column_dimensions["A"].width = 58
    for j in range(2, 11):
        ws.column_dimensions[get_column_letter(j)].width = 12

    ws = wb["notes"]
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    for cell in ws[1]:
        cell.font = F(bold=True); cell.fill = hfill
    ws.column_dimensions["A"].width = 22; ws.column_dimensions["B"].width = 125
    wb.save(OUT)

    pd.set_option("display.width", 260); pd.set_option("display.max_rows", 100)
    print(res[~res["_panel"]][["Variable", "N_T", "Mean_T", "Median_T", "N_C", "Mean_C", "Median_C", "Diff", "p_t", "p_mw", "Inc_T", "Inc_C", "p_inc", "n_excl"]].round(3).to_string(index=False))
    print(shares[["Section", "Mean_T", "Mean_C", "Diff", "p_t", "p_mw"]].round(3).to_string(index=False))
    print(rob[rob["var"].isin(["cpa_total", "cpa_t_pre", "cpa_t_aug", "cpa_t_sep"])].round(3).to_string(index=False))
    print("n10 activities:", n10)


# =====================================================================================
# STEP 4 — bubble plot of weekly cost per acre  ->  Round2_cost_per_acre_bubble_*.jpg
# =====================================================================================

def bubble_plot2(ITEMS, DATA, OUT, BASIS="act", MODE="overlay"):
    """Round 2 (Virippu 2026) static bubble plot: weekly cost per acre by activity, with the cumulative cost per acre of the average farmer. (from bubble2.py)"""
    it = pd.read_csv(ITEMS, parse_dates=[f"{BASIS}_wk"])
    hh = pd.read_excel(DATA, sheet_name="data", usecols=["_index", "rice_acres_2026"])
    acres = hh.set_index("_index")["rice_acres_2026"]
    valid = acres[acres > 0]
    it = it[it["_index"].isin(valid.index)].dropna(subset=[f"{BASIS}_wk"]).copy()
    it["cpa"] = it["cost"] / it["_index"].map(acres)
    N_HH = len(valid)
    # RCT cash transfers: 1st payment in the first week of August, 2nd payment in the first week of September 2026
    PAYWEEKS = [(pd.Timestamp("2026-08-01"), "1st payment\n1st week of Aug"), (pd.Timestamp("2026-09-01"), "2nd payment\n1st week of Sep")]
    dated = it[it[f"{BASIS}_wk"].notna()]
    ANTIC = dated.loc[dated["anticipated"] == 1, "cpa"].sum() / dated["cpa"].sum()


    # bubbles: per farmer, per activity, per week → cost per acre; then average over the farmers who spent that week
    fw = it.groupby(["key", "sec", "label", "_index", f"{BASIS}_wk"], as_index=False)["cpa"].sum()
    bub = fw.groupby(["key", "sec", "label", f"{BASIS}_wk"]).agg(farmers=("_index", "nunique"), cpa=("cpa", "mean")).reset_index()
    bub["x"] = bub[f"{BASIS}_wk"] + pd.Timedelta(days=3.5)

    # cumulative line: average over all households with rice area of cumulative cost per acre
    wk = it.groupby(f"{BASIS}_wk")["cpa"].sum().sort_index() / N_HH
    weeks = pd.date_range("2026-01-26", "2026-12-28", freq="W-MON")
    cum = wk.reindex(weeks, fill_value=0).cumsum()
    cum = cum[(cum.index >= wk.index.min() - pd.Timedelta(days=7)) & (cum.index <= wk.index.max())]
    cx = cum.index + pd.Timedelta(days=7)          # value reached by the end of the week
    end_val = cum.iloc[-1]
    half_wk = cum.index[np.argmax(cum.values >= end_val / 2)]

    # ---------------- figure ----------------
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10.5})
    INK, INK2, MUTED, GRID, AXIS, SURF = "#141613", "#4d524a", "#868a80", "#e2e4dc", "#c3c6ba", "#fcfcfb"
    PAYCOL = "#6250d6"          # violet: not used by any stage colour
    COL = {"A": "#2a78d6", "B": "#eb6834", "C": "#1baf7a", "D": "#eda100", "E": "#e87ba4", "F": "#008300"}
    SECLAB = {"A": "A. Pre-nursery preparation", "B": "B. Nursery preparation", "C": "C. Transplanting",
              "D": "D. Weedicide / 2nd fertiliser dose", "E": "E. Pesticide / 3rd fertiliser dose", "F": "F. Harvesting"}
    basis_word = "activity" if BASIS == "act" else "payment"
    inr = lambda v: f"₹{v:,.0f}"

    def style(ax, xlabel=True):
        ax.set_facecolor(SURF)
        ax.set_xlim(pd.Timestamp("2026-02-01"), pd.Timestamp("2026-12-31"))
        for p0, _ in PAYWEEKS:
            ax.axvspan(p0, p0 + pd.Timedelta(days=7), color=PAYCOL, alpha=0.14, lw=0, zorder=1)
        ax.xaxis.set_major_locator(mdates.MonthLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b"))
        ax.grid(axis="y", color=GRID, lw=0.8, zorder=0); ax.grid(axis="x", color=GRID, lw=0.6, zorder=0)
        for sp in ("top", "right", "left"):
            ax.spines[sp].set_visible(False)
        ax.spines["bottom"].set_color(AXIS)
        ax.tick_params(colors=MUTED, length=0, labelsize=10)
        if xlabel:
            ax.set_xlabel(f"Week of {basis_word}, 2026", color=INK2)

    panels = MODE == "panels"
    if panels:
        fig, (ax, ax2) = plt.subplots(2, 1, figsize=(13, 9.4), dpi=200, sharex=True, gridspec_kw={"height_ratios": [3, 1.35], "hspace": 0.12})
        fig.subplots_adjust(left=0.085, right=0.78, top=0.865, bottom=0.135)
    else:
        fig, ax = plt.subplots(figsize=(13, 7.6), dpi=200)
        fig.subplots_adjust(left=0.085, right=0.78, top=0.82, bottom=0.185)
    fig.patch.set_facecolor("white")

    maxf = bub["farmers"].max()
    size = lambda f: (np.sqrt(f / maxf) * 26) ** 2          # marker area proportional to farmers
    b = bub.sort_values("farmers", ascending=False)
    ax.scatter(b["x"], b["cpa"], s=size(b["farmers"]), c=b["sec"].map(COL), alpha=0.78, edgecolors="white", linewidths=0.7, zorder=3)
    ax.set_yscale("log")
    yt = [10, 50, 100, 500, 1000, 5000, 10000, 50000]
    lo = max(5, bub["cpa"].min() * 0.6)
    hi = (bub["cpa"].max() if panels else max(end_val, bub["cpa"].max())) * 1.9
    ax.set_ylim(lo, hi)
    ax.yaxis.set_major_locator(FixedLocator([t for t in yt if lo <= t <= hi]))
    ax.yaxis.set_minor_locator(NullLocator())
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: inr(v)))
    style(ax, xlabel=not panels)
    ax.set_ylabel("Cost per acre (₹, log scale)", color=INK2)

    for (p0, lab), side in zip(PAYWEEKS, ("right", "left")):      # 1st label left of its band, 2nd label right of its band
        xt = p0 - pd.Timedelta(days=2) if side == "right" else p0 + pd.Timedelta(days=9)
        ax.text(xt, hi / 1.15, lab, ha=side, va="top", fontsize=8.8, color=PAYCOL, fontweight="bold",
                bbox=dict(boxstyle="round,pad=0.25", fc="white", ec="none", alpha=0.85))
    hx = half_wk + pd.Timedelta(days=7); hy = cum.loc[half_wk]
    half_txt = f"Half the season's cost\nspent by {hx - pd.Timedelta(days=1):%-d %b}"
    end_txt = f"{inr(end_val)} per acre\nby season end"
    if panels:
        ax2.fill_between(cx, 0, cum.values, color=INK, alpha=0.07, lw=0, zorder=2)
        ax2.plot(cx, cum.values, color=INK, lw=2.2, zorder=4, solid_capstyle="round")
        ax2.set_ylim(0, end_val * 1.18)
        ax2.yaxis.set_major_formatter(FuncFormatter(lambda v, _: inr(v)))
        ax2.set_yticks([0, 10000, 20000, 30000, 40000])
        style(ax2)
        ax2.set_ylabel("Cumulative cost\nper acre (₹)", color=INK2)
        tgt = ax2
    else:
        m = cum.values >= lo
        ax.plot(cx[m], cum.values[m], color=INK, lw=2.2, zorder=4, solid_capstyle="round")
        tgt = ax
    tgt.scatter([cx[-1]], [end_val], s=28, color=INK, zorder=5)
    tgt.annotate(end_txt, (cx[-1], end_val), xytext=(-8, -22) if panels else (-10, -10), textcoords="offset points",
                 va="top", ha="right", fontsize=10, color=INK, fontweight="bold")
    tgt.scatter([hx], [hy], s=26, color=INK, zorder=5)
    tgt.annotate(half_txt, (hx, hy), xytext=(12, -4), textcoords="offset points", ha="left", va="top", fontsize=9.5, color=INK2)

    title = "Rice cultivation cost per acre by week and activity, Virippu 2026 (Survey Round 2)"
    sub = ("Bubbles: average cost per acre among farmers who spent on that activity in the week; bubble size is the number of farmers. "
           + ("Lower panel" if panels else "Black line") + f": cumulative cost per acre of the average farmer ({N_HH} households).")
    fig.text(0.085, 0.955 if panels else 0.945, title, fontsize=16, fontweight="bold", color=INK)
    fig.text(0.085, 0.94 if panels else 0.928, "\n".join(textwrap.wrap(sub, 150)), fontsize=10.5, color=INK2, va="top")

    h_sec = [Line2D([0], [0], marker="o", ls="", markersize=9, markerfacecolor=COL[k], markeredgecolor="white", label=SECLAB[k]) for k in COL]
    top = 0.865 if panels else 0.82
    leg1 = fig.legend(handles=h_sec, title="Stage of cultivation", loc="upper left", bbox_to_anchor=(0.79, top), frameon=False,
                      fontsize=9.5, title_fontsize=10, alignment="left")
    h_sz = [plt.scatter([], [], s=size(v), facecolor="none", edgecolor=MUTED, lw=1, label=f"{v} farmers") for v in (10, 50, 100)]
    leg2 = fig.legend(handles=h_sz, title="Farmers in the week", loc="upper left", bbox_to_anchor=(0.79, top - (0.25 if panels else 0.30)), frameon=False,
                      fontsize=9.5, title_fontsize=10, labelspacing=1.5, borderpad=0.8, alignment="left")
    h3 = ([] if panels else [Line2D([0], [0], color=INK, lw=2.2, label="Cumulative cost per acre,\naverage farmer")]) + \
         [plt.Rectangle((0, 0), 1, 1, fc=PAYCOL, alpha=0.25, ec="none", label="Week of an RCT cash\ntransfer payment")]
    fig.legend(handles=h3, loc="upper left", bbox_to_anchor=(0.79, top - (0.47 if panels else 0.56)), frameon=False, fontsize=9.5)
    for lg in (leg1, leg2):
        lg.get_title().set_color(INK)

    note = (f"Source: Survey Round 2, Section 3 (Cash transfer RCT); {N_HH} households with rice area > 0 (rice_acres_2026). Cost = quantity × price per unit; "
            f"'other items' = amount paid. Weeks start Monday; {basis_word} dates. Shaded weeks: RCT cash transfer payments (1st week of August and 1st week of September 2026). "
            f"Harvest costs not yet incurred at the interview were reported as expected amounts (about {ANTIC:.0%} of the season's cost per acre). "
            "Items flagged as entry errors and items without a date are excluded." + ("" if panels else " The y-axis is logarithmic, which compresses the later rise of the cumulative line."))
    fig.text(0.085, 0.025, "\n".join(textwrap.wrap(note, 180)), fontsize=8.3, color=MUTED, va="bottom")
    fig.savefig(OUT, dpi=200, pil_kwargs={"quality": 95})
    print(f"{MODE}: end {end_val:,.0f}; half by week of {half_wk:%Y-%m-%d}; bubbles {len(bub)}; households {N_HH}; anticipated {ANTIC:.1%}")


# =====================================================================================
# STEP 5 — pesticide vs 3rd fertiliser dose, Round 2  ->  Round2_pesticide_vs_3rd_fertiliser.jpg
# =====================================================================================

def pesticide_fert3(ITEMS, DATA, OUT):
    """Round 2: pesticide vs 3rd fertiliser dose — weekly spending per acre and spending before / after the 1st payment, treatment vs control. (from pest_fert3.py)"""
    hh = pd.read_excel(DATA, sheet_name="data", usecols=["_index", "payment", "rice_acres_2026"])
    hh = hh[hh["rice_acres_2026"] > 0].set_index("_index")
    it = pd.read_csv(ITEMS, parse_dates=["act_n", "act_wk"])
    it = it[it["_index"].isin(hh.index) & it["key"].isin(["e_pest", "e_fert3"])].dropna(subset=["act_n"]).copy()
    it["cpa"] = it["cost"] / it["_index"].map(hh["rice_acres_2026"])
    N = hh["payment"].value_counts()                      # households with rice area, by group
    PAY1, PAY2 = pd.Timestamp("2026-08-01"), pd.Timestamp("2026-09-01")

    INK, INK2, MUTED, GRID, AXIS, SURF = "#141613", "#4d524a", "#868a80", "#e2e4dc", "#c3c6ba", "#fcfcfb"
    COL = {1: "#2a78d6", 0: "#eb6834"}                    # treatment blue, control orange
    GLAB = {1: f"Treatment (payment = 1, n = {N[1]})", 0: f"Control (payment = 0, n = {N[0]})"}
    ACT = {"e_pest": "Pesticide", "e_fert3": "3rd fertiliser dose"}
    inr = lambda v, _=None: f"₹{v:,.0f}"
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10.5})

    fig, axes = plt.subplots(2, 2, figsize=(13, 9.2), dpi=200, gridspec_kw={"height_ratios": [1.25, 1], "hspace": 0.42, "wspace": 0.16})
    fig.patch.set_facecolor("white")
    fig.subplots_adjust(left=0.07, right=0.98, top=0.82, bottom=0.12)
    weeks = pd.date_range("2026-06-15", "2026-10-05", freq="W-MON")
    stats_out = {}

    for j, key in enumerate(["e_pest", "e_fert3"]):
        x = it[it["key"] == key]
        # ---- top: weekly cost per acre, average over all households in the group (non-buyers count as 0)
        ax = axes[0, j]
        ax.set_facecolor(SURF)
        for p0 in (PAY1, PAY2):
            ax.axvspan(p0, p0 + pd.Timedelta(days=7), color=INK, alpha=0.08, lw=0, zorder=1)
        ymax = 0
        for g in (1, 0):
            w = x[x["payment"] == g].groupby("act_wk")["cpa"].sum().reindex(weeks, fill_value=0) / N[g]
            ax.plot(w.index + pd.Timedelta(days=3.5), w.values, color=COL[g], lw=2, marker="o", ms=3.5, zorder=3)
            ymax = max(ymax, w.max())
        ax.set_ylim(0, ymax * 1.28)
        for p0, lab in ((PAY1, "1st payment"), (PAY2, "2nd payment")):
            ax.text(p0 + pd.Timedelta(days=3.5), ymax * 1.24, lab, ha="center", va="top", fontsize=8.8, color=INK2)
        ax.set_xlim(weeks[0], weeks[-1] + pd.Timedelta(days=7))
        ax.xaxis.set_major_locator(mdates.MonthLocator()); ax.xaxis.set_major_formatter(mdates.DateFormatter("%b"))
        ax.yaxis.set_major_formatter(FuncFormatter(inr))
        ax.set_title(ACT[key], loc="left", fontsize=13, fontweight="bold", color=INK)
        if j == 0:
            ax.set_ylabel("Spent in the week\n(₹ per acre, average household)", color=INK2)

        # ---- bottom: per-household cost per acre before / from the 1st payment, mean with 95% CI
        ax2 = axes[1, j]
        ax2.set_facecolor(SURF)
        tot = x.assign(per=np.where(x["act_n"] < PAY1, "before", "after")).pivot_table(
            index="_index", columns="per", values="cpa", aggfunc="sum", fill_value=0)
        tot = tot.reindex(hh.index).fillna(0)
        tot["season"] = tot.get("before", 0) + tot.get("after", 0)
        groups = [("before", "Before 1st payment\n(before 1 Aug)"), ("after", "From 1st payment\n(1 Aug onwards)"), ("season", "Whole season")]
        xs = np.arange(len(groups)); bw = 0.36
        top = 0
        for gi, g in enumerate((1, 0)):
            means, los, his = [], [], []
            for col, _ in groups:
                v = tot.loc[hh["payment"] == g, col]
                m, se = v.mean(), v.std(ddof=1) / np.sqrt(len(v))
                tcrit = stats.t.ppf(0.975, len(v) - 1)
                means.append(m); los.append(m - tcrit * se); his.append(m + tcrit * se)
            pos = xs + (gi - 0.5) * bw
            ax2.bar(pos, means, width=bw - 0.04, color=COL[g], zorder=3)
            ax2.errorbar(pos, means, yerr=[np.array(means) - los, np.array(his) - means], fmt="none", ecolor=INK, elinewidth=1.1, capsize=3, zorder=4)
            for p, m, h in zip(pos, means, his):
                ax2.annotate(inr(m), (p, h), xytext=(0, 4), textcoords="offset points", ha="center", va="bottom", fontsize=8.6, color=INK2)
            top = max(top, max(his))
        for k, (col, _) in enumerate(groups):
            a, b = tot.loc[hh["payment"] == 1, col], tot.loc[hh["payment"] == 0, col]
            p = stats.ttest_ind(a, b, equal_var=False).pvalue
            stats_out[(key, col)] = (a.mean(), b.mean(), p)
            ax2.text(xs[k], top * 1.2, f"p = {p:.2f}", ha="center", va="bottom", fontsize=9, color=INK)
        ax2.set_ylim(0, top * 1.35)
        ax2.set_xticks(xs); ax2.set_xticklabels([g[1] for g in groups])
        ax2.yaxis.set_major_formatter(FuncFormatter(inr))
        if j == 0:
            ax2.set_ylabel("Cost per acre (₹), mean and 95% CI", color=INK2)
        for a_ in (ax, ax2):
            a_.grid(axis="y", color=GRID, lw=0.8, zorder=0)
            for sp in ("top", "right", "left"):
                a_.spines[sp].set_visible(False)
            a_.spines["bottom"].set_color(AXIS)
            a_.tick_params(colors=MUTED, length=0, labelsize=9.8)

    fig.text(0.07, 0.955, "Pesticide and 3rd fertiliser dose: spending per acre by payment status, Virippu 2026", fontsize=16, fontweight="bold", color=INK)
    fig.text(0.07, 0.928, "Top: weekly spending per acre, averaged over all households in each group (households that did not buy count as zero).\n"
             "Bottom: each household's spending per acre before and from the 1st RCT payment.", fontsize=10.3, color=INK2, va="top")
    fig.legend(handles=[Line2D([0], [0], color=COL[1], lw=2.5, label=GLAB[1]), Line2D([0], [0], color=COL[0], lw=2.5, label=GLAB[0]),
                        Patch(fc=INK, alpha=0.12, label="Week of an RCT payment")],
               loc="upper left", bbox_to_anchor=(0.065, 0.885), ncol=3, frameon=False, fontsize=9.8)
    note = ("Source: Survey Round 2, Section 3 (Cash transfer RCT); households with rice area > 0. Material costs only (application labour not included); "
            "cost = quantity × price per unit, ÷ rice_acres_2026. Weeks start Monday; activity dates. Items flagged as entry errors and undated items excluded. "
            "p-values: Welch two-sample t-test, treatment vs control. Payment status from RCT_payment_round1.")
    fig.text(0.07, 0.025, "\n".join(textwrap.wrap(note, 200)), fontsize=8.3, color=MUTED, va="bottom")
    fig.savefig(OUT, dpi=200, pil_kwargs={"quality": 95})
    # change in spending from before to after the 1st payment, treatment vs control (difference-in-differences of means)
    for key in ("e_pest", "e_fert3"):
        x = it[it["key"] == key]
        t = x.assign(per=np.where(x["act_n"] < PAY1, "before", "after")).pivot_table(index="_index", columns="per", values="cpa", aggfunc="sum", fill_value=0).reindex(hh.index).fillna(0)
        ch = t.get("after", 0) - t.get("before", 0)
        a, b = ch[hh["payment"] == 1], ch[hh["payment"] == 0]
        print(key, f"DiD {a.mean() - b.mean():,.0f}  p {stats.ttest_ind(a, b, equal_var=False).pvalue:.3f}")
    for k, v in stats_out.items():
        print(k, f"T {v[0]:,.0f}  C {v[1]:,.0f}  p {v[2]:.3f}")


# =====================================================================================
# STEP 6 — pesticide and 3rd dose, Round 1 as pre-payment check  ->  Pesticide_3rd_fertiliser_round1_vs_round2.jpg
# =====================================================================================

def pesticide_fert3_rounds(ITEMS1, DATA1, ITEMS2, DATA2, OUT):
    """Pesticide and 3rd fertiliser dose by payment status: Round 1 (Virippu 2025, before any payment) as a placebo check
    against Round 2 (Virippu 2026, payments in the 1st weeks of August and September). (from pest_fert3_both.py)"""
    ACT = {"e_pest": "Pesticide", "e_fert3": "3rd fertiliser dose"}
    INK, INK2, MUTED, GRID, AXIS, SURF = "#141613", "#4d524a", "#868a80", "#e2e4dc", "#c3c6ba", "#fcfcfb"
    COL = {1: "#2a78d6", 0: "#eb6834"}
    inr = lambda v, _=None: f"₹{v:,.0f}" if v >= 0 else f"−₹{-v:,.0f}"

    def load(items, data, acres_col, year):
        hh = pd.read_excel(data, sheet_name="data", usecols=["_index", "payment", acres_col])
        hh = hh[hh[acres_col] > 0].set_index("_index")
        it = pd.read_csv(items, parse_dates=["act_n", "act_wk"])
        it = it[it["_index"].isin(hh.index) & it["key"].isin(list(ACT))].dropna(subset=["act_n"]).copy()
        it["cpa"] = it["cost"] / it["_index"].map(hh[acres_col])
        cut = pd.Timestamp(f"{year}-08-01")
        it["per"] = np.where(it["act_n"] < cut, "before", "after")
        return {"hh": hh, "it": it, "year": year, "cut": cut, "N": hh["payment"].value_counts()}

    R = {1: load(ITEMS1, DATA1, "rice_acres_2025", 2025), 2: load(ITEMS2, DATA2, "rice_acres_2026", 2026)}

    def household_split(r, key):
        x = r["it"][r["it"]["key"] == key]
        t = x.pivot_table(index="_index", columns="per", values="cpa", aggfunc="sum", fill_value=0).reindex(r["hh"].index).fillna(0)
        for c in ("before", "after"):
            if c not in t:
                t[c] = 0.0
        t["change"] = t["after"] - t["before"]
        t["payment"] = r["hh"]["payment"]
        return t

    def mean_ci(v):
        m, se = v.mean(), v.std(ddof=1) / np.sqrt(len(v))
        h = stats.t.ppf(0.975, len(v) - 1) * se
        return m, m - h, m + h

    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10.5})
    fig, axes = plt.subplots(3, 2, figsize=(13, 12.6), dpi=200, gridspec_kw={"height_ratios": [1, 1, 1.15], "hspace": 0.5, "wspace": 0.16})
    fig.patch.set_facecolor("white")
    fig.subplots_adjust(left=0.08, right=0.98, top=0.835, bottom=0.095)

    def style(ax):
        ax.set_facecolor(SURF)
        ax.grid(axis="y", color=GRID, lw=0.8, zorder=0)
        for sp in ("top", "right", "left"):
            ax.spines[sp].set_visible(False)
        ax.spines["bottom"].set_color(AXIS)
        ax.tick_params(colors=MUTED, length=0, labelsize=9.6)

    results = {}
    for j, key in enumerate(ACT):
        # ---- rows 1-2: weekly spending per acre, same y scale in both rounds; Round 2 dates shifted back a year to align the calendar
        ymax = 0
        series = {}
        for rnd in (1, 2):
            r = R[rnd]
            x = r["it"][r["it"]["key"] == key]
            weeks = pd.date_range(f"{r['year']}-05-26", f"{r['year']}-10-06", freq="W-MON")
            for g in (1, 0):
                w = x[x["payment"] == g].groupby("act_wk")["cpa"].sum().reindex(weeks, fill_value=0) / r["N"][g]
                series[(rnd, g)] = w
                ymax = max(ymax, w.max())
        for rnd, row in ((1, 0), (2, 1)):
            ax = axes[row, j]
            style(ax)
            yr = R[rnd]["year"]
            for p0, lab in ((pd.Timestamp(f"{yr}-08-01"), "1st payment"), (pd.Timestamp(f"{yr}-09-01"), "2nd payment")):
                if rnd == 2:
                    ax.axvspan(p0, p0 + pd.Timedelta(days=7), color=INK, alpha=0.08, lw=0, zorder=1)
                    ax.text(p0 + pd.Timedelta(days=3.5), ymax * 1.22, lab, ha="center", va="top", fontsize=8.6, color=INK2)
                else:
                    ax.axvspan(p0, p0 + pd.Timedelta(days=7), fc="none", ec=MUTED, ls=(0, (3, 2)), lw=0.9, zorder=1)
                    ax.text(p0 + pd.Timedelta(days=3.5), ymax * 1.22, "same week,\nno payment", ha="center", va="top", fontsize=8, color=MUTED)
            for g in (1, 0):
                w = series[(rnd, g)]
                ax.plot(w.index + pd.Timedelta(days=3.5), w.values, color=COL[g], lw=2, marker="o", ms=3.2, zorder=3)
            ax.set_ylim(0, ymax * 1.26)
            ax.set_xlim(pd.Timestamp(f"{yr}-05-26"), pd.Timestamp(f"{yr}-10-13"))
            ax.xaxis.set_major_locator(mdates.MonthLocator()); ax.xaxis.set_major_formatter(mdates.DateFormatter("%b"))
            ax.yaxis.set_major_formatter(FuncFormatter(inr))
            head = (f"{ACT[key]} · Round 1, {yr} (no payment yet)" if rnd == 1 else f"{ACT[key]} · Round 2, {yr} (payments made)")
            ax.set_title(head, loc="left", fontsize=11.5, fontweight="bold", color=INK)
            if j == 0:
                ax.set_ylabel("Spent in the week\n(₹ per acre, average household)", color=INK2, fontsize=9.8)

        # ---- row 3: change from before to after 1 August (after − before), treatment vs control, both rounds
        ax = axes[2, j]
        style(ax)
        xs = np.array([0, 1])
        bw = 0.34
        top, bot = 0, 0
        for gi, g in enumerate((1, 0)):
            ms, lo, hi = [], [], []
            for rnd in (1, 2):
                t = household_split(R[rnd], key)
                m, l, h = mean_ci(t.loc[t["payment"] == g, "change"])
                ms.append(m); lo.append(l); hi.append(h)
            pos = xs + (gi - 0.5) * bw
            ax.bar(pos, ms, width=bw - 0.04, color=COL[g], zorder=3)
            ax.errorbar(pos, ms, yerr=[np.array(ms) - lo, np.array(hi) - ms], fmt="none", ecolor=INK, elinewidth=1.1, capsize=3, zorder=4)
            for p, m, h, l in zip(pos, ms, hi, lo):
                ax.annotate(inr(m), (p, h if m >= 0 else l), xytext=(0, 4 if m >= 0 else -4), textcoords="offset points",
                            ha="center", va="bottom" if m >= 0 else "top", fontsize=8.6, color=INK2)
            top, bot = max(top, max(hi)), min(bot, min(lo))
        ax.axhline(0, color=AXIS, lw=1, zorder=2)
        for k, rnd in enumerate((1, 2)):
            t = household_split(R[rnd], key)
            a, b = t.loc[t["payment"] == 1], t.loc[t["payment"] == 0]
            p_ch = stats.ttest_ind(a["change"], b["change"], equal_var=False).pvalue
            did = a["change"].mean() - b["change"].mean()
            results[(key, rnd)] = {"T_before": a["before"].mean(), "C_before": b["before"].mean(), "T_after": a["after"].mean(),
                                   "C_after": b["after"].mean(), "did": did, "p": p_ch,
                                   "p_season": stats.ttest_ind(a["before"] + a["after"], b["before"] + b["after"], equal_var=False).pvalue}
            ax.text(xs[k], top + (top - bot) * 0.16, f"T − C = {inr(did)}\np = {p_ch:.2f}",
                    ha="center", va="bottom", fontsize=9, color=INK)
        span = top - bot
        ax.set_ylim(bot - span * 0.12, top + span * 0.45)
        ax.set_xticks(xs); ax.set_xticklabels(["Round 1, Virippu 2025\n(no payment: placebo)", "Round 2, Virippu 2026\n(payments made)"])
        ax.yaxis.set_major_formatter(FuncFormatter(inr))
        ax.set_title(f"{ACT[key]} · spending from 1 Aug minus before 1 Aug", loc="left", fontsize=11.5, fontweight="bold", color=INK)
        if j == 0:
            ax.set_ylabel("Change in ₹ per acre,\nmean and 95% CI", color=INK2, fontsize=9.8)

    fig.text(0.08, 0.962, "Pesticide and 3rd fertiliser dose by payment status: Round 1 as a pre-payment check", fontsize=16, fontweight="bold", color=INK)
    fig.text(0.08, 0.937, "Rows 1–2: weekly spending per acre, averaged over all households in each group (non-buyers count as zero), on the same scale in both rounds.\n"
             "Row 3: each household's spending from 1 August minus spending before 1 August. If the payments moved purchases later, the treatment–control gap\n"
             "should appear in Round 2 but not in Round 1, when nobody had been paid.", fontsize=10, color=INK2, va="top")
    fig.legend(handles=[Line2D([0], [0], color=COL[1], lw=2.5, label="Treatment (payment = 1)"), Line2D([0], [0], color=COL[0], lw=2.5, label="Control (payment = 0)"),
                        Patch(fc=INK, alpha=0.12, label="Week of an RCT payment (2026)"),
                        Patch(fc="none", ec=MUTED, ls=(0, (3, 2)), label="Same calendar week in 2025 (no payment)")],
               loc="upper left", bbox_to_anchor=(0.075, 0.888), ncol=4, frameon=False, fontsize=9.6)
    n1, n2 = R[1]["N"], R[2]["N"]
    note = (f"Households with rice area > 0: Round 1 treatment {n1[1]}, control {n1[0]}; Round 2 treatment {n2[1]}, control {n2[0]}. Material costs only (application labour not included); "
            "cost = quantity × price per unit ÷ rice area of that season. Round 1 dates placed on the 2025 calendar (most were recorded as 2026 by mistake). Items flagged as entry "
            "errors and undated items excluded. p-values: Welch two-sample t-test on the change, treatment vs control. Payment status from RCT_payment_round1 in both rounds.")
    fig.text(0.08, 0.02, "\n".join(textwrap.wrap(note, 190)), fontsize=8.2, color=MUTED, va="bottom")
    fig.savefig(OUT, dpi=200, pil_kwargs={"quality": 95})
    for k, v in results.items():
        print(k, {a: round(b, 3) if isinstance(b, float) else b for a, b in v.items()})


# =====================================================================================
# STEP 7a — use of the Rs 2,000 (pmkisan_survey_use), pies by payment group  ->  Round2_pmkisan_survey_use_pies.jpg
# =====================================================================================

def pmkisan_pies_round2(DATA, OUT, CSV):
    """Categorise pmkisan_survey_use (free text) and draw pie charts by payment group (Round 2). (from pmkisan_use.py)"""
    d = pd.read_excel(DATA, sheet_name="data", usecols=["farmers_id", "payment", "pmkisan_survey", "pmkisan_survey_use"])
    d = d[d["pmkisan_survey"] == 1].copy()          # question asked only of households that received the Rs 2,000

    RULES = [  # (pattern, theme), first match wins
        (r"dont know|don't know|not decided$|no idea", "dontknow"),
        (r"harvest|drying", "harvest"),
        (r"agri|fertili[sz]|pesticide|insecticide|weed|grass|plough|farm|seed|labour|wage|crop|tractor", "farm"),
        (r"hospital|medic|health|treatment", "health"),
        (r"household|family|own exp|personal|immediate|consumption|dress|onam|food|grocer|shop|current bill|electric|lpg|gas", "household"),
        (r"petrol|education|school|loan|temple|insurance|vehicle|sav(ed|ing)", "other"),
    ]
    def themes(text):
        t = str(text).lower()
        parts = [p.strip() for p in re.split(r",|/| and | or |&|\+", t) if p.strip()]
        found = []
        for p in parts:
            for pat, th in RULES:
                if re.search(pat, p):
                    found.append(th)
                    break
            else:
                found.append("unmatched")
        return found

    def category(found):
        s = set(found)
        farm = s & {"harvest", "farm"}
        nonfarm = s - {"harvest", "farm"}
        if farm and (nonfarm & {"household", "health", "other"}):
            return "Farming and household"
        if farm:
            return "Harvesting" if s == {"harvest"} else "Other farming"
        if "household" in s:
            return "Household and family"
        if "health" in s:
            return "Health care"
        return "Other or don't know"

    d["themes"] = d["pmkisan_survey_use"].map(themes)
    d["category"] = d["themes"].map(category)
    unmatched = d[d["themes"].map(lambda f: "unmatched" in f)]
    d[["farmers_id", "payment", "pmkisan_survey_use", "category"]].to_csv(CSV, index=False)

    ORDER = ["Household and family", "Harvesting", "Other farming", "Farming and household", "Health care", "Other or don't know"]
    COL = ["#2a78d6", "#1baf7a", "#008300", "#eda100", "#e87ba4", "#868a80"]   # categorical slots; grey for the residual
    tab = pd.crosstab(d["category"], d["payment"]).reindex(ORDER).fillna(0).astype(int)
    chi2, p, dof, exp = stats.chi2_contingency(tab.loc[tab.sum(axis=1) > 0])
    farm_any = d["category"].isin(["Harvesting", "Other farming", "Farming and household"])
    ft = pd.crosstab(farm_any, d["payment"])
    p_farm = stats.fisher_exact(ft.values)[1] if ft.shape == (2, 2) else float("nan")

    INK, INK2, MUTED = "#141613", "#4d524a", "#868a80"
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10.5})
    fig, axes = plt.subplots(1, 2, figsize=(13, 7.2), dpi=200)
    fig.patch.set_facecolor("white")
    fig.subplots_adjust(left=0.02, right=0.98, top=0.80, bottom=0.17, wspace=0.05)
    for ax, g, name in ((axes[0], 1, "Treatment (payment = 1)"), (axes[1], 0, "Control (payment = 0)")):
        v = tab[g]
        n = int(v.sum())
        keep = v > 0
        wedges, _ = ax.pie(v[keep], colors=[c for c, k in zip(COL, keep) if k], startangle=90, counterclock=False, radius=0.9,
                           wedgeprops=dict(edgecolor="white", linewidth=2))
        ax.set_xlim(-2.15, 2.15); ax.set_ylim(-1.35, 1.55)
        labels = []
        for w, (lab, cnt) in zip(wedges, v[keep].items()):
            ang = np.deg2rad((w.theta2 + w.theta1) / 2)
            x, y = np.cos(ang), np.sin(ang)
            labels.append({"x": x, "y": y, "ty": 1.12 * y, "txt": f"{lab}\n{cnt / n:.0%} ({cnt})"})
        for side in (1, -1):                                   # push labels apart vertically on each side
            grp = sorted([l for l in labels if np.sign(l["x"] if abs(l["x"]) > 1e-9 else 1) == side], key=lambda l: l["ty"])
            for k in range(1, len(grp)):
                if grp[k]["ty"] - grp[k - 1]["ty"] < 0.27:
                    grp[k]["ty"] = grp[k - 1]["ty"] + 0.27
        for l in labels:
            tx = 1.2 * l["x"] if abs(l["x"]) > 0.25 else (0.42 if l["x"] >= 0 else -0.42) + 0.75 * l["x"]
            ax.annotate(l["txt"], xy=(0.89 * l["x"], 0.89 * l["y"]), xytext=(tx, l["ty"]),
                        ha="left" if l["x"] >= 0 else "right", va="center", fontsize=9.6, color=INK,
                        arrowprops=dict(arrowstyle="-", color=MUTED, lw=0.8, shrinkA=0, shrinkB=0))
        ax.set_title(f"{name}\nn = {n} households", fontsize=12.5, fontweight="bold", color=INK, pad=6)
        ax.set_aspect("equal")
    fig.text(0.03, 0.95, "What households did with the Rs 2,000 transfer, by payment group (Survey Round 2)",
             fontsize=15.5, fontweight="bold", color=INK)
    farm_t = farm_any[d["payment"] == 1].mean(); farm_c = farm_any[d["payment"] == 0].mean()
    fig.text(0.03, 0.91, f"Any farm use: treatment {farm_t:.0%}, control {farm_c:.0%} (Fisher's exact p = {p_farm:.2f}). "
             f"Distribution across all categories: chi-square p = {p:.2f}.", fontsize=10.5, color=INK2)
    note = ("Question 4 (asked of households that received the Rs 2,000): 'What did you do with the cash received from the PM-KISAN survey team? "
            "(If not yet used, what do they plan to use it for?)'. Free-text answers grouped by Claude: Household and family = household, family, personal or immediate "
            "expenses, Onam purchases, clothes, shop, electricity and LPG bills; Harvesting = harvesting only; Other farming = agriculture, fertiliser, pesticide, weeding, grass cutting, wages, next season's crop (with or without "
            "harvesting); Farming and household = farm use together with household, health or other use; Health care = medicine, hospital; Other = petrol, education, "
            "loan repayment, temple, vehicle insurance, savings, don't know. Answers naming household and health together are counted as Household and family. "
            "Treatment = Payment 'Yes' in RCT_payment_round1.")
    fig.text(0.03, 0.025, "\n".join(textwrap.wrap(note, 195)), fontsize=8.2, color=MUTED, va="bottom")
    fig.savefig(OUT, dpi=200, pil_kwargs={"quality": 95})
    print(tab); print("chi2 p", round(p, 3), "| farm any T/C", round(farm_t, 3), round(farm_c, 3), "p", round(p_farm, 3))
    print("unmatched answers:", unmatched["pmkisan_survey_use"].tolist())


# =====================================================================================
# STEP 7b — use of PM-KISAN money (pmkisan_use), Round 1 combined pie  ->  Round1_pmkisan_use_pie.jpg
# =====================================================================================

def pmkisan_pie_round1(DATA, OUT, CSV):
    """Round 1: categorise pmkisan_use (free text, use of the PM-KISAN government transfer) and draw one combined pie chart. (from pmkisan_use_r1.py)"""
    d = pd.read_excel(DATA, sheet_name="data", usecols=["farmers_id", "payment", "pmkisan_received", "pmkisan_use"])
    d = d[(d["pmkisan_received"] == 1) & d["pmkisan_use"].notna()].copy()

    # answers saying the money was not received (no use to report) are left out of the pie
    NOT_RECEIVED = re.compile(r"hasn't received|has not received|not received|not taking|received once", re.I)
    d["not_received"] = d["pmkisan_use"].astype(str).str.contains(NOT_RECEIVED)

    RULES = [  # (pattern, theme), first match wins; same groups as the Round 2 chart
        (r"^999$|dont know|don't know|cannot attribute$|^na$|^nag$|no idea", "dontknow"),
        (r"loan|kcc", "other"),
        (r"harvest|drying", "harvest"),
        (r"agri|fertili[sz]|pesticide|insecticide|weed|grass|plough|farm|seed|labour|wage|crop|tractor", "farm"),
        (r"hospital|medic|health|treatment|surgery", "health"),
        (r"household|family|own exp|personal|immediate|consumption|dress|onam|food|grocer|shop|current bill|electric|lpg|gas|expense", "household"),
        (r"petrol|education|school|temple|insurance|vehicle|sav(ed|ing)", "other"),
    ]
    def themes(text):
        t = str(text).lower().strip()
        parts = [p.strip() for p in re.split(r",|/| and | or |&|\+", t) if p.strip()] or [t]
        out = []
        for p in parts:
            for pat, th in RULES:
                if re.search(pat, p):
                    out.append(th)
                    if th in ("farm", "harvest") and re.search(r"household|family", p):   # e.g. "Agriculture purpose household"
                        out.append("household")
                    break
            else:
                out.append("unmatched")
        return out

    def category(found):
        s = set(found)
        farm = s & {"harvest", "farm"}
        if farm and (s & {"household", "health", "other"}):
            return "Farming and household"
        if farm:
            return "Harvesting" if s == {"harvest"} else "Other farming"
        if "household" in s:
            return "Household and family"
        if "health" in s:
            return "Health care"
        return "Other or don't know"

    d["themes"] = d["pmkisan_use"].map(themes)
    d["category"] = np.where(d["not_received"], "(not received: excluded)", d["themes"].map(category))
    d[["farmers_id", "payment", "pmkisan_use", "category"]].to_csv(CSV, index=False)
    unmatched = d[d["themes"].map(lambda f: "unmatched" in f) & ~d["not_received"]]
    use = d[~d["not_received"]]

    ORDER = ["Household and family", "Harvesting", "Other farming", "Farming and household", "Health care", "Other or don't know"]
    COL = ["#2a78d6", "#1baf7a", "#008300", "#eda100", "#e87ba4", "#868a80"]
    v = use["category"].value_counts().reindex(ORDER).fillna(0).astype(int)
    n = int(v.sum())

    INK, INK2, MUTED = "#141613", "#4d524a", "#868a80"
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10.5})
    fig, ax = plt.subplots(figsize=(10.5, 7.4), dpi=200)
    fig.patch.set_facecolor("white")
    fig.subplots_adjust(left=0.02, right=0.98, top=0.86, bottom=0.16)
    keep = v > 0
    wedges, _ = ax.pie(v[keep], colors=[c for c, k in zip(COL, keep) if k], startangle=90, counterclock=False, radius=0.9,
                       wedgeprops=dict(edgecolor="white", linewidth=2))
    ax.set_xlim(-2.2, 2.2); ax.set_ylim(-1.3, 1.3); ax.set_aspect("equal")
    labels = []
    for w, (lab, cnt) in zip(wedges, v[keep].items()):
        ang = np.deg2rad((w.theta2 + w.theta1) / 2)
        x, y = np.cos(ang), np.sin(ang)
        shown = "Farming (agriculture, fertiliser,\npesticide; no harvest-only answers)" if lab == "Other farming" else lab
        labels.append({"x": x, "y": y, "ty": 1.12 * y, "txt": f"{shown}\n{cnt / n:.0%} ({cnt})"})
    for side in (1, -1):
        grp = sorted([l for l in labels if np.sign(l["x"] if abs(l["x"]) > 1e-9 else 1) == side], key=lambda l: l["ty"])
        for k in range(1, len(grp)):
            if grp[k]["ty"] - grp[k - 1]["ty"] < 0.27:
                grp[k]["ty"] = grp[k - 1]["ty"] + 0.27
    for l in labels:
        tx = 1.2 * l["x"] if abs(l["x"]) > 0.25 else (0.42 if l["x"] >= 0 else -0.42) + 0.75 * l["x"]
        ax.annotate(l["txt"], xy=(0.89 * l["x"], 0.89 * l["y"]), xytext=(tx, l["ty"]), ha="left" if l["x"] >= 0 else "right",
                    va="center", fontsize=10, color=INK, arrowprops=dict(arrowstyle="-", color=MUTED, lw=0.8, shrinkA=0, shrinkB=0))
    farm_share = use["category"].isin(["Harvesting", "Other farming", "Farming and household"]).mean()
    fig.text(0.04, 0.94, "What households did with their PM-KISAN money (Survey Round 1, all households)", fontsize=15, fontweight="bold", color=INK)
    fig.text(0.04, 0.90, f"{n} households that received PM-KISAN and reported a use. Any farm use: {farm_share:.0%}.", fontsize=10.5, color=INK2)
    note = ("Round 1, Section 6, question 4 (asked if the household ever received PM-KISAN): 'What did you do with the cash received from the PM-KISAN scheme?'. "
            "This is the government PM-KISAN transfer, not the Rs 2,000 from the survey team asked about in Round 2. Free-text answers grouped by Claude into the "
            "same six groups as the Round 2 chart; loan repayments (incl. agricultural / KCC loans) count as Other; '999' counts as don't know. "
            f"{int(d['not_received'].sum())} answers saying the money had not been received are left out.")
    fig.text(0.04, 0.03, "\n".join(textwrap.wrap(note, 165)), fontsize=8.3, color=MUTED, va="bottom")
    fig.savefig(OUT, dpi=200, pil_kwargs={"quality": 95})
    print(v.to_string(), "\nn", n, "| farm share", round(farm_share, 3), "| excluded not received", int(d["not_received"].sum()))
    print("unmatched:", unmatched["pmkisan_use"].tolist())
    print(d.groupby(["category"])["pmkisan_use"].apply(lambda s: sorted(set(s.astype(str).str.strip())))
          .to_string())


# =====================================================================================
# STEP 7c — Excel workbook for the pies (Round 2 sheets)  ->  PMKISAN_use_pies_round1_round2.xlsx
# =====================================================================================

def pmkisan_excel(CSV, OUT, JPG):
    """Excel output for the pmkisan_survey_use pie charts: summary table with formulas, native pie charts, coding sheet, notes. (from pmkisan_excel.py)"""
    c = pd.read_csv(CSV)
    ORDER = ["Household and family", "Harvesting", "Other farming", "Farming and household", "Health care", "Other or don't know"]
    HEX = ["2A78D6", "1BAF7A", "008300", "EDA100", "E87BA4", "868A80"]

    # tests (computed here, written as values and documented)
    tab = pd.crosstab(c["category"], c["payment"]).reindex(ORDER).fillna(0).astype(int)
    chi2, p_chi, dof, _ = stats.chi2_contingency(tab)
    farm = c["category"].isin(["Harvesting", "Other farming", "Farming and household"])
    p_farm = stats.fisher_exact(pd.crosstab(farm, c["payment"]).values)[1]

    wb = Workbook()
    F = lambda **k: Font(**{"name": "Arial", "size": 10, **k})
    head_fill = PatternFill("solid", start_color="DDE4EE")
    thin = Side(style="thin", color="9A9A9A")

    # ---------------- coding sheet (the data the summary counts from) ----------------
    cs = wb.active
    cs.title = "coding"
    cols = ["farmers_id", "payment", "group", "pmkisan_survey_use (as recorded)", "use_category"]
    cs.append(cols)
    for r in c.itertuples():
        cs.append([r.farmers_id, int(r.payment), "Treatment" if r.payment == 1 else "Control",
                   str(r.pmkisan_survey_use).replace("\n", " ").strip(), r.category])
    n_rows = cs.max_row
    for cell in cs[1]:
        cell.font = F(bold=True); cell.fill = head_fill
    for row in cs.iter_rows(min_row=2):
        for cell in row:
            cell.font = F()
    for col, w in zip("ABCDE", [11, 9, 11, 62, 24]):
        cs.column_dimensions[col].width = w
    cs.freeze_panes = "A2"
    cs.auto_filter.ref = f"A1:E{n_rows}"

    # ---------------- summary sheet ----------------
    ws = wb.create_sheet("summary", 0)
    ws["A1"] = "Use of the Rs 2,000 transfer (pmkisan_survey_use), by payment group, Survey Round 2"
    ws["A1"].font = F(bold=True, size=13)
    ws["A2"] = ("Households that received the Rs 2,000 (pmkisan_survey = 1). Counts are formulas on the 'coding' sheet, "
                "so changing a category there updates this table and the charts.")
    ws["A2"].font = F(italic=True, color="4D524A")
    hdr = ["Use category", "Treatment (n)", "Treatment (%)", "Control (n)", "Control (%)", "All (n)", "All (%)", "Difference T − C (pct points)"]
    ws.append([])
    ws.append(hdr)                                              # row 4
    for cell in ws[4]:
        cell.font = F(bold=True); cell.fill = head_fill
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws["A4"].alignment = Alignment(vertical="center")
    first = 5
    for i, cat in enumerate(ORDER):
        r = first + i
        ws.cell(r, 1, cat)
        ws.cell(r, 2, f'=COUNTIFS(coding!$E$2:$E${n_rows},$A{r},coding!$B$2:$B${n_rows},1)')
        ws.cell(r, 4, f'=COUNTIFS(coding!$E$2:$E${n_rows},$A{r},coding!$B$2:$B${n_rows},0)')
        ws.cell(r, 6, f"=B{r}+D{r}")
        tot = first + len(ORDER)
        ws.cell(r, 3, f"=IF(B${tot}=0,0,B{r}/B${tot})")
        ws.cell(r, 5, f"=IF(D${tot}=0,0,D{r}/D${tot})")
        ws.cell(r, 7, f"=IF(F${tot}=0,0,F{r}/F${tot})")
        ws.cell(r, 8, f"=(C{r}-E{r})*100")
    last = first + len(ORDER) - 1
    tot = last + 1
    ws.cell(tot, 1, "Total")
    for col in "BDF":
        ws[f"{col}{tot}"] = f"=SUM({col}{first}:{col}{last})"
    for col in "CEG":
        ws[f"{col}{tot}"] = f"=SUM({col}{first}:{col}{last})"
    farm_r = tot + 1
    ws.cell(farm_r, 1, "Any farm use (Harvesting + Other farming + Farming and household)")
    ws[f"B{farm_r}"] = f"=B{first + 1}+B{first + 2}+B{first + 3}"
    ws[f"D{farm_r}"] = f"=D{first + 1}+D{first + 2}+D{first + 3}"
    ws[f"F{farm_r}"] = f"=B{farm_r}+D{farm_r}"
    ws[f"C{farm_r}"] = f"=IF(B{tot}=0,0,B{farm_r}/B{tot})"
    ws[f"E{farm_r}"] = f"=IF(D{tot}=0,0,D{farm_r}/D{tot})"
    ws[f"G{farm_r}"] = f"=IF(F{tot}=0,0,F{farm_r}/F{tot})"
    ws[f"H{farm_r}"] = f"=(C{farm_r}-E{farm_r})*100"
    for r in range(first, farm_r + 1):
        for col in range(1, 9):
            cell = ws.cell(r, col)
            cell.font = F(bold=(r >= tot))
            if col in (3, 5, 7):
                cell.number_format = "0.0%"
            elif col == 8:
                cell.number_format = "+0.0;-0.0;0.0"
        ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="center")
    for col in range(1, 9):
        ws.cell(tot, col).border = Border(top=thin)
    ws.row_dimensions[farm_r].height = 28

    t0 = farm_r + 2
    ws.cell(t0, 1, "Tests (computed in Python from the coding sheet as delivered; they do not update if categories are changed)").font = F(bold=True)
    ws.cell(t0 + 1, 1, "Any farm use, treatment vs control: Fisher's exact test p-value")
    ws.cell(t0 + 1, 2, round(p_farm, 4))
    ws.cell(t0 + 2, 1, f"Distribution across the 6 categories: chi-square ({dof} df) p-value")
    ws.cell(t0 + 2, 2, round(p_chi, 4))
    for r in (t0 + 1, t0 + 2):
        ws.cell(r, 1).font = F(); ws.cell(r, 2).font = F(color="0000FF"); ws.cell(r, 2).number_format = "0.000"

    for col, w in zip("ABCDEFGH", [44, 13, 13, 12, 12, 10, 10, 16]):
        ws.column_dimensions[col].width = w
    ws.row_dimensions[4].height = 32

    # native pie charts
    def pie(title, data_col, anchor):
        ch = PieChart()
        ch.title = title
        labels = Reference(ws, min_col=1, min_row=first, max_row=last)
        data = Reference(ws, min_col=data_col, min_row=first - 1, max_row=last)
        ch.add_data(data, titles_from_data=True)
        ch.set_categories(labels)
        s = ch.series[0]
        for i, h in enumerate(HEX):
            pt = DataPoint(idx=i)
            pt.graphicalProperties.solidFill = h
            pt.graphicalProperties.line.solidFill = "FFFFFF"
            s.dPt.append(pt)
        ch.dataLabels = DataLabelList()
        ch.dataLabels.showPercent = True
        ch.dataLabels.showVal = False
        ch.dataLabels.showCatName = False
        ch.dataLabels.showSerName = False
        ch.dataLabels.showLegendKey = False
        ch.legend.position = "r"
        ch.height, ch.width = 9.5, 14
        ws.add_chart(ch, anchor)

    chart_row = t0 + 4
    pie("Treatment (payment = 1)", 2, f"A{chart_row}")
    pie("Control (payment = 0)", 4, f"D{chart_row}")

    # the figure as delivered (JPEG)
    fs = wb.create_sheet("figure")
    fs["A1"] = "Pie charts as delivered (Round2_pmkisan_survey_use_pies.jpg)"
    fs["A1"].font = F(bold=True, size=12)
    img = XLImage(JPG)
    img.width, img.height = int(img.width * 0.42), int(img.height * 0.42)
    fs.add_image(img, "A3")

    # notes
    ns = wb.create_sheet("notes")
    notes = [
        ("Question", "Survey Round 2, question 4 (asked if pmkisan_survey = 1): 'What did you do with the cash received from the PM-KISAN survey team? "
                     "(If they have not used, for what activity they are planning to use it for?)'"),
        ("Sample", f"{len(c)} households that reported receiving the Rs 2,000: {int((c.payment == 1).sum())} treatment, {int((c.payment == 0).sum())} control."),
        ("Groups", "Treatment = Payment 'Yes' in RCT_payment_round1 (payment = 1); control = payment = 0. Note: 141 of the 158 control households also "
                   "report receiving Rs 2,000 from the survey team, so payment does not separate who received money in Round 2."),
        ("Categories", "Household and family = household, family, personal or immediate expenses, Onam purchases, clothes, shop, electricity and LPG bills. "
                       "Harvesting = harvesting only. Other farming = agriculture, fertiliser, pesticide, weeding, grass cutting, wages, ploughing, next season's crop "
                       "(with or without harvesting). Farming and household = a farm use named together with household, health or other use. "
                       "Health care = medicine, hospital. Other or don't know = petrol, education, loan repayment, temple, vehicle insurance, savings, don't know. "
                       "Answers naming household and health together are counted as Household and family."),
        ("Recoding", "To change a category, edit column E of the 'coding' sheet using one of the six category names exactly; the summary table and charts update."),
        ("Tests", "Fisher's exact test (any farm use) and chi-square test of independence (six categories), computed in Python; values in blue are hardcoded."),
    ]
    ns.append(["Item", "Note"])
    for n in notes:
        ns.append(list(n))
    for cell in ns[1]:
        cell.font = F(bold=True); cell.fill = head_fill
    for row in ns.iter_rows(min_row=2):
        for cell in row:
            cell.font = F(); cell.alignment = Alignment(wrap_text=True, vertical="top")
    ns.column_dimensions["A"].width = 14
    ns.column_dimensions["B"].width = 120
    wb.calculation.fullCalcOnLoad = True      # Excel calculates the formulas on opening
    wb.save(OUT)
    print("saved", OUT, "| rows", n_rows - 1)


# =====================================================================================
# STEP 7d — add the Round 1 sheets to the PM-KISAN workbook
# =====================================================================================

def pmkisan_excel_add_round1(XLSX, CSV, JPG):
    """Add Round 1 (pmkisan_use, combined) sheets to the pmkisan Excel workbook. (from pmkisan_excel_r1.py)"""
    c = pd.read_csv(CSV)
    wb = load_workbook(XLSX)
    for nm in ("round1_summary", "round1_coding"):
        if nm in wb.sheetnames:
            del wb[nm]
    F = lambda **k: Font(**{"name": "Arial", "size": 10, **k})
    head_fill = PatternFill("solid", start_color="DDE4EE")
    thin = Side(style="thin", color="9A9A9A")
    ORDER = ["Household and family", "Harvesting", "Other farming", "Farming and household", "Health care", "Other or don't know"]
    HEX = ["2A78D6", "1BAF7A", "008300", "EDA100", "E87BA4", "868A80"]

    cs = wb.create_sheet("round1_coding")
    cs.append(["farmers_id", "payment", "pmkisan_use (as recorded)", "use_category"])
    for r in c.itertuples():
        cs.append([r.farmers_id, int(r.payment), str(r.pmkisan_use).replace("\n", " ").strip(), r.category])
    n_rows = cs.max_row
    for cell in cs[1]:
        cell.font = F(bold=True); cell.fill = head_fill
    for row in cs.iter_rows(min_row=2):
        for cell in row:
            cell.font = F()
    for col, w in zip("ABCD", [11, 9, 70, 26]):
        cs.column_dimensions[col].width = w
    cs.freeze_panes = "A2"; cs.auto_filter.ref = f"A1:D{n_rows}"

    ws = wb.create_sheet("round1_summary", 1)
    ws["A1"] = "Use of PM-KISAN money (pmkisan_use), Survey Round 1, all households combined"
    ws["A1"].font = F(bold=True, size=13)
    ws["A2"] = ("Households that received PM-KISAN and reported a use. This is the government PM-KISAN transfer, not the survey-team Rs 2,000 of Round 2. "
                "Counts are formulas on 'round1_coding'; answers saying the money was not received are excluded.")
    ws["A2"].font = F(italic=True, color="4D524A")
    ws.append([])
    ws.append(["Use category", "Households (n)", "Share (%)"])
    for cell in ws[4]:
        cell.font = F(bold=True); cell.fill = head_fill
    first = 5
    last = first + len(ORDER) - 1
    tot = last + 1
    for i, cat in enumerate(ORDER):
        r = first + i
        ws.cell(r, 1, cat)
        ws.cell(r, 2, f'=COUNTIFS(round1_coding!$D$2:$D${n_rows},$A{r})')
        ws.cell(r, 3, f"=IF($B${tot}=0,0,B{r}/$B${tot})")
    ws.cell(tot, 1, "Total")
    ws[f"B{tot}"] = f"=SUM(B{first}:B{last})"
    ws[f"C{tot}"] = f"=SUM(C{first}:C{last})"
    fr = tot + 1
    ws.cell(fr, 1, "Any farm use (Harvesting + Other farming + Farming and household)")
    ws[f"B{fr}"] = f"=B{first + 1}+B{first + 2}+B{first + 3}"
    ws[f"C{fr}"] = f"=IF($B${tot}=0,0,B{fr}/$B${tot})"
    ex = fr + 1
    ws.cell(ex, 1, "Excluded: money not received")
    ws[f"B{ex}"] = f'=COUNTIFS(round1_coding!$D$2:$D${n_rows},"(not received: excluded)")'
    for r in range(first, ex + 1):
        for col in (1, 2, 3):
            ws.cell(r, col).font = F(bold=(r in (tot, fr)))
        ws.cell(r, 3).number_format = "0.0%"
        ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="center")
    for col in (1, 2, 3):
        ws.cell(tot, col).border = Border(top=thin)
    ws.row_dimensions[fr].height = 28
    for col, w in zip("ABC", [46, 15, 12]):
        ws.column_dimensions[col].width = w
    ws.cell(ex + 2, 1, "Note: 'Other farming' here is all farm use (agriculture, fertiliser, pesticide); no Round 1 answer named harvesting alone.").font = F(italic=True, color="4D524A")

    ch = PieChart()
    ch.title = "Round 1: use of PM-KISAN money (all households)"
    ch.add_data(Reference(ws, min_col=2, min_row=first - 1, max_row=last), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last))
    s = ch.series[0]
    for i, h in enumerate(HEX):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = h
        pt.graphicalProperties.line.solidFill = "FFFFFF"
        s.dPt.append(pt)
    ch.dataLabels = DataLabelList()
    ch.dataLabels.showPercent = True
    for attr in ("showVal", "showCatName", "showSerName", "showLegendKey"):
        setattr(ch.dataLabels, attr, False)
    ch.legend.position = "r"
    ch.height, ch.width = 9.5, 15
    ws.add_chart(ch, "E4")

    fs = wb["figure"]
    fs["A1"] = "Pie charts as delivered: Round 2 by payment group (top), Round 1 combined (below)"
    img = XLImage(JPG)
    img.width, img.height = int(img.width * 0.42), int(img.height * 0.42)
    fs.add_image(img, "A34")

    ns = wb["notes"]
    ns.append(["Round 1", "Round 1, Section 6, question 4 (asked if the household ever received PM-KISAN): 'What did you do with the cash received from the "
                          "PM-KISAN scheme?'. Same six groups as Round 2; loan repayments (incl. agricultural / KCC loans) count as Other; '999' counts as don't know; "
                          "answers saying the money had not been received are excluded. Treatment and control are combined."])
    for cell in ns[ns.max_row]:
        cell.font = F(); cell.alignment = Alignment(wrap_text=True, vertical="top")
    wb.calculation.fullCalcOnLoad = True
    wb.save(XLSX)
    print("ok", n_rows - 1)


# =====================================================================================
# STEP 8 — spending right after each transfer, Round 1 as placebo  ->  Round2_spending_after_transfer.xlsx
# =====================================================================================

def spending_after_transfer(ITEMS1, DATA1, ITEMS2, DATA2, OUT):
    """Spending right after each RCT transfer (Round 2, Virippu 2026) with Round 1 (Virippu 2025) as the
    pre-payment check: tidy Excel output for the paper. (from event_excel.py)"""
    def load(items, data, acres, extra=()):
        hh = pd.read_excel(data, sheet_name="data", usecols=["_index", "farmers_id", "payment", acres, *extra])
        hh["fid"] = hh["farmers_id"].astype(str).str.strip()
        hh = hh[hh[acres] > 0].set_index("_index")
        it = pd.read_csv(items, parse_dates=["act_n", "pay_n"])
        it = it[it["_index"].isin(hh.index)].copy()
        it["cpa"] = it["cost"] / it["_index"].map(hh[acres])
        return hh, it

    h1, i1 = load(ITEMS1, DATA1, "rice_acres_2025")
    h2, i2all = load(ITEMS2, DATA2, "rice_acres_2026", ("pmkisan_survey", "pmkisan_survey_date", "interview_date"))
    i2 = i2all[i2all["anticipated"] == 0]          # harvest costs reported as planned (after the interview) left out of windows

    def win(hh, it, keys, yr, a, b):
        x = it if keys is None else it[it["key"].isin(keys)]
        x = x[(x["act_n"] >= pd.Timestamp(f"{yr}-{a}")) & (x["act_n"] < pd.Timestamp(f"{yr}-{b}"))]
        return x.groupby("_index")["cpa"].sum().reindex(hh.index, fill_value=0)

    def welch(v, g):
        a, c = v[g == 1], v[g == 0]
        p = stats.ttest_ind(a, c, equal_var=False).pvalue if (a.std() > 0 or c.std() > 0) else np.nan
        return a.mean(), c.mean(), a.mean() - c.mean(), p

    def fisher(v, g):
        t = pd.crosstab(v, g)
        return stats.fisher_exact(t.values)[1] if t.shape == (2, 2) else np.nan

    def panel(v1, v2):
        p1 = pd.DataFrame({"fid": h1["fid"], "y1": v1}).groupby("fid")["y1"].sum()
        p2 = pd.DataFrame({"fid": h2["fid"], "y2": v2, "payment": h2["payment"]}).groupby("fid").agg(y2=("y2", "sum"), payment=("payment", "max"))
        pn = p2.join(p1, how="inner")
        d = pn["y2"] - pn["y1"]
        return welch(d, pn["payment"]), len(pn)

    OUTCOMES = [  # label, keys, kind ("amount" per acre or "any" = share of households with a purchase)
        ("Pesticide bought (share of households)", ["e_pest"], "any"),
        ("Pesticide, material (Rs/acre)", ["e_pest"], "amount"),
        ("Pesticide, material + application labour (Rs/acre)", ["e_pest", "e_pestlab"], "amount"),
        ("3rd fertiliser dose applied (share of households)", ["e_fert3"], "any"),
        ("3rd fertiliser dose, material (Rs/acre)", ["e_fert3"], "amount"),
        ("3rd fertiliser dose, material + labour (Rs/acre)", ["e_fert3", "e_fert3lab"], "amount"),
        ("Section E total: pesticide / 3rd dose incl. labour and other items (Rs/acre)", ["e_pest", "e_pestlab", "e_fert3", "e_fert3lab", "e_other"], "amount"),
        ("Harvesting, Section F (Rs/acre)", ["f_mach", "f_tractor", "f_labour", "f_other"], "amount"),
        ("All activities (Rs/acre)", None, "amount"),
    ]
    WINDOWS = [("Before 1 Aug (before any transfer)", "01-01", "08-01"),
               ("1–21 Aug (3 weeks after 1st payment)", "08-01", "08-22"),
               ("1–21 Sep (3 weeks after 2nd payment)", "09-01", "09-22"),
               ("Whole season", "01-01", "12-31")]
    rows = []
    for wl, a, b in WINDOWS:
        rows.append({"panel": wl})
        for lab, keys, kind in OUTCOMES:
            if lab.startswith("Harvesting") and wl.startswith("Before"):
                continue
            v2 = win(h2, i2all if wl == "Whole season" else i2, keys, 2026, a, b)
            v1 = win(h1, i1, keys, 2025, a, b)
            if kind == "any":
                v2, v1 = (v2 > 0).astype(float), (v1 > 0).astype(float)
            r2, r1 = welch(v2, h2["payment"]), welch(v1, h1["payment"])
            (dd, n) = panel(v1, v2)
            rows.append({"Outcome": lab, "kind": kind,
                         "R2 T": r2[0], "R2 C": r2[1], "R2 diff": r2[2], "R2 p": fisher(v2 > 0, h2["payment"]) if kind == "any" else r2[3],
                         "R1 T": r1[0], "R1 C": r1[1], "R1 diff": r1[2], "R1 p": fisher(v1 > 0, h1["payment"]) if kind == "any" else r1[3],
                         "DiD": dd[2], "DiD p": dd[3], "n": n})

    # timing of first purchase
    tim = []
    for (hh, it, yr, rnd) in ((h1, i1, 2025, "Round 1 (2025)"), (h2, i2, 2026, "Round 2 (2026)")):
        for k, lab in (("d_fert2", "2nd fertiliser dose"), ("e_fert3", "3rd fertiliser dose"), ("e_pest", "Pesticide")):
            first = it[it["key"] == k].dropna(subset=["act_n"]).groupby("_index")["act_n"].min()
            g = hh.loc[first.index, "payment"]; doy = first.dt.dayofyear
            md = lambda s: (pd.Timestamp(f"{yr}-01-01") + pd.Timedelta(days=float(s.median()) - 1)).strftime("%d %b")
            tim.append([rnd, lab, int((g == 1).sum()), md(doy[g == 1]), int((g == 0).sum()), md(doy[g == 0]),
                        float(doy[g == 1].median() - doy[g == 0].median()), stats.mannwhitneyu(doy[g == 1], doy[g == 0]).pvalue])

    # season profile, Round 2 (cost per acre of the average household)
    N = len(h2); tot = i2all["cpa"].sum() / N
    prof = []
    for lab, m in (("Before 1 August", i2["act_n"] < "2026-08-01"),
                   ("1–31 August (after the 1st payment)", (i2["act_n"] >= "2026-08-01") & (i2["act_n"] < "2026-09-01")),
                   ("1 September to the interview (after the 2nd payment)", i2["act_n"] >= "2026-09-01"),
                   ("No usable date", i2["act_n"].isna())):
        prof.append([lab, i2.loc[m, "cpa"].sum() / N, i2.loc[m, "cpa"].sum() / N / tot])
    ant = i2all.loc[i2all["anticipated"] == 1, "cpa"].sum() / N
    prof.append(["After the interview (harvest, reported as planned)", ant, ant / tot])
    prof.append(["Season total", tot, 1.0])

    # credit: share of spending paid more than a week after the activity
    x = i2.dropna(subset=["act_n", "pay_n"]); lag = (x["pay_n"] - x["act_n"]).dt.days
    credit = [["Items paid on the day of the activity", (lag == 0).mean()],
              ["Spending paid more than 7 days after the activity", x.loc[lag > 7, "cpa"].sum() / x["cpa"].sum()]]

    # reported receipt date of the Rs 2,000 (Round 2)
    dt = h2["pmkisan_survey_date"]
    bucket = np.select([h2["pmkisan_survey"] != 1, dt.isna(), dt < "2026-09-01", dt < "2026-09-15"],
                       ["Did not receive", "Received, no date", "August", "1–14 September"], "15–30 September")
    rec = pd.crosstab(pd.Series(bucket, index=h2.index, name="b"), h2["payment"]).reindex(
        ["August", "1–14 September", "15–30 September", "Received, no date", "Did not receive"]).fillna(0).astype(int)
    same_day = pd.crosstab(dt.dt.normalize() == h2["interview_date"].dt.normalize(), h2["payment"]).reindex([True]).fillna(0).astype(int)

    # ---------------- write ----------------
    wb = Workbook()
    F = lambda **k: Font(**{"name": "Arial", "size": 10, **k})
    HF = PatternFill("solid", start_color="DDE4EE"); PF = PatternFill("solid", start_color="F2F2EE")
    thin = Side(style="thin", color="9A9A9A")
    ws = wb.active; ws.title = "after_transfer"
    ws["A1"] = "Spending right after each cash transfer: treatment vs control, Round 2 (payments made) and Round 1 (same weeks, no payment)"
    ws["A1"].font = F(bold=True, size=12)
    ws["A2"] = ("Per acre of rice, averaged over all households with rice area (households without the item count as 0). Activity dates. "
                "DiD = change from Round 1 to Round 2, treatment minus control, households in both rounds. Bold p-values are below 0.10.")
    ws["A2"].font = F(italic=True, color="4D524A")
    hdr = ["Outcome", "Treatment", "Control", "T − C", "p", "Treatment", "Control", "T − C", "p", "DiD (T − C)", "p", "Panel N"]
    ws.append([]); ws.append(["", "Round 2, Virippu 2026 (payments made)", "", "", "", "Round 1, Virippu 2025 (no payment: placebo)", "", "", "",
                              "Difference-in-differences", "", ""])
    ws.merge_cells("B4:E4"); ws.merge_cells("F4:I4"); ws.merge_cells("J4:L4")
    ws.append(hdr)
    for r in (4, 5):
        for c in ws[r]:
            c.font = F(bold=True); c.fill = HF; c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["A5"].alignment = Alignment(vertical="center")
    for row in rows:
        if "panel" in row:
            ws.append([row["panel"]]); rr = ws.max_row
            for c in range(1, 13):
                ws.cell(rr, c).fill = PF
            ws.cell(rr, 1).font = F(bold=True)
            continue
        ws.append([row["Outcome"], row["R2 T"], row["R2 C"], row["R2 diff"], row["R2 p"], row["R1 T"], row["R1 C"], row["R1 diff"], row["R1 p"],
                   row["DiD"], row["DiD p"], row["n"]])
        rr = ws.max_row
        pct = row["kind"] == "any"
        for c in range(1, 13):
            cell = ws.cell(rr, c); cell.font = F()
            if c in (2, 3, 6, 7):
                cell.number_format = "0%" if pct else "#,##0"
            elif c in (4, 8, 10):
                cell.number_format = '+0" pp";-0" pp";0" pp"' if pct else "+#,##0;-#,##0;0"
                if pct and cell.value is not None:
                    cell.value = cell.value * 100
            elif c in (5, 9, 11):
                cell.number_format = "0.000"
                if cell.value is not None and not pd.isna(cell.value) and cell.value < 0.10:
                    cell.font = F(bold=True)
                if cell.value is not None and pd.isna(cell.value):
                    cell.value = None
        ws.cell(rr, 1).alignment = Alignment(wrap_text=True, vertical="center")
    for col, w in zip("ABCDEFGHIJKL", [52, 11, 11, 10, 8, 11, 11, 10, 8, 12, 8, 9]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "B6"
    ws.row_dimensions[4].height = 30
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0; ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.cell(ws.max_row + 2, 1, f"Households with rice area: Round 2 treatment {int((h2.payment == 1).sum())}, control {int((h2.payment == 0).sum())}; "
            f"Round 1 treatment {int((h1.payment == 1).sum())}, control {int((h1.payment == 0).sum())}. p-values: Welch t-test (amounts, DiD), "
            "Fisher's exact test (shares in each round). Treatment = Payment 'Yes' in RCT_payment_round1.").font = F(italic=True, color="4D524A")

    ws = wb.create_sheet("timing")
    ws["A1"] = "Date of the first purchase / application, by group (median)"; ws["A1"].font = F(bold=True, size=12)
    ws.append([]); ws.append(["Round", "Activity", "N treatment", "Median date, treatment", "N control", "Median date, control", "Difference (days, T − C)", "p (Mann–Whitney)"])
    for c in ws[3]:
        c.font = F(bold=True); c.fill = HF; c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    for t in tim:
        ws.append(t)
        for c in ws[ws.max_row]:
            c.font = F()
        ws.cell(ws.max_row, 7).number_format = "+0;-0;0"; ws.cell(ws.max_row, 8).number_format = "0.000"
    for col, w in zip("ABCDEFGH", [16, 22, 12, 14, 11, 14, 14, 12]):
        ws.column_dimensions[col].width = w
    ws.row_dimensions[3].height = 30

    ws = wb.create_sheet("season_profile")
    ws["A1"] = "When the Virippu 2026 money was spent (Round 2, cost per acre of the average household)"; ws["A1"].font = F(bold=True, size=12)
    ws.append([]); ws.append(["Period (activity date)", "Rs per acre", "Share of season"])
    for c in ws[3]:
        c.font = F(bold=True); c.fill = HF
    for p in prof:
        ws.append(p)
        ws.cell(ws.max_row, 2).number_format = "#,##0"; ws.cell(ws.max_row, 3).number_format = "0%"
        for c in ws[ws.max_row]:
            c.font = F(bold=p[0] == "Season total")
    ws.append([]); ws.append(["How items were paid (Round 2)", "", "Share"]); ws.cell(ws.max_row, 1).font = F(bold=True)
    for c in credit:
        ws.append([c[0], None, c[1]]); ws.cell(ws.max_row, 3).number_format = "0.0%"
    ws.column_dimensions["A"].width = 58; ws.column_dimensions["B"].width = 13; ws.column_dimensions["C"].width = 15

    ws = wb.create_sheet("receipt_dates")
    ws["A1"] = "Reported date of receiving the Rs 2,000 from the survey team (Round 2, pmkisan_survey_date), by group"; ws["A1"].font = F(bold=True, size=12)
    ws.append([]); ws.append(["Reported date", "Treatment (payment = 1)", "Control (payment = 0)"])
    for c in ws[3]:
        c.font = F(bold=True); c.fill = HF
    for b, r in rec.iterrows():
        ws.append([b, int(r.get(1, 0)), int(r.get(0, 0))])
    ws.append(["Of the dated reports: date equals the interview date", int(same_day.get(1, pd.Series([0])).iloc[0]), int(same_day.get(0, pd.Series([0])).iloc[0])])
    ws.column_dimensions["A"].width = 52; ws.column_dimensions["B"].width = 22; ws.column_dimensions["C"].width = 22

    ns = wb.create_sheet("notes")
    notes = [
        ("Question", "Did households that received cash spend more on the activity that comes right after the transfer? The 1st payment (1st week of August) "
                     "precedes pesticide application and the 3rd fertiliser dose; the 2nd payment (1st week of September) precedes late crop protection and harvesting."),
        ("Windows", "Spending by activity date in the three weeks after each payment week (1–21 August, 1–21 September), before 1 August, and the whole season. "
                    "The same calendar weeks in Virippu 2025 (Round 1, before any payment) serve as the pre-payment check."),
        ("Round 2 harvest", "Interviews took place 17–30 September 2026, so most harvesting costs were reported as planned. These are left out of the windows "
                            "but included in the whole-season row."),
        ("DiD", "Households in both rounds (matched on farmer ID). The estimate equals a household fixed-effects regression with two periods."),
        ("Treatment contrast", "141 of 158 control households also report receiving Rs 2,000 from the survey team; see 'receipt_dates'. If control households were paid later "
                               "than treatment households, the comparison measures the effect of having the cash at the time of the activity."),
        ("Source", "Item-level costs: quantity × price per unit (or amount paid) ÷ rice area of that season. Items flagged as entry errors and undated items excluded."),
    ]
    ns.append(["Item", "Note"])
    for n in notes:
        ns.append(list(n))
    for c in ns[1]:
        c.font = F(bold=True); c.fill = HF
    for row in ns.iter_rows(min_row=2):
        for c in row:
            c.font = F(); c.alignment = Alignment(wrap_text=True, vertical="top")
    ns.column_dimensions["A"].width = 18; ns.column_dimensions["B"].width = 120
    wb.save(OUT)
    print(pd.DataFrame([r for r in rows if "panel" not in r]).round(3).to_string())
    print(tim); print(prof); print(credit); print(rec); print(same_day)


# =====================================================================================
# STEP 9a — household consumption per person, treatment vs control, both rounds (results)
# =====================================================================================

def household_expenditure(R1, R2):
    """Household consumption expenditure (Section 4, last 30 days) per person: treatment vs control in Round 1
    (baseline, interviews July 2026, before any transfer) and Round 2 (interviews 17-30 Sep 2026, after the
    transfers), panel difference-in-differences and ANCOVA. Round 2 did not ask household size, so the Round 1
    household size of the same farmer is used in both rounds. (from hh_exp.py)"""
    ITEMS = [("exp_food_amt", "Food"), ("exp_util_amt", "Housing & utilities"), ("exp_comm_amt", "Transport & communication"),
             ("exp_health_amt", "Health"), ("exp_edu_amt", "Education"), ("exp_personal_amt", "Personal & household goods"),
             ("exp_misc_amt", "Miscellaneous (religious, social, other)")]
    IV = [v for v, _ in ITEMS]
    DK = [999, 9999]                                     # "don't know" codes (Round 1 only)

    d1 = pd.read_excel(R1, sheet_name="data", usecols=["_index", "farmers_id", "payment", "interview_date", "hh_members",
                                                        "working_age", "children", "migrants", "rice_acres_2025", "exp_total"] + IV)
    d2 = pd.read_excel(R2, sheet_name="data", usecols=["_index", "farmers_id", "payment", "interview_date", "rice_acres_2026",
                                                        "exp_total", "exp_total_onam"] + IV)
    for d in (d1, d2):
        d["fid"] = d["farmers_id"].astype(str).str.strip()
        for v in IV + ["exp_total"]:
            d[v] = pd.to_numeric(d[v], errors="coerce").where(lambda s: ~s.isin(DK))
        d["exp_sum"] = d[IV].sum(axis=1, min_count=len(IV))          # missing if any item missing / don't know
    log = []

    # ---------------- household size (Round 1) ----------------
    d1["hh_size"] = d1["hh_members"].where(d1["hh_members"] <= 20)  # 58 and 61 are entry errors
    log.append(("Household size", "Round 1 hh_members; 2 values (58, 61; farmers 157, 136) set to missing. Round 2 did not ask household size, "
                "so Round 1 size is used for the same farmer in Round 2."))
    incons = d1["hh_members"] < d1["working_age"] + d1["children"]
    log.append(("Household size check", f"{int(incons.sum())} households report fewer members than working-age adults + children "
                f"(farmers {', '.join(d1.loc[incons, 'fid'])}); kept as reported, alternative size used in robustness."))
    d1["hh_alt"] = np.maximum(d1["hh_size"], d1["working_age"] + d1["children"]).where(d1["hh_size"].notna())
    d1["hh_res"] = (d1["hh_size"] - d1["migrants"]).clip(lower=1)   # members living in the house

    # Round 1 interview after the transfers: recall window not a baseline
    late = pd.to_datetime(d1["interview_date"]) >= "2026-08-01"
    log.append(("Round 1 late interview", f"{int(late.sum())} Round 1 interview after 1 Aug 2026 (farmer {', '.join(d1.loc[late, 'fid'])}, "
                f"{pd.to_datetime(d1.loc[late, 'interview_date']).dt.strftime('%d %b %Y').tolist()}): its 30-day recall falls after the transfers, so it is left out."))
    d1 = d1[~late].copy()

    # ---------------- match Round 2 to Round 1 ----------------
    # Round 2 has 4 duplicated farmer IDs (different households). Match the row whose rice area is closest to the
    # Round 1 rice area; leave the ID unmatched if both rows are equally close.
    r1 = d1.set_index("fid")
    d2["match"] = False
    for fid, g in d2.groupby("fid"):
        if fid not in r1.index:
            continue
        if len(g) == 1:
            d2.loc[g.index, "match"] = True
            continue
        dist = (g["rice_acres_2026"] - r1.at[fid, "rice_acres_2025"]).abs()
        if (dist == dist.min()).sum() == 1:
            d2.loc[dist.idxmin(), "match"] = True
    dups = d2[d2["fid"].duplicated(keep=False)].sort_values("fid")
    log.append(("Round 2 duplicate IDs", "; ".join(f"{fid}: _index " + ", ".join(f"{int(r['_index'])} ({r['rice_acres_2026']} ac{', matched' if r['match'] else ''})" for _, r in g.iterrows())
                                                  + f" vs Round 1 {r1.at[fid, 'rice_acres_2025']} ac" for fid, g in dups.groupby("fid"))
                + ". Matched by closest rice area; ties left unmatched. Unmatched duplicates have no household size, so they enter household totals only."))
    unm = d2.loc[~d2["fid"].isin(r1.index), "fid"].tolist()
    log.append(("Round 2 IDs not in Round 1", f"{len(unm)}: {', '.join(unm)} (no household size; household totals only)."))
    for c in ("hh_size", "hh_alt", "hh_res"):
        d2[c] = np.where(d2["match"], d2["fid"].map(r1[c]), np.nan)

    for d in (d1, d2):
        for v in IV + ["exp_sum", "exp_total"]:
            d[f"pp_{v}"] = d[v] / d["hh_size"]
        d["pp_exp_sum_alt"] = d["exp_sum"] / d["hh_alt"]
        d["pp_exp_sum_res"] = d["exp_sum"] / d["hh_res"]
        d["pp_exp_sum_nohealth"] = (d["exp_sum"] - d["exp_health_amt"]) / d["hh_size"]
        d["ln_pp_exp_sum"] = np.log(d["pp_exp_sum"])
    d2["pp_exp_total_onam"] = d2["exp_total_onam"] / d2["hh_size"]
    for d in (d1, d2):                                    # top 1% winsorized within the round
        cap = d["pp_exp_sum"].quantile(0.99)
        d["pp_exp_sum_w"] = d["pp_exp_sum"].clip(upper=cap)

    # ---------------- tests ----------------
    def ols_hc1(y, X):
        X = np.column_stack([np.ones(len(y))] + X)
        b = np.linalg.lstsq(X, y, rcond=None)[0]; e = y - X @ b
        n, k = X.shape; XtXi = np.linalg.inv(X.T @ X)
        V = XtXi @ (X.T * e ** 2) @ X @ XtXi * n / (n - k)
        se = np.sqrt(np.diag(V)); p = 2 * stats.t.sf(np.abs(b / se), n - k)
        return b, se, p

    def compare(d, v):
        x = d[[v, "payment"]].dropna()
        a, c = x.loc[x.payment == 1, v], x.loc[x.payment == 0, v]
        return {"N_T": len(a), "mean_T": a.mean(), "sd_T": a.std(), "med_T": a.median(), "N_C": len(c), "mean_C": c.mean(), "sd_C": c.std(),
                "med_C": c.median(), "diff": a.mean() - c.mean(), "pct": (a.mean() - c.mean()) / c.mean() if c.mean() else np.nan,
                "p_t": stats.ttest_ind(a, c, equal_var=False).pvalue, "p_mw": stats.mannwhitneyu(a, c).pvalue,
                "nd": (a.mean() - c.mean()) / np.sqrt((a.var() + c.var()) / 2)}

    def panel(v):
        p = d2.loc[d2["match"], ["fid", "payment", v]].merge(d1[["fid", "payment", v]], on="fid", suffixes=("_2", "_1")).dropna()
        t = p["payment_2"].values.astype(float)
        ch = p[f"{v}_2"] - p[f"{v}_1"]
        b, se, pv = ols_hc1(ch.values, [t])                               # DiD = household fixed effects, 2 periods
        b2, se2, pv2 = ols_hc1(p[f"{v}_2"].values, [t, p[f"{v}_1"].values])  # ANCOVA (McKenzie 2012)
        return {"N_panel": len(p), "N_panel_T": int(t.sum()), "chg_T": ch[t == 1].mean(), "chg_C": ch[t == 0].mean(),
                "did": b[1], "did_se": se[1], "did_p": pv[1], "anc": b2[1], "anc_se": se2[1], "anc_p": pv2[1],
                "switch": int((p["payment_2"] != p["payment_1"]).sum())}

    OUTCOMES = [("pp_exp_sum", "TOTAL consumption per person (sum of 7 items)")] + [(f"pp_{v}", l) for v, l in ITEMS]
    HH = [("exp_sum", "TOTAL household consumption (sum of 7 items)")] + [(v, l) for v, l in ITEMS]
    ROB = [("pp_exp_sum", "Main: sum of 7 items ÷ household members"),
           ("pp_exp_sum_w", "Top 1% winsorized within each round"),
           ("ln_pp_exp_sum", "Log of per-person total (difference ≈ % gap)"),
           ("pp_exp_sum_nohealth", "Excluding health (lumpy, e.g. surgery)"),
           ("pp_exp_sum_alt", "Size = max(members, working-age + children)"),
           ("pp_exp_sum_res", "Size = resident members (members − migrants)"),
           ("pp_exp_total", "Respondent's own total (question 8) ÷ members")]

    def table(rows):
        out = []
        for v, lab in rows:
            a, b, p = compare(d1, v), compare(d2, v), panel(v)
            out.append({"var": v, "label": lab, **{f"R1_{k}": x for k, x in a.items()}, **{f"R2_{k}": x for k, x in b.items()}, **p})
        return pd.DataFrame(out)

    pp, hh, rob = table(OUTCOMES), table(HH), table(ROB)
    onam = pd.DataFrame([{"label": "Onam spending, household (question 9)", **compare(d2, "exp_total_onam")},
                         {"label": "Onam spending per person", **compare(d2, "pp_exp_total_onam")}])
    # 30-day recall window vs Onam (26 Aug 2026)
    iv2 = pd.to_datetime(d2["interview_date"])
    onam_in = (iv2 - pd.Timedelta(days=30) <= pd.Timestamp("2026-08-26"))
    log.append(("Round 2 recall window", f"Interviews {iv2.min():%d %b}–{iv2.max():%d %b %Y}; the 30 days before the interview include Onam (Thiruvonam, 26 Aug 2026) "
                f"for {onam_in.mean():.0%} of households. Onam spending was asked separately (question 9); whether respondents also counted it in the 30-day items is not known."))
    iv1 = pd.to_datetime(d1["interview_date"])
    log.append(("Round 1 recall window", f"Interviews {iv1.min():%d %b}–{iv1.max():%d %b %Y}: the 30 days before the interview fall in June–July 2026, before the first transfer (1st week of August)."))

    # distribution sheet
    dist = []
    for rnd, d in (("Round 1", d1), ("Round 2", d2)):
        for g, gl in ((1, "Treatment"), (0, "Control")):
            x = d.loc[d.payment == g]
            for v, lab in [("hh_size", "Household members (Round 1)"), ("exp_sum", "Household total, Rs/30 days"), ("pp_exp_sum", "Per person total, Rs/30 days")]:
                s = x[v].dropna()
                dist.append([rnd, gl, lab, len(s), s.mean(), s.std(), s.quantile(.25), s.median(), s.quantile(.75), s.min(), s.max()])
    dist = pd.DataFrame(dist, columns=["Round", "Group", "Variable", "N", "Mean", "SD", "P25", "Median", "P75", "Min", "Max"])

    # household-level data
    keep = ["fid", "payment", "hh_size"] + IV + ["exp_sum", "pp_exp_sum", "exp_total"]
    hd = d1[keep].merge(d2.loc[d2["match"], ["fid", "payment"] + IV + ["exp_sum", "pp_exp_sum", "exp_total", "exp_total_onam"]],
                        on="fid", how="outer", suffixes=("_r1", "_r2"))
    hd2u = d2.loc[~d2["match"], ["fid", "payment"] + IV + ["exp_sum", "exp_total", "exp_total_onam"]].add_suffix("_r2").rename(columns={"fid_r2": "fid"})
    hd = pd.concat([hd, hd2u.assign(note="Round 2 row not matched to Round 1")], ignore_index=True)

    R = {"pp": pp, "hh": hh, "rob": rob, "onam": onam, "dist": dist, "log": log, "hd": hd,
         "N": {"r1": d1["payment"].value_counts().to_dict(), "r2": d2["payment"].value_counts().to_dict(),
               "r1_pp": d1.dropna(subset=["pp_exp_sum"])["payment"].value_counts().to_dict(),
               "r2_pp": d2.dropna(subset=["pp_exp_sum"])["payment"].value_counts().to_dict()}}
    pd.set_option("display.width", 250); pd.set_option("display.max_columns", 40)
    cols = ["label", "R1_mean_T", "R1_mean_C", "R1_diff", "R1_p_t", "R1_p_mw", "R2_mean_T", "R2_mean_C", "R2_diff", "R2_pct", "R2_p_t", "R2_p_mw",
            "N_panel", "did", "did_p", "anc", "anc_p", "switch"]
    for nm, t in (("PER PERSON", pp), ("HOUSEHOLD", hh), ("ROBUSTNESS", rob)):
        print("=====", nm); print(t[cols].round(3).to_string(index=False))
    print(onam.round(3).to_string())
    print(dist.round(0).to_string())
    for k, v in log:
        print("-", k, ":", v)
    return R


# =====================================================================================
# STEP 9b — household consumption per person  ->  Household_expenditure_per_person_round1_round2.xlsx
# =====================================================================================

def household_expenditure_excel(R, OUT):
    """Excel output for the household consumption comparison (hh_exp.py results). (from hh_exp_excel.py)"""
    F = lambda **k: Font(**{"name": "Arial", "size": 10, **k})
    HF = PatternFill("solid", start_color="DDE4EE"); TF = PatternFill("solid", start_color="F2F2EE")
    thin = Side(style="thin", color="9A9A9A")
    wb = Workbook()

    def money(c, signed=False):
        c.number_format = "+#,##0;-#,##0;0" if signed else "#,##0"

    def pval(c):
        c.number_format = "0.000"
        if c.value is not None and not pd.isna(c.value) and c.value < 0.10:
            c.font = F(bold=True)

    def main_sheet(ws, df, title, sub, log_row=None):
        ws["A1"] = title; ws["A1"].font = F(bold=True, size=12)
        ws["A2"] = sub; ws["A2"].font = F(italic=True, color="4D524A")
        groups = [("Round 1 (baseline: interviews July 2026, before any transfer)", 2, 10),
                  ("Round 2 (interviews 17–30 Sep 2026, after the transfers)", 11, 20),
                  ("Change from Round 1 to Round 2 (households in both rounds)", 21, 27)]
        for lab, a, b in groups:
            ws.cell(4, a, lab)
            ws.merge_cells(start_row=4, start_column=a, end_row=4, end_column=b)
        hdr = ["30-day consumption, Rs", "N T", "N C", "Mean T", "Mean C", "Median T", "Median C", "T − C (mean)", "p (t-test)", "p (Mann–Whitney)",
               "N T", "N C", "Mean T", "Mean C", "Median T", "Median C", "T − C (mean)", "T − C (% of C)", "p (t-test)", "p (Mann–Whitney)",
               "N", "DiD: T − C change", "SE", "p", "ANCOVA: T − C", "SE", "p"]
        for j, h in enumerate(hdr, 1):
            ws.cell(5, j, h)
        for r in (4, 5):
            for j in range(1, 28):
                c = ws.cell(r, j); c.font = F(bold=True); c.fill = HF
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.cell(5, 1).alignment = Alignment(vertical="center")
        r = 6
        for _, x in df.iterrows():
            vals = [x["label"], x["R1_N_T"], x["R1_N_C"], x["R1_mean_T"], x["R1_mean_C"], x["R1_med_T"], x["R1_med_C"], x["R1_diff"], x["R1_p_t"], x["R1_p_mw"],
                    x["R2_N_T"], x["R2_N_C"], x["R2_mean_T"], x["R2_mean_C"], x["R2_med_T"], x["R2_med_C"], x["R2_diff"], x["R2_pct"], x["R2_p_t"], x["R2_p_mw"],
                    x["N_panel"], x["did"], x["did_se"], x["did_p"], x["anc"], x["anc_se"], x["anc_p"]]
            for j, v in enumerate(vals, 1):
                c = ws.cell(r, j, None if (isinstance(v, float) and np.isnan(v)) else v); c.font = F(bold=(r == 6 and j == 1))
                if j in (4, 5, 6, 7, 13, 14, 15, 16, 23, 26):
                    money(c)
                elif j in (8, 17, 22, 25):
                    money(c, True)
                elif j == 18:
                    c.number_format = "+0%;-0%;0%"
                elif j in (9, 10, 19, 20, 24, 27):
                    pval(c)
            if r == 6:
                for j in range(1, 28):
                    ws.cell(r, j).fill = TF
            ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="center")
            r += 1
        for j in range(1, 28):
            ws.cell(r - 1, j).border = Border(bottom=thin)
        widths = [40] + [7, 7, 9, 9, 9, 9, 10, 8, 9] + [7, 7, 9, 9, 9, 9, 10, 9, 8, 9] + [6, 11, 8, 7, 11, 8, 7]
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = w
        ws.row_dimensions[4].height = 30; ws.row_dimensions[5].height = 30
        ws.freeze_panes = "B6"
        ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        return r

    # ---------------- per person ----------------
    ws = wb.active; ws.title = "per_person"
    r = main_sheet(ws, R["pp"], "Household consumption per person (Rs, last 30 days): treatment vs control, Round 1 and Round 2",
                   "Per person = household spending ÷ household members (Round 1). T = treatment (payment = 1), C = control. "
                   "DiD = household's change from Round 1 to Round 2, T minus C. ANCOVA = Round 2 value on treatment, controlling for the Round 1 value. "
                   "Robust (HC1) SEs. Bold p-values are below 0.10.")
    ws.cell(r + 1, 1, "Reading the table: means are pulled up by a few very large health and miscellaneous amounts in Round 2; see 'robustness' "
            "for winsorized, log and excluding-health versions.").font = F(italic=True, color="4D524A")

    # ---------------- household ----------------
    ws = wb.create_sheet("household")
    main_sheet(ws, R["hh"], "Household consumption (Rs per household, last 30 days): treatment vs control, Round 1 and Round 2",
               "Not divided by household size. Includes Round 2 households without a Round 1 match. Bold p-values are below 0.10.")

    # ---------------- robustness ----------------
    ws = wb.create_sheet("robustness")
    ws["A1"] = "Total consumption per person: alternative definitions"; ws["A1"].font = F(bold=True, size=12)
    ws["A2"] = "Log rows: differences are in log points (×100 ≈ % difference). 95% CI = estimate ± 1.97 × SE. MDE = minimum detectable effect at 80% power, 5% level (2.8 × SE)."
    ws["A2"].font = F(italic=True, color="4D524A")
    hdr = ["Definition", "Round 1: T − C", "p", "Round 2: T − C", "p", "N panel", "DiD", "SE", "p", "95% CI low", "95% CI high", "MDE (DiD)", "ANCOVA", "SE", "p"]
    ws.append([]); ws.append(hdr)
    for c in ws[4]:
        c.font = F(bold=True); c.fill = HF; c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["A4"].alignment = Alignment(vertical="center")
    for _, x in R["rob"].iterrows():
        lg = x["var"].startswith("ln_")
        ws.append([x["label"], x["R1_diff"], x["R1_p_t"], x["R2_diff"], x["R2_p_t"], x["N_panel"], x["did"], x["did_se"], x["did_p"],
                   x["did"] - 1.97 * x["did_se"], x["did"] + 1.97 * x["did_se"], 2.8 * x["did_se"], x["anc"], x["anc_se"], x["anc_p"]])
        rr = ws.max_row
        for j in range(1, 16):
            c = ws.cell(rr, j); c.font = F()
            if j in (3, 5, 9, 15):
                pval(c)
            elif j != 1 and j != 6:
                c.number_format = ("+0.000;-0.000;0.000" if j in (2, 4, 7, 10, 11, 13) else "0.000") if lg else ("+#,##0;-#,##0;0" if j in (2, 4, 7, 10, 11, 13) else "#,##0")
        ws.cell(rr, 1).alignment = Alignment(wrap_text=True, vertical="center")
    for j, w in enumerate([46, 13, 7, 13, 7, 8, 10, 9, 7, 11, 11, 11, 10, 9, 7], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[4].height = 30

    # ---------------- Onam ----------------
    ws = wb.create_sheet("onam")
    ws["A1"] = "Spending during Onam (Round 2, question 9: 'Total household consumption expenditure during Onam')"; ws["A1"].font = F(bold=True, size=12)
    ws["A2"] = "Thiruvonam was on 26 August 2026, after the 1st transfer. Per person uses Round 1 household size."
    ws["A2"].font = F(italic=True, color="4D524A")
    ws.append([]); ws.append(["", "N T", "Mean T", "Median T", "N C", "Mean C", "Median C", "T − C (mean)", "T − C (% of C)", "p (t-test)", "p (Mann–Whitney)"])
    for c in ws[4]:
        c.font = F(bold=True); c.fill = HF; c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for _, x in R["onam"].iterrows():
        ws.append([x["label"], x["N_T"], x["mean_T"], x["med_T"], x["N_C"], x["mean_C"], x["med_C"], x["diff"], x["pct"], x["p_t"], x["p_mw"]])
        rr = ws.max_row
        for j in range(1, 12):
            c = ws.cell(rr, j); c.font = F()
            if j in (3, 4, 6, 7):
                money(c)
            elif j == 8:
                money(c, True)
            elif j == 9:
                c.number_format = "+0%;-0%;0%"
            elif j in (10, 11):
                pval(c)
    for j, w in enumerate([40, 7, 10, 10, 7, 10, 10, 12, 12, 10, 12], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[4].height = 30

    # ---------------- distribution ----------------
    ws = wb.create_sheet("distribution")
    ws["A1"] = "Distribution by round and group"; ws["A1"].font = F(bold=True, size=12)
    ws.append([]); ws.append(list(R["dist"].columns))
    for c in ws[3]:
        c.font = F(bold=True); c.fill = HF
    for row in R["dist"].itertuples(index=False):
        ws.append(list(row))
        for j in range(1, 12):
            c = ws.cell(ws.max_row, j); c.font = F()
            if j >= 5:
                c.number_format = "#,##0.0" if "members" in row[2] else "#,##0"
    for j, w in enumerate([10, 11, 30, 6, 10, 10, 10, 10, 10, 10, 10], 1):
        ws.column_dimensions[get_column_letter(j)].width = w

    # ---------------- household data ----------------
    ws = wb.create_sheet("household_data")
    hd = R["hd"].copy()
    ws.append(list(hd.columns))
    for c in ws[1]:
        c.font = F(bold=True); c.fill = HF
    for row in hd.itertuples(index=False):
        ws.append([None if (isinstance(v, float) and np.isnan(v)) else v for v in row])
    ws.freeze_panes = "B2"
    for j in range(1, len(hd.columns) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 13

    # ---------------- notes ----------------
    ns = wb.create_sheet("notes")
    N = R["N"]
    notes = [
        ("Outcome", "Section 4, 'During the past 30 days, how much did your household spend on (or consume from own production/gifts of)': food; housing & utilities; "
                    "transport & communication; health; education; personal & household goods; miscellaneous. Total = sum of the 7 items (missing if any item is "
                    "missing or 'don't know'). The respondent's own total (question 8) rarely equals the sum of items and is used only as a robustness check."),
        ("Don't know", "999 / 9999 set to missing (Round 1 only; Round 2 has none)."),
        ("Per person", "Household spending ÷ household members (Round 1, question 11)."),
        ("Groups", f"Treatment = Payment 'Yes' in RCT_payment_round1 (payment = 1). Households with a per-person total: Round 1 treatment {N['r1_pp'].get(1)}, control {N['r1_pp'].get(0)}; "
                   f"Round 2 treatment {N['r2_pp'].get(1)}, control {N['r2_pp'].get(0)}. Payment status is the same for every household in both rounds."),
        ("Tests", "Welch t-test and Mann–Whitney test within each round. DiD: OLS of the change (Round 2 − Round 1) on treatment, equal to a household fixed-effects "
                  "regression with two periods. ANCOVA: OLS of the Round 2 value on treatment and the Round 1 value (more precise than DiD when outcomes are noisy). "
                  "Robust (HC1) standard errors."),
        ("Size of the transfer", "Rs 2,000 per household is about Rs 500 per person for a household of 4, about 6% of median Round 2 spending per person (Rs 7,650–8,270). "
                                 "The smallest effect the study can reliably detect (MDE) on spending per person is about 20% (log DiD) or Rs 2,460 (winsorized DiD), "
                                 "so an effect the size of the transfer could not be detected."),
    ] + [(k, v) for k, v in R["log"]]
    ns.append(["Item", "Note"])
    for n in notes:
        ns.append(list(n))
    for c in ns[1]:
        c.font = F(bold=True); c.fill = HF
    for row in ns.iter_rows(min_row=2):
        for c in row:
            c.font = F(); c.alignment = Alignment(wrap_text=True, vertical="top")
    ns.column_dimensions["A"].width = 24; ns.column_dimensions["B"].width = 130
    wb.save(OUT)
    print("saved", OUT)


# =====================================================================================
# RUN ALL STEPS
# =====================================================================================
def main(data_dir=DATA_DIR, r1_dir=R1_DIR, out_dir=OUT_DIR):
    data_dir, r1_dir, out_dir = Path(data_dir), Path(r1_dir), Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    r1_data, r1_items = r1_dir / "Survey_Round_1_merged_wide.xlsx", r1_dir / "items.csv"
    for f in (r1_data, r1_items):
        if not f.exists():
            raise FileNotFoundError(f"{f} not found: run RCT_round1_analysis.py first (or set R1_DIR).")
    merged = out_dir / "Survey_round_2_merged_wide.xlsx"
    items2 = out_dir / "items2.csv"
    print("Step 1: merging sheets and computing costs")
    merge_round2(data_dir, data_dir / "RCT_payment_round1.xls", r1_data, merged)
    print("Step 2: item-level costs and weekly timeline data")
    timeline_items2(merged, out_dir / "timeline_weekly_round2.json")          # also writes items2.csv
    print("Step 3: cost comparison")
    cost_comparison2(merged, out_dir / "Round2_cost_comparison_payment.xlsx", items2)
    print("Step 4: bubble plots")
    for mode in ("overlay", "panels"):
        bubble_plot2(items2, merged, out_dir / f"Round2_cost_per_acre_bubble_{mode}.jpg", "act", mode)
    print("Step 5: pesticide vs 3rd fertiliser dose")
    pesticide_fert3(items2, merged, out_dir / "Round2_pesticide_vs_3rd_fertiliser.jpg")
    print("Step 6: pesticide and 3rd dose, Round 1 vs Round 2")
    pesticide_fert3_rounds(r1_items, r1_data, items2, merged, out_dir / "Pesticide_3rd_fertiliser_round1_vs_round2.jpg")
    print("Step 7: PM-KISAN use pies and workbook")
    jpg2, csv2 = out_dir / "Round2_pmkisan_survey_use_pies.jpg", out_dir / "pmkisan_survey_use_categories_round2.csv"
    jpg1, csv1 = out_dir / "Round1_pmkisan_use_pie.jpg", out_dir / "pmkisan_use_categories_round1.csv"
    pmkisan_pies_round2(merged, jpg2, csv2)
    pmkisan_pie_round1(r1_data, jpg1, csv1)
    xlsx = out_dir / "PMKISAN_use_pies_round1_round2.xlsx"
    pmkisan_excel(csv2, xlsx, jpg2)
    pmkisan_excel_add_round1(xlsx, csv1, jpg1)
    print("Step 8: spending right after each transfer")
    spending_after_transfer(r1_items, r1_data, items2, merged, out_dir / "Round2_spending_after_transfer.xlsx")
    print("Step 9: household consumption per person")
    res = household_expenditure(r1_data, merged)
    household_expenditure_excel(res, out_dir / "Household_expenditure_per_person_round1_round2.xlsx")
    print(f"Done. Outputs in {out_dir.resolve()}")


if __name__ == "__main__":
    if len(sys.argv) == 4:
        main(sys.argv[1], sys.argv[2], sys.argv[3])
    else:
        main()

"""Generate offline SEC-shaped fixtures for FICTIONAL companies, used by tests and demo mode.

The JSON mirrors the real EDGAR `companyfacts` format and deliberately includes real-world messiness:
tag switches (SalesRevenueNet -> ASC 606 tag), restatements, duplicate facts across filings, a missing
GrossProfit tag, multi-class cover-page shares, a company that stopped tagging cover-page shares, a
January fiscal year-end, and 10-Q balance sheets after the last 10-K.
"""
import json
from pathlib import Path

import numpy as np
import pandas as pd

OUT = Path(__file__).parent / "fixtures"
RNG = np.random.default_rng(7)

COMPANIES = [
    # ticker, cik, name, fy_end_month, rev0 ($bn), growth path, gm, ebit margin path, debt, cash, shares (m), beta, quirks
    ("DEMO", 900001, "Demo Athletic Brands Inc.", 5, 30.0, [.07, .06, -.05, .19, .05, .10, .00, -.10, .00],
     .44, [.13, .13, .10, .16, .14, .12, .12, .08, .07], 9.5, 9.0, 1500, 1.05, {"no_dei_shares", "restatement", "no_operating_income"}),
    ("PEERA", 900002, "Peer Apparel Group", 1, 4.0, [.15, .20, .11, .42, .30, .19, .10, .08, .04],
     .56, [.18, .19, .20, .22, .21, .22, .23, .22, .20], 0.0, 1.5, 125, 1.20, {"jan_ye"}),
    ("PEERB", 900003, "Peer Footwear Holdings", 3, 1.9, [.06, .13, .08, .38, .18, .14, .18, .16, .09],
     .50, [.13, .14, .15, .19, .18, .18, .21, .23, .22], 0.0, 1.2, 28, 1.10, {"no_gross_profit"}),
    ("PEERC", 900004, "Peer Outdoor Brands", 12, 2.4, [.06, .08, .15, -.10, .25, .10, .06, -.03, .01],
     .50, [.10, .11, .12, .07, .13, .12, .10, .09, .09], 0.0, 0.8, 62, 0.95, {"revenues_tag"}),
    ("PEERD", 900005, "Peer Performance Wear", 12, 5.0, [.02, .03, .04, -.15, .27, .03, .00, -.05, -.10],
     .46, [.02, .03, .05, -.05, .07, .05, .03, .02, .01], 1.0, 1.2, 440, 1.40, {"multi_class"}),
    ("PEERE", 900006, "Peer Lifestyle Co.", 12, 2.2, [.10, .14, .12, .05, .60, .50, .11, .04, .02],
     .52, [.15, .16, .18, .22, .27, .25, .25, .24, .22], 2.0, 0.5, 60, 1.60, set()),
]
YEARS = list(range(2017, 2026))   # fiscal year labels


def fy_end(label, month):
    year = label + 1 if month <= 2 else label
    return pd.Timestamp(year=year, month=month, day=1) + pd.offsets.MonthEnd(0)


def build(c):
    tk, cik, name, m, rev0, gpath, gm, mpath, debt0, cash0, sh0, beta, quirks = c
    rows = {}   # (tag, unit) -> list of facts
    rev = rev0 * 1e9
    hist = []
    for i, fy in enumerate(YEARS):
        rev = rev * (1 + gpath[i])
        ebit = rev * mpath[i]
        da = rev * 0.025; capex = rev * 0.03
        inv = rev * (1 - gm) * 0.30; rec = rev * 0.10; pay = rev * (1 - gm) * 0.18
        debt = debt0 * 1e9 * (1 + 0.03 * i); cash = cash0 * 1e9 * (1 + 0.05 * i)
        interest = debt * 0.035
        pretax = ebit - interest + rev * 0.002
        tax = pretax * 0.19; ni = pretax - tax
        shares = sh0 * 1e6 * (1 - 0.01 * i)
        cfo = ni + da + rev * 0.01
        hist.append(dict(fy=fy, end=fy_end(fy, m), start=fy_end(fy - 1, m) + pd.Timedelta(days=1), rev=rev, cogs=rev * (1 - gm),
                         gp=rev * gm, ebit=ebit, da=da, interest=interest, pretax=pretax, tax=tax, ni=ni, cfo=cfo,
                         capex=capex, sbc=rev * 0.008, div=ni * 0.3, bb=ni * 0.3, sh=shares, eps=ni / shares,
                         cash=cash, sti=cash * 0.3, rec=rec, inv=inv, pay=pay, ca=cash + rec + inv, cl=pay * 2.2,
                         ta=rev * 0.9, tl=rev * 0.55, eq=rev * 0.35, ppe=rev * 0.15, gw=rev * 0.05,
                         ltd=debt * 0.9, ltdc=debt * 0.1))

    def add(tag, unit, fact, ns="us-gaap"):
        rows.setdefault((ns, tag, unit), []).append(fact)

    for k, h in enumerate(hist):                       # one 10-K per year, reporting 3 yrs of flows, 2 of stocks
        filed = (h["end"] + pd.Timedelta(days=55)).strftime("%Y-%m-%d")
        accn = f"0000{cik}-{str(h['end'].year + (1 if h['end'].month > 9 else 0))[2:]}-{k:06d}"
        for j in range(max(0, k - 2), k + 1):
            p = hist[j]
            restated = 1.0 + (0.004 if ("restatement" in quirks and j == k - 1) else 0)
            fl = dict(start=p["start"].strftime("%Y-%m-%d"), end=p["end"].strftime("%Y-%m-%d"), accn=accn,
                      fy=h["end"].year, fp="FY", form="10-K", filed=filed)
            rev_tag = ("Revenues" if "revenues_tag" in quirks else
                       ("SalesRevenueNet" if p["fy"] < 2018 else "RevenueFromContractWithCustomerExcludingAssessedTax"))
            add(rev_tag, "USD", {**fl, "val": round(p["rev"] * restated)})
            add("CostOfGoodsAndServicesSold", "USD", {**fl, "val": round(p["cogs"])})
            if "no_gross_profit" not in quirks: add("GrossProfit", "USD", {**fl, "val": round(p["gp"])})
            if "no_operating_income" in quirks:
                add("SellingGeneralAndAdministrativeExpense", "USD", {**fl, "val": round(p["gp"] - p["ebit"])})
            else:
                add("OperatingIncomeLoss", "USD", {**fl, "val": round(p["ebit"])})
            for tag, key in [("DepreciationDepletionAndAmortization", "da"),
                             ("InterestExpense", "interest"), ("IncomeTaxExpenseBenefit", "tax"), ("NetIncomeLoss", "ni"),
                             ("IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest", "pretax"),
                             ("NetCashProvidedByUsedInOperatingActivities", "cfo"),
                             ("PaymentsToAcquirePropertyPlantAndEquipment", "capex"), ("ShareBasedCompensation", "sbc"),
                             ("PaymentsOfDividends", "div"), ("PaymentsForRepurchaseOfCommonStock", "bb")]:
                add(tag, "USD", {**fl, "val": round(p[key])})
            add("WeightedAverageNumberOfDilutedSharesOutstanding", "shares", {**fl, "val": round(p["sh"])})
            add("EarningsPerShareDiluted", "USD/shares", {**fl, "val": round(p["eps"], 2)})
        for j in range(max(0, k - 1), k + 1):
            p = hist[j]
            fi = dict(end=p["end"].strftime("%Y-%m-%d"), accn=accn, fy=h["end"].year, fp="FY", form="10-K", filed=filed)
            for tag, key in [("CashAndCashEquivalentsAtCarryingValue", "cash"), ("ShortTermInvestments", "sti"),
                             ("AccountsReceivableNetCurrent", "rec"), ("InventoryNet", "inv"), ("AccountsPayableCurrent", "pay"),
                             ("AssetsCurrent", "ca"), ("LiabilitiesCurrent", "cl"), ("Assets", "ta"), ("Liabilities", "tl"),
                             ("StockholdersEquity", "eq"), ("PropertyPlantAndEquipmentNet", "ppe"), ("Goodwill", "gw"),
                             ("LongTermDebtNoncurrent", "ltd"), ("LongTermDebtCurrent", "ltdc")]:
                add(tag, "USD", {**fi, "val": round(p[key])})
        cover = dict(end=(h["end"] + pd.Timedelta(days=40)).strftime("%Y-%m-%d"), accn=accn, fy=h["end"].year,
                     fp="FY", form="10-K", filed=filed)
        if "no_dei_shares" in quirks and k > 2: pass
        elif "multi_class" in quirks:
            add("EntityCommonStockSharesOutstanding", "shares", {**cover, "val": round(h["sh"] * 0.45)}, ns="dei")
            add("EntityCommonStockSharesOutstanding", "shares", {**cover, "val": round(h["sh"] * 0.55)}, ns="dei")
        else:
            add("EntityCommonStockSharesOutstanding", "shares", {**cover, "val": round(h["sh"])}, ns="dei")

    # a 10-Q after the last 10-K: newer balance sheet + quarterly diluted shares
    last = hist[-1]; qend = last["end"] + pd.offsets.MonthEnd(3)
    fq = dict(end=qend.strftime("%Y-%m-%d"), accn=f"0000{cik}-26-900000", fy=qend.year, fp="Q1", form="10-Q",
              filed=(qend + pd.Timedelta(days=35)).strftime("%Y-%m-%d"))
    add("CashAndCashEquivalentsAtCarryingValue", "USD", {**fq, "val": round(last["cash"] * 0.95)})
    add("ShortTermInvestments", "USD", {**fq, "val": round(last["sti"])})
    add("LongTermDebtNoncurrent", "USD", {**fq, "val": round(last["ltd"])})
    add("LongTermDebtCurrent", "USD", {**fq, "val": round(last["ltdc"])})
    add("WeightedAverageNumberOfDilutedSharesOutstanding", "shares",
        {**fq, "start": (last["end"] + pd.Timedelta(days=1)).strftime("%Y-%m-%d"), "val": round(last["sh"] * 0.995)})

    facts = {"cik": cik, "entityName": name, "facts": {}}
    for (ns, tag, unit), lst in rows.items():
        facts["facts"].setdefault(ns, {}).setdefault(tag, {"label": tag, "units": {}})["units"][unit] = lst
    subs = {"cik": str(cik), "name": name, "sic": "3021", "sicDescription": "Rubber & Plastics Footwear",
            "fiscalYearEnd": f"{m:02d}{fy_end(2025, m).day:02d}", "tickers": [tk]}
    L = hist[-1]
    ebitda = L["ebit"] + L["da"]; net_debt = (L["ltd"] + L["ltdc"]) - L["cash"] * 0.95 - L["sti"]
    return facts, subs, (ebitda, net_debt, L["sh"])


def prices(specs):
    """specs: (ticker, beta, final price). Paths are random but rescaled to end at a price that gives a
    sensible EV/EBITDA, so the demo comps look like a real peer group."""
    idx = pd.bdate_range(end="2026-10-02", periods=3 * 252)
    mkt = RNG.normal(0.0004, 0.010, len(idx))
    out = {"SPY": 500 * np.exp(np.cumsum(mkt))}
    for tk, b, final in specs:
        path = np.exp(np.cumsum(b * mkt + RNG.normal(0, 0.012, len(idx))))
        out[tk] = path / path[-1] * final
    return pd.DataFrame(out, index=idx).round(2)

TARGET_EV_EBITDA = {"DEMO": 15.0, "PEERA": 14.0, "PEERB": 12.5, "PEERC": 10.0, "PEERD": 9.0, "PEERE": 13.0}


if __name__ == "__main__":
    OUT.mkdir(exist_ok=True)
    tick, specs = {}, []
    for c in COMPANIES:
        f, s, (ebitda, nd, sh) = build(c)
        specs.append((c[0], c[11], round((TARGET_EV_EBITDA[c[0]] * ebitda - nd) / sh, 2)))
        (OUT / f"facts_{c[1]}.json").write_text(json.dumps(f))
        (OUT / f"subs_{c[1]}.json").write_text(json.dumps(s))
        tick[c[0]] = c[1]
    (OUT / "tickers.json").write_text(json.dumps(tick))
    prices(specs).to_csv(OUT / "prices.csv")
    (OUT / "market.json").write_text(json.dumps({"risk_free": 0.042}))
    print("fixtures written to", OUT)

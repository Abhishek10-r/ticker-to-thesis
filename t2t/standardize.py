"""Turn raw XBRL company facts into a clean, comparable annual financial table, with provenance.

Companies tag the same line item differently, and switch tags over time (e.g. SalesRevenueNet ->
RevenueFromContractWithCustomerExcludingAssessedTax after ASC 606). For each line item we try a priority
list of tags *per period*, keep the most recently filed value (so restatements win), and record exactly
which tag, filing and date every number came from.
"""
from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np
import pandas as pd

ANNUAL_FORMS = {"10-K", "10-K/A", "10-KT", "10-KT/A"}

# name: (period type, unit, [tags in priority order])
ITEMS: dict[str, tuple[str, str, list[str]]] = {
    "revenue": ("duration", "USD", ["RevenueFromContractWithCustomerExcludingAssessedTax", "Revenues", "SalesRevenueNet",
                                    "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueGoodsNet"]),
    "cogs": ("duration", "USD", ["CostOfRevenue", "CostOfGoodsAndServicesSold", "CostOfGoodsSold",
                                 "CostOfGoodsAndServiceExcludingDepreciationDepletionAndAmortization"]),
    "gross_profit": ("duration", "USD", ["GrossProfit"]),
    "operating_income": ("duration", "USD", ["OperatingIncomeLoss"]),
    "d_and_a": ("duration", "USD", ["DepreciationDepletionAndAmortization", "DepreciationAndAmortization",
                                    "DepreciationAmortizationAndAccretionNet"]),
    "depreciation": ("duration", "USD", ["Depreciation"]),
    "amortization": ("duration", "USD", ["AmortizationOfIntangibleAssets"]),
    "interest_expense": ("duration", "USD", ["InterestExpense", "InterestExpenseNonoperating", "InterestExpenseDebt",
                                             "InterestAndDebtExpense", "InterestPaidNet", "InterestPaid"]),
    # used only to derive EBIT when a company doesn't tag operating income (e.g. Nike)
    "costs_and_expenses": ("duration", "USD", ["CostsAndExpenses"]),
    "operating_expenses": ("duration", "USD", ["OperatingExpenses"]),
    "sga": ("duration", "USD", ["SellingGeneralAndAdministrativeExpense"]),
    "rnd": ("duration", "USD", ["ResearchAndDevelopmentExpense"]),
    "pretax_income": ("duration", "USD", [
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesDomestic"]),
    "income_tax": ("duration", "USD", ["IncomeTaxExpenseBenefit"]),
    "net_income": ("duration", "USD", ["NetIncomeLoss", "ProfitLoss", "NetIncomeLossAvailableToCommonStockholdersBasic"]),
    "cfo": ("duration", "USD", ["NetCashProvidedByUsedInOperatingActivities",
                                "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"]),
    "capex": ("duration", "USD", ["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets",
                                  "PaymentsForCapitalImprovements"]),
    "sbc": ("duration", "USD", ["ShareBasedCompensation", "AllocatedShareBasedCompensationExpense"]),
    "dividends": ("duration", "USD", ["PaymentsOfDividends", "PaymentsOfDividendsCommonStock"]),
    "buybacks": ("duration", "USD", ["PaymentsForRepurchaseOfCommonStock"]),
    "diluted_shares": ("duration", "shares", ["WeightedAverageNumberOfDilutedSharesOutstanding"]),
    "eps_diluted": ("duration", "USD/shares", ["EarningsPerShareDiluted"]),
    "cash": ("instant", "USD", ["CashAndCashEquivalentsAtCarryingValue",
                                "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents", "Cash"]),
    "st_investments": ("instant", "USD", ["ShortTermInvestments", "MarketableSecuritiesCurrent",
                                          "AvailableForSaleSecuritiesDebtSecuritiesCurrent"]),
    "receivables": ("instant", "USD", ["AccountsReceivableNetCurrent", "ReceivablesNetCurrent"]),
    "inventory": ("instant", "USD", ["InventoryNet"]),
    "payables": ("instant", "USD", ["AccountsPayableCurrent", "AccountsPayableAndAccruedLiabilitiesCurrent"]),
    "current_assets": ("instant", "USD", ["AssetsCurrent"]),
    "current_liabilities": ("instant", "USD", ["LiabilitiesCurrent"]),
    "total_assets": ("instant", "USD", ["Assets"]),
    "total_liabilities": ("instant", "USD", ["Liabilities"]),
    "equity": ("instant", "USD", ["StockholdersEquity",
                                  "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"]),
    "ppe": ("instant", "USD", ["PropertyPlantAndEquipmentNet"]),
    "goodwill": ("instant", "USD", ["Goodwill"]),
    # debt pieces (combined in `debt` below)
    "lt_debt_total": ("instant", "USD", ["LongTermDebt", "LongTermDebtAndCapitalLeaseObligations", "DebtInstrumentCarryingAmount"]),
    "lt_debt_noncurrent": ("instant", "USD", ["LongTermDebtNoncurrent", "LongTermDebtAndCapitalLeaseObligationsNoncurrent"]),
    "lt_debt_current": ("instant", "USD", ["LongTermDebtCurrent", "DebtCurrent", "LongTermDebtAndCapitalLeaseObligationsCurrent"]),
    "st_borrowings": ("instant", "USD", ["ShortTermBorrowings", "CommercialPaper", "NotesPayableCurrent"]),
    "shares_outstanding": ("instant", "shares", ["dei:EntityCommonStockSharesOutstanding", "CommonStockSharesOutstanding"]),
}

LABELS = {
    "revenue": "Revenue", "cogs": "Cost of revenue", "gross_profit": "Gross profit", "operating_income": "Operating income (EBIT)",
    "d_and_a": "Depreciation & amortisation", "ebitda": "EBITDA", "interest_expense": "Interest expense",
    "pretax_income": "Pre-tax income", "income_tax": "Income tax", "net_income": "Net income", "cfo": "Cash from operations",
    "capex": "Capital expenditure", "fcf": "Free cash flow", "sbc": "Stock-based compensation", "dividends": "Dividends paid",
    "buybacks": "Share repurchases", "diluted_shares": "Diluted shares (weighted avg)", "eps_diluted": "Diluted EPS",
    "cash": "Cash & equivalents", "st_investments": "Short-term investments", "receivables": "Accounts receivable",
    "inventory": "Inventory", "payables": "Accounts payable", "current_assets": "Current assets",
    "current_liabilities": "Current liabilities", "total_assets": "Total assets", "total_liabilities": "Total liabilities",
    "equity": "Shareholders' equity", "ppe": "PP&E, net", "goodwill": "Goodwill", "debt": "Total debt",
    "net_debt": "Net debt", "nwc": "Operating working capital",
}


@dataclass
class Financials:
    ticker: str
    cik: int
    name: str
    sic: str
    sic_description: str
    fiscal_year_end: str
    annual: pd.DataFrame                    # index = fiscal year label (int), columns = line items (USD)
    provenance: pd.DataFrame                # one row per (item, fiscal year): tag, accession, form, filed
    latest: dict = field(default_factory=dict)   # most recent balance-sheet values (any filing) for the EV bridge
    warnings: list = field(default_factory=list)


def _fy_label(end: pd.Timestamp) -> int:
    """Fiscal year label: years ending Jan/Feb belong to the prior year (retail convention)."""
    return end.year - 1 if end.month <= 2 else end.year


def _entries(facts: dict, tag: str, unit: str) -> list[dict]:
    ns, name = ("dei", tag[4:]) if tag.startswith("dei:") else ("us-gaap", tag)
    return facts.get("facts", {}).get(ns, {}).get(name, {}).get("units", {}).get(unit, [])


def _series(facts: dict, item: str, annual_only: bool = True) -> pd.DataFrame:
    """All candidate values for one line item, one row per period end, highest-priority tag then latest filing."""
    kind, unit, tags = ITEMS[item]
    rows = []
    for rank, tag in enumerate(tags):
        for e in _entries(facts, tag, unit):
            if annual_only and e.get("form") not in ANNUAL_FORMS: continue
            if kind == "duration":
                if "start" not in e: continue
                days = (pd.Timestamp(e["end"]) - pd.Timestamp(e["start"])).days
                if annual_only and not 335 <= days <= 400: continue
            rows.append({"end": pd.Timestamp(e["end"]), "val": float(e["val"]), "tag": tag, "rank": rank,
                         "accn": e.get("accn"), "form": e.get("form"), "filed": e.get("filed", ""),
                         "start": e.get("start")})
    if not rows:
        return pd.DataFrame(columns=["end", "val", "tag", "rank", "accn", "form", "filed", "start"])
    df = pd.DataFrame(rows)
    if item == "shares_outstanding":
        # multi-class companies report one fact per class in the same filing: sum within a filing
        df = (df.groupby(["end", "accn", "tag", "rank", "form", "filed"], as_index=False)["val"].sum())
    df = df.sort_values(["end", "rank", "filed"], ascending=[True, True, False])
    return df.drop_duplicates("end", keep="first").reset_index(drop=True)


def standardize(ticker: str, source, years: int = 8) -> Financials:
    cik = source.cik(ticker)
    facts = source.company_facts(cik)
    subs = source.submissions(cik)
    warnings = []

    rev = _series(facts, "revenue")
    if rev.empty:
        raise ValueError(f"{ticker}: no annual revenue found in XBRL facts (financial companies and "
                         "foreign filers using IFRS are not supported).")
    fy_ends = sorted(rev["end"].unique())[-years:]
    index = [_fy_label(pd.Timestamp(e)) for e in fy_ends]

    data, prov = {}, []
    for item in ITEMS:
        if item == "shares_outstanding": continue
        s = _series(facts, item)
        s = s[s["end"].isin(fy_ends)].set_index("end")
        data[item] = [s.loc[e, "val"] if e in s.index else np.nan for e in fy_ends]
        for e in fy_ends:
            if e in s.index:
                r = s.loc[e]
                prov.append({"item": item, "fy": _fy_label(pd.Timestamp(e)), "period_end": pd.Timestamp(e).date(),
                             "value": r["val"], "tag": r["tag"], "form": r["form"], "accn": r["accn"], "filed": r["filed"]})
    a = pd.DataFrame(data, index=index)
    a.index.name = "fy"

    # ---- derived lines, with fallbacks ----
    a["gross_profit"] = a["gross_profit"].fillna(a["revenue"] - a["cogs"])
    a["operating_income"], how = _derive_ebit(a)
    if how: warnings.append(f"{ticker}: no OperatingIncomeLoss tag for FY {sorted(how)}; EBIT derived as "
                            f"{'; '.join(sorted(set(how.values())))}.")
    for fy, method in how.items():
        prov.append({"item": "operating_income", "fy": fy, "period_end": None, "value": a.loc[fy, "operating_income"],
                     "tag": f"DERIVED: {method}", "form": "", "accn": "", "filed": ""})
    a["d_and_a"] = a["d_and_a"].fillna(a["depreciation"] + a["amortization"].fillna(0))
    a["ebitda"] = a["operating_income"] + a["d_and_a"]
    a["capex"] = a["capex"].abs()
    a["fcf"] = a["cfo"] - a["capex"]
    a["debt"] = _combine_debt(a)
    a["net_debt"] = a["debt"] - a["cash"].fillna(0) - a["st_investments"].fillna(0)
    a["nwc"] = a["receivables"].fillna(0) + a["inventory"].fillna(0) - a["payables"].fillna(0)
    a = a.drop(columns=["depreciation", "amortization", "lt_debt_total", "lt_debt_noncurrent", "lt_debt_current",
                        "st_borrowings", "costs_and_expenses", "operating_expenses", "sga", "rnd"])

    for col, why in [("operating_income", "EBIT"), ("cfo", "cash flow"), ("d_and_a", "D&A")]:
        if a[col].isna().all(): warnings.append(f"{ticker}: no {why} data found; related metrics will be blank.")
    missing_years = a.index[a["operating_income"].isna()].tolist()
    if missing_years and not a["operating_income"].isna().all():
        warnings.append(f"{ticker}: EBIT missing for FY {missing_years}.")

    latest = _latest_balance_sheet(facts, a)
    return Financials(ticker=ticker.upper(), cik=cik, name=subs.get("name", ticker), sic=str(subs.get("sic", "")),
                      sic_description=subs.get("sicDescription", ""), fiscal_year_end=subs.get("fiscalYearEnd", ""),
                      annual=a, provenance=pd.DataFrame(prov), latest=latest, warnings=warnings)


def _derive_ebit(a: pd.DataFrame):
    """Operating income, falling back (per year) to income-statement identities when it isn't tagged."""
    ebit = a["operating_income"].copy(); how = {}
    fallbacks = [("revenue − CostsAndExpenses", a["revenue"] - a["costs_and_expenses"]),
                 ("gross profit − OperatingExpenses", a["gross_profit"] - a["operating_expenses"]),
                 ("gross profit − SG&A − R&D", a["gross_profit"] - a["sga"] - a["rnd"].fillna(0)),
                 ("pre-tax income + interest expense", a["pretax_income"] + a["interest_expense"].fillna(0))]
    for label, series in fallbacks:
        fill = ebit.isna() & series.notna()
        for fy in a.index[fill]: how[int(fy)] = label
        ebit = ebit.where(~fill, series)
    return ebit, how


def _combine_debt(a: pd.DataFrame) -> pd.Series:
    split = a["lt_debt_noncurrent"].fillna(0) + a["lt_debt_current"].fillna(0)
    has_split = a["lt_debt_noncurrent"].notna() | a["lt_debt_current"].notna()
    lt = a["lt_debt_total"].where(a["lt_debt_total"].notna(), split.where(has_split))
    # LongTermDebt usually includes the current portion; if only the split is reported, use it
    lt = lt.where(~(a["lt_debt_total"].notna() & has_split & (split > a["lt_debt_total"])), split)
    return lt.fillna(0) + a["st_borrowings"].fillna(0)


def _latest_instant(facts: dict, item: str):
    s = _series(facts, item, annual_only=False)
    if item != "shares_outstanding":
        s = s[s["form"].isin(ANNUAL_FORMS | {"10-Q", "10-Q/A"})]
    if s.empty: return None, None
    r = s.iloc[-1]
    return float(r["val"]), pd.Timestamp(r["end"]).date()


def _latest_balance_sheet(facts: dict, a: pd.DataFrame) -> dict:
    """Most recent cash, investments, debt and share count (10-Q or 10-K) for the EV-to-equity bridge."""
    out = {}
    for item in ["cash", "st_investments", "lt_debt_total", "lt_debt_noncurrent", "lt_debt_current", "st_borrowings"]:
        out[item], out[item + "_date"] = _latest_instant(facts, item)
    dates = [out["cash_date"]] + [out[k + "_date"] for k in ["lt_debt_total", "lt_debt_noncurrent"]]
    asof = max([d for d in dates if d is not None], default=None)
    # only use pieces reported at the same latest date; otherwise fall back to the last annual balance sheet
    piece = lambda k: out[k] if out[k + "_date"] == asof and out[k] is not None else np.nan
    tmp = pd.DataFrame({"lt_debt_total": [piece("lt_debt_total")], "lt_debt_noncurrent": [piece("lt_debt_noncurrent")],
                        "lt_debt_current": [piece("lt_debt_current")], "st_borrowings": [piece("st_borrowings")]})
    debt = float(_combine_debt(tmp).iloc[0])
    cash = piece("cash")
    if np.isnan(cash):
        cash, debt, asof = a["cash"].iloc[-1], a["debt"].iloc[-1], None
    sti = piece("st_investments")
    shares, shares_date = _latest_instant(facts, "shares_outstanding")
    if shares is None or (asof and shares_date and (pd.Timestamp(asof) - pd.Timestamp(shares_date)).days > 400):
        # e.g. multi-class companies stop tagging cover-page shares; use the latest diluted share count instead
        d = _series(facts, "diluted_shares", annual_only=False)
        shares, shares_date = (float(d.iloc[-1]["val"]), pd.Timestamp(d.iloc[-1]["end"]).date()) if not d.empty else (np.nan, None)
        shares_src = "WeightedAverageNumberOfDilutedSharesOutstanding (latest period)"
    else:
        shares_src = "EntityCommonStockSharesOutstanding (cover page)"
    return {"asof": asof, "cash": float(cash), "st_investments": 0.0 if np.isnan(sti) else float(sti), "debt": debt,
            "shares": shares, "shares_date": shares_date, "shares_source": shares_src}

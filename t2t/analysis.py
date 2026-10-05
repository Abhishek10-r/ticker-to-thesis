"""Ratios, peer benchmarking and industry-trend analysis."""
from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd

from .standardize import Financials

# key: (label, format, higher_is_better, category)
METRICS = {
    "revenue_growth": ("Revenue growth", "pct", True, "Growth"),
    "revenue_cagr_3y": ("Revenue CAGR (3y)", "pct", True, "Growth"),
    "gross_margin": ("Gross margin", "pct", True, "Profitability"),
    "ebit_margin": ("EBIT margin", "pct", True, "Profitability"),
    "ebitda_margin": ("EBITDA margin", "pct", True, "Profitability"),
    "net_margin": ("Net margin", "pct", True, "Profitability"),
    "roic": ("ROIC", "pct", True, "Returns"),
    "roe": ("ROE", "pct", True, "Returns"),
    "fcf_margin": ("FCF margin", "pct", True, "Cash generation"),
    "fcf_conversion": ("FCF conversion (FCF / net income)", "pct", True, "Cash generation"),
    "capex_pct": ("Capex / revenue", "pct", None, "Cash generation"),
    "net_debt_ebitda": ("Net debt / EBITDA", "x", False, "Balance sheet"),
    "interest_coverage": ("Interest coverage (EBIT / interest)", "x", True, "Balance sheet"),
    "dso": ("Receivable days", "days", False, "Efficiency"),
    "dio": ("Inventory days", "days", False, "Efficiency"),
    "dpo": ("Payable days", "days", True, "Efficiency"),
    "ccc": ("Cash conversion cycle", "days", False, "Efficiency"),
}
SCORECARD = ["revenue_growth", "revenue_cagr_3y", "gross_margin", "ebit_margin", "roic", "fcf_margin",
             "fcf_conversion", "net_debt_ebitda", "interest_coverage", "dio", "ccc"]
TREND_METRICS = ["revenue_growth", "gross_margin", "ebit_margin", "roic", "fcf_margin", "net_debt_ebitda"]
DEFAULT_TAX = 0.21


def fmt(value, kind):
    if value is None or (isinstance(value, float) and np.isnan(value)): return "n/a"
    return {"pct": f"{value:.1%}", "x": f"{value:.1f}x", "days": f"{value:.0f} days"}.get(kind, f"{value:,.2f}")


def _div(a, b):
    b = b.where(b != 0) if isinstance(b, pd.Series) else (np.nan if b == 0 else b)
    return a / b


def ratios(f: Financials) -> pd.DataFrame:
    a = f.annual
    r = pd.DataFrame(index=a.index)
    r["revenue_growth"] = a["revenue"].pct_change(fill_method=None)
    r["revenue_cagr_3y"] = (a["revenue"] / a["revenue"].shift(3)) ** (1 / 3) - 1
    for m, num in [("gross_margin", "gross_profit"), ("ebit_margin", "operating_income"), ("ebitda_margin", "ebitda"),
                   ("net_margin", "net_income"), ("fcf_margin", "fcf"), ("capex_pct", "capex")]:
        r[m] = _div(a[num], a["revenue"])
    r["da_pct"] = _div(a["d_and_a"], a["revenue"])
    r["nwc_pct"] = _div(a["nwc"], a["revenue"])
    r["fcf_conversion"] = _div(a["fcf"], a["net_income"].where(a["net_income"] > 0))
    r["tax_rate"] = _div(a["income_tax"], a["pretax_income"].where(a["pretax_income"] > 0)).clip(0, 0.5)
    tax = r["tax_rate"].fillna(DEFAULT_TAX)
    ic = a["debt"].fillna(0) + a["equity"] - a["cash"].fillna(0) - a["st_investments"].fillna(0)
    avg_ic = ((ic + ic.shift(1)) / 2).fillna(ic)
    r["roic"] = _div(a["operating_income"] * (1 - tax), avg_ic.where(avg_ic > 0))
    avg_eq = ((a["equity"] + a["equity"].shift(1)) / 2).fillna(a["equity"])
    r["roe"] = _div(a["net_income"], avg_eq.where(avg_eq > 0))
    r["net_debt_ebitda"] = _div(a["net_debt"], a["ebitda"].where(a["ebitda"] > 0))
    r["interest_coverage"] = _div(a["operating_income"], a["interest_expense"].where(a["interest_expense"] > 0))
    r["dso"] = _div(a["receivables"], a["revenue"]) * 365
    r["dio"] = _div(a["inventory"], a["cogs"]) * 365
    r["dpo"] = _div(a["payables"], a["cogs"]) * 365
    r["ccc"] = r["dso"] + r["dio"] - r["dpo"]
    return r


@dataclass
class Benchmark:
    target: str
    peers: list
    latest: pd.DataFrame        # rows = companies, cols = metrics (latest fiscal year)
    scorecard: pd.DataFrame     # target vs peer distribution, percentile, verdict
    trends: dict                # metric -> DataFrame(index=fy, cols=[target, peer_median, peer_q1, peer_q3])
    industry: dict              # summary of how the peer group itself is moving
    ratios: dict                # ticker -> ratios DataFrame


def benchmark(fins: dict[str, Financials], target: str) -> Benchmark:
    rat = {t: ratios(f) for t, f in fins.items()}
    peers = [t for t in fins if t != target]
    latest = pd.DataFrame({t: r.iloc[-1] for t, r in rat.items()}).T
    rows = []
    for m in SCORECARD:
        label, kind, hib, cat = METRICS[m]
        tv, pv = latest.loc[target, m], latest.loc[peers, m].dropna()
        if len(pv) == 0 or pd.isna(tv):
            pct, verdict = np.nan, "n/a"
        else:
            # percentile among the target + peers, oriented so 100 = best
            pool = np.append(pv.values, tv)
            pct = (pool < tv).mean() * 100 + (pool == tv).mean() * 50
            if hib is False: pct = 100 - pct
            med = pv.median(); spread = max(abs(med) * 0.10, 1e-9) if kind != "days" else 5
            better = (tv > med) if hib else (tv < med)
            verdict = "In line" if abs(tv - med) <= spread else ("Better" if better else "Worse")
        rows.append({"metric": m, "label": label, "category": cat, "format": kind, "target": tv,
                     "peer_median": pv.median() if len(pv) else np.nan,
                     "peer_q1": pv.quantile(.25) if len(pv) else np.nan, "peer_q3": pv.quantile(.75) if len(pv) else np.nan,
                     "percentile": pct, "verdict": verdict})
    scorecard = pd.DataFrame(rows).set_index("metric")

    trends = {}
    years = sorted(set().union(*[set(r.index) for r in rat.values()]))
    for m in TREND_METRICS:
        peer_tbl = pd.DataFrame({t: rat[t][m] for t in peers}).reindex(years)
        trends[m] = pd.DataFrame({target: rat[target][m].reindex(years), "peer_median": peer_tbl.median(axis=1),
                                  "peer_q1": peer_tbl.quantile(.25, axis=1), "peer_q3": peer_tbl.quantile(.75, axis=1),
                                  "peers_reporting": peer_tbl.notna().sum(axis=1)})
        trends[m] = trends[m][trends[m]["peers_reporting"] >= max(2, len(peers) // 2)]
    industry = _industry_summary(trends, target)
    return Benchmark(target, peers, latest, scorecard, trends, industry, rat)


def _change(series: pd.Series, n=3):
    s = series.dropna()
    return (s.iloc[-1] - s.iloc[-1 - n]) if len(s) > n else (s.iloc[-1] - s.iloc[0] if len(s) > 1 else np.nan)


def _industry_summary(trends, target):
    g, m = trends["revenue_growth"], trends["ebit_margin"]
    return {
        "years": f"{int(g.index[0])}–{int(g.index[-1])}" if len(g) else "",
        "peer_growth_latest": g["peer_median"].iloc[-1] if len(g) else np.nan,
        "peer_growth_avg_3y": g["peer_median"].tail(3).mean() if len(g) else np.nan,
        "target_growth_avg_3y": g[target].tail(3).mean() if len(g) else np.nan,
        "peer_margin_latest": m["peer_median"].iloc[-1] if len(m) else np.nan,
        "peer_margin_change_3y": _change(m["peer_median"]) if len(m) else np.nan,
        "target_margin_change_3y": _change(m[target]) if len(m) else np.nan,
        "peer_gm_change_3y": _change(trends["gross_margin"]["peer_median"]) if len(trends["gross_margin"]) else np.nan,
        "target_gm_change_3y": _change(trends["gross_margin"][target]) if len(trends["gross_margin"]) else np.nan,
    }

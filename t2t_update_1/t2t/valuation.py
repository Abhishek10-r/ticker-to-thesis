"""Valuation: forecast, WACC (bottom-up beta + synthetic rating), DCF, sensitivities, trading comps,
scenarios, football field, price target and rating.

The Excel model rebuilds the same maths with live formulas; tests check the two agree.
"""
from __future__ import annotations

import copy
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

from .analysis import Benchmark
from .market import regression_beta
from .standardize import Financials


# ------------------------------------------------------------------ assumptions
@dataclass
class Assumptions:
    revenue_growth: list
    ebit_margin: list
    tax_rate: float
    da_pct: float
    capex_pct: float
    nwc_pct: float
    terminal_growth: float
    wacc: float
    exit_multiple: float
    notes: dict = field(default_factory=dict)      # assumption -> rationale shown in Excel and deck

    @property
    def years(self): return len(self.revenue_growth)


def _avg(s: pd.Series, n=3, default=np.nan):
    v = s.dropna().tail(n)
    return float(v.mean()) if len(v) else default


def default_assumptions(fin: Financials, bench: Benchmark, wacc: float, exit_multiple: float, cfg: dict) -> Assumptions:
    r = bench.ratios[fin.ticker]; d = cfg["dcf"]; n = d["forecast_years"]; tg = d["terminal_growth"]
    g_last, g_cagr = r["revenue_growth"].dropna(), r["revenue_cagr_3y"].dropna()
    g0 = np.nanmean([g_last.iloc[-1] if len(g_last) else np.nan, g_cagr.iloc[-1] if len(g_cagr) else np.nan])
    g1 = float(np.clip(0.03 if np.isnan(g0) else g0, *d["growth_clip"]))
    ind_g = bench.industry.get("peer_growth_avg_3y", np.nan)
    gN = float(np.clip(0.5 * (tg if np.isnan(ind_g) else ind_g) + 0.5 * tg, tg, 0.08))
    growth = list(np.linspace(g1, gN, n))

    m0 = float(r["ebit_margin"].dropna().iloc[-1])
    own5 = _avg(r["ebit_margin"], 5, m0)
    peer_m = bench.scorecard.loc["ebit_margin", "peer_median"]
    w = d["margin_blend_peer_weight"]
    mN = float((1 - w) * own5 + w * peer_m) if not np.isnan(peer_m) else own5
    margins = [m0 + (mN - m0) * (i + 1) / n for i in range(n)]

    tax = float(np.clip(_avg(r["tax_rate"], 3, 0.21), 0.15, 0.30))
    a = Assumptions(
        revenue_growth=growth, ebit_margin=margins, tax_rate=tax,
        da_pct=_avg(r["da_pct"], 3, 0.03), capex_pct=_avg(r["capex_pct"], 3, 0.03), nwc_pct=_avg(r["nwc_pct"], 3, 0.10),
        terminal_growth=tg, wacc=wacc, exit_multiple=exit_multiple)
    a.notes = {
        "revenue_growth": f"Year 1 = average of last year's growth and the 3-year CAGR ({g1:.1%}, clipped to "
                          f"{d['growth_clip'][0]:.0%}–{d['growth_clip'][1]:.0%}); fades linearly to {gN:.1%} by year {n}, "
                          "halfway between the peer median growth rate and terminal growth.",
        "ebit_margin": f"Moves linearly from the latest margin ({m0:.1%}) to {mN:.1%} in year {n}: "
                       f"{1 - w:.0%} own 5-year average ({own5:.1%}) and {w:.0%} peer median ({peer_m:.1%}).",
        "tax_rate": "3-year average effective tax rate, bounded 15–30%.",
        "da_pct": "3-year average D&A / revenue.", "capex_pct": "3-year average capex / revenue.",
        "nwc_pct": "3-year average operating working capital (receivables + inventory − payables) / revenue.",
        "terminal_growth": "Long-run nominal growth assumption (config).",
        "wacc": "See WACC build.", "exit_multiple": "Peer median EV / EBITDA (latest fiscal year).",
    }
    return a


# ------------------------------------------------------------------ WACC
@dataclass
class WaccResult:
    wacc: float
    cost_of_equity: float
    cost_of_debt_pre: float
    cost_of_debt_post: float
    beta_relevered: float
    beta_unlevered_median: float
    beta_own_adjusted: float
    rf: float
    rf_source: str
    erp: float
    tax: float
    weight_equity: float
    weight_debt: float
    rating: str
    spread: float
    coverage: float
    beta_table: pd.DataFrame


def synthetic_rating(coverage: float, table: list) -> tuple[str, float]:
    if coverage is None or np.isnan(coverage) or coverage == np.inf: return table[0]["rating"], table[0]["spread"]
    for row in table:
        if coverage >= row["min_coverage"]: return row["rating"], row["spread"]
    return table[-1]["rating"], table[-1]["spread"]


def compute_wacc(target: str, fins: dict, market_caps: dict, prices: pd.DataFrame, rf: float, rf_source: str,
                 cfg: dict) -> WaccResult:
    m, b = cfg["market"], cfg["beta"]; t = m["marginal_tax_rate"]
    rows = []
    for tk, f in fins.items():
        if tk not in prices or tk not in market_caps: continue
        raw, r2, nobs = regression_beta(prices, tk, m["market_proxy"], b["regression_years"])
        adj = 0.67 * raw + 0.33 if b["blume_adjust"] else raw
        de = f.latest["debt"] / market_caps[tk] if market_caps[tk] else np.nan
        rows.append({"ticker": tk, "raw_beta": raw, "r2": r2, "weeks": nobs, "adjusted_beta": adj,
                     "debt_to_equity": de, "unlevered_beta": adj / (1 + (1 - t) * de), "is_target": tk == target})
    bt = pd.DataFrame(rows).set_index("ticker")
    peer_bu = bt.loc[~bt["is_target"], "unlevered_beta"].dropna()
    bu = float(peer_bu.median()) if len(peer_bu) else float(bt["unlevered_beta"].median())
    de_t = bt.loc[target, "debt_to_equity"] if target in bt.index else 0.0
    beta = bu * (1 + (1 - t) * de_t)
    ke = rf + beta * m["equity_risk_premium"]

    a = fins[target].annual
    ie = a["interest_expense"].iloc[-1]
    cov = a["operating_income"].iloc[-1] / ie if ie and ie > 0 else np.inf
    rating, spread = synthetic_rating(cov, cfg["synthetic_rating"])
    kd = rf + spread
    E, D = market_caps[target], fins[target].latest["debt"]
    we, wd = E / (E + D), D / (E + D)
    wacc = we * ke + wd * kd * (1 - t)
    floor = cfg["dcf"]["terminal_growth"] + cfg["dcf"]["min_wacc_spread_over_g"]
    return WaccResult(wacc=max(wacc, floor), cost_of_equity=ke, cost_of_debt_pre=kd, cost_of_debt_post=kd * (1 - t),
                      beta_relevered=beta, beta_unlevered_median=bu,
                      beta_own_adjusted=float(bt.loc[target, "adjusted_beta"]) if target in bt.index else np.nan,
                      rf=rf, rf_source=rf_source, erp=m["equity_risk_premium"], tax=t, weight_equity=we, weight_debt=wd,
                      rating=rating, spread=spread, coverage=cov, beta_table=bt)


# ------------------------------------------------------------------ DCF
@dataclass
class DCFResult:
    table: pd.DataFrame          # forecast years x line items
    pv_fcf: float
    tv_perpetuity: float
    pv_tv_perpetuity: float
    ev_perpetuity: float
    price_perpetuity: float
    tv_exit: float
    pv_tv_exit: float
    ev_exit: float
    price_exit: float
    implied_exit_multiple: float
    implied_growth_from_exit: float
    tv_share: float
    net_debt: float
    shares: float


def run_dcf(a: Assumptions, base_revenue: float, net_debt: float, shares: float, mid_year=True) -> DCFResult:
    n = a.years
    rev, rows, nwc_prev = base_revenue, [], base_revenue * a.nwc_pct
    for i in range(n):
        rev = rev * (1 + a.revenue_growth[i])
        ebit = rev * a.ebit_margin[i]
        nopat = ebit * (1 - a.tax_rate)
        da, capex, nwc = rev * a.da_pct, rev * a.capex_pct, rev * a.nwc_pct
        d_nwc = nwc - nwc_prev; nwc_prev = nwc
        ufcf = nopat + da - capex - d_nwc
        period = i + 0.5 if mid_year else i + 1
        df = 1 / (1 + a.wacc) ** period
        rows.append({"year": i + 1, "revenue": rev, "growth": a.revenue_growth[i], "ebit": ebit, "ebit_margin": a.ebit_margin[i],
                     "taxes": ebit * a.tax_rate, "nopat": nopat, "d_and_a": da, "capex": capex, "change_nwc": d_nwc,
                     "nwc": nwc, "ufcf": ufcf, "ebitda": ebit + da, "period": period, "discount_factor": df, "pv_ufcf": ufcf * df})
    t = pd.DataFrame(rows).set_index("year")
    pv = t["pv_ufcf"].sum()
    last = t.iloc[-1]
    df_n = 1 / (1 + a.wacc) ** n
    tv_g = last["ufcf"] * (1 + a.terminal_growth) / (a.wacc - a.terminal_growth)
    tv_x = last["ebitda"] * a.exit_multiple
    ev_g, ev_x = pv + tv_g * df_n, pv + tv_x * df_n
    return DCFResult(table=t, pv_fcf=pv, tv_perpetuity=tv_g, pv_tv_perpetuity=tv_g * df_n, ev_perpetuity=ev_g,
                     price_perpetuity=(ev_g - net_debt) / shares, tv_exit=tv_x, pv_tv_exit=tv_x * df_n, ev_exit=ev_x,
                     price_exit=(ev_x - net_debt) / shares, implied_exit_multiple=tv_g / last["ebitda"],
                     implied_growth_from_exit=(tv_x * a.wacc - last["ufcf"]) / (tv_x + last["ufcf"]),
                     tv_share=tv_g * df_n / ev_g, net_debt=net_debt, shares=shares)


def sensitivity(a: Assumptions, base_revenue, net_debt, shares, mid_year=True, wacc_step=0.005, g_step=0.0025, mult_step=1.0):
    waccs = [a.wacc + k * wacc_step for k in range(-2, 3)]
    gs = [a.terminal_growth + k * g_step for k in range(-2, 3)]
    mults = [a.exit_multiple + k * mult_step for k in range(-2, 3)]
    def price(**kw):
        b = copy.deepcopy(a)
        for k, v in kw.items(): setattr(b, k, v)
        if b.wacc <= b.terminal_growth: return np.nan
        return run_dcf(b, base_revenue, net_debt, shares, mid_year)
    g_tbl = pd.DataFrame({g: [getattr(price(wacc=w, terminal_growth=g), "price_perpetuity", np.nan) for w in waccs] for g in gs}, index=waccs)
    m_tbl = pd.DataFrame({m: [getattr(price(wacc=w, exit_multiple=m), "price_exit", np.nan) for w in waccs] for m in mults}, index=waccs)
    g_tbl.index.name = m_tbl.index.name = "wacc"
    return g_tbl, m_tbl


# ------------------------------------------------------------------ comps
def trading_comps(fins: dict, target: str, prices_now: dict, market_caps: dict, cfg: dict) -> pd.DataFrame:
    rows = []
    for tk, f in fins.items():
        if tk not in market_caps: continue
        a, L = f.annual.iloc[-1], f.latest
        ev = market_caps[tk] + L["debt"] - L["cash"] - L["st_investments"]
        ebitda, ni = a["ebitda"], a["net_income"]
        ev_ebitda = ev / ebitda if ebitda and ebitda > 0 else np.nan
        pe = market_caps[tk] / ni if ni and ni > 0 else np.nan
        if not 0 < (ev_ebitda if not np.isnan(ev_ebitda) else 1) <= cfg["comps"]["max_ev_ebitda"]: ev_ebitda = np.nan
        if not 0 < (pe if not np.isnan(pe) else 1) <= cfg["comps"]["max_pe"]: pe = np.nan
        rows.append({"ticker": tk, "name": f.name, "fiscal_year": int(f.annual.index[-1]), "price": prices_now.get(tk, np.nan),
                     "market_cap": market_caps[tk], "enterprise_value": ev, "revenue": a["revenue"], "ebitda": ebitda,
                     "net_income": ni, "ev_revenue": ev / a["revenue"], "ev_ebitda": ev_ebitda, "pe": pe,
                     "ebit_margin": a["operating_income"] / a["revenue"], "fcf_yield": a["fcf"] / market_caps[tk],
                     "is_target": tk == target})
    return pd.DataFrame(rows).set_index("ticker")


def implied_from_comps(comps: pd.DataFrame, fin: Financials, net_debt: float, shares: float) -> dict:
    peers = comps[~comps["is_target"]]
    a = fin.annual.iloc[-1]; out = {}
    for col, metric in [("ev_ebitda", "ebitda"), ("ev_revenue", "revenue"), ("pe", "net_income")]:
        v = peers[col].dropna()
        if len(v) == 0 or pd.isna(a[metric]) or a[metric] <= 0: continue
        q = {"low": v.quantile(.25), "median": v.median(), "high": v.quantile(.75)}
        if col == "pe": out[col] = {k: m * a[metric] / shares for k, m in q.items()}
        else: out[col] = {k: (m * a[metric] - net_debt) / shares for k, m in q.items()}
        out[col]["multiples"] = q
    return out


# ------------------------------------------------------------------ everything together
@dataclass
class Valuation:
    price: float
    market_cap: float
    net_debt: float
    shares: float
    assumptions: Assumptions
    wacc: WaccResult
    dcf: DCFResult
    sens_growth: pd.DataFrame
    sens_exit: pd.DataFrame
    comps: pd.DataFrame
    comps_implied: dict
    scenarios: dict
    football: pd.DataFrame
    price_target: float
    pt_components: dict
    upside: float
    rating: str
    week52: tuple


def shift_assumptions(a: Assumptions, growth=0.0, margin=0.0, wacc=0.0, terminal_growth=0.0) -> Assumptions:
    b = copy.deepcopy(a)
    b.revenue_growth = [g + growth for g in a.revenue_growth]
    b.ebit_margin = [m + margin for m in a.ebit_margin]
    b.wacc, b.terminal_growth = a.wacc + wacc, a.terminal_growth + terminal_growth
    return b


def apply_overrides(a: Assumptions, ov: dict | None) -> Assumptions:
    if not ov: return a
    b = copy.deepcopy(a)
    for k, v in ov.items():
        if v is None: continue
        if k in ("revenue_growth", "ebit_margin") and np.isscalar(v): v = [v] * b.years
        setattr(b, k, list(v) if isinstance(v, (list, tuple, np.ndarray)) else float(v))
        b.notes[k] = "User override."
    return b


def value_company(target: str, fins: dict, bench: Benchmark, prices: pd.DataFrame, rf: float, rf_source: str,
                  cfg: dict, overrides: dict | None = None, raw_prices: pd.DataFrame | None = None) -> Valuation:
    quote = prices.copy()
    if raw_prices is not None and len(raw_prices):
        for t in raw_prices.columns: quote[t] = raw_prices[t].reindex(quote.index).fillna(quote[t])
    prices_now = {t: float(quote[t].dropna().iloc[-1]) for t in quote.columns if quote[t].notna().any()}
    caps = {t: prices_now[t] * f.latest["shares"] for t, f in fins.items() if t in prices_now and f.latest["shares"] > 0}
    if target not in caps:
        raise ValueError(f"No price or share count for {target}; can't value it.")
    f = fins[target]; L = f.latest
    net_debt = L["debt"] - L["cash"] - L["st_investments"]; shares = L["shares"]

    w = compute_wacc(target, fins, caps, prices, rf, rf_source, cfg)
    comps = trading_comps(fins, target, prices_now, caps, cfg)
    peer_mult = comps.loc[~comps["is_target"], "ev_ebitda"].median()
    a = default_assumptions(f, bench, w.wacc, float(peer_mult) if not np.isnan(peer_mult) else 10.0, cfg)
    a = apply_overrides(a, overrides)
    mid = cfg["dcf"]["mid_year_convention"]
    base_rev = f.annual["revenue"].iloc[-1]
    if a.wacc - a.terminal_growth < 0.005:
        raise ValueError("WACC must exceed terminal growth.")
    d = run_dcf(a, base_rev, net_debt, shares, mid)
    sg, sx = sensitivity(a, base_rev, net_debt, shares, mid)
    implied = implied_from_comps(comps, f, net_debt, shares)

    scen = {"Base": d}
    for name, s in cfg["scenarios"].items():
        b = shift_assumptions(a, **s)
        scen[name.capitalize()] = run_dcf(b, base_rev, net_debt, shares, mid)

    wts = cfg["price_target_weights"]
    comp_val = implied.get("ev_ebitda", {}).get("median", np.nan)
    parts = {"dcf_perpetuity": d.price_perpetuity, "dcf_exit_multiple": d.price_exit, "comps_ev_ebitda": comp_val}
    used = {k: v for k, v in parts.items() if not np.isnan(v)}
    tw = sum(wts[k] for k in used)
    pt = sum(wts[k] * v for k, v in used.items()) / tw
    price = prices_now[target]; up = pt / price - 1
    r = cfg["rating"]
    rating = "BUY" if up > r["buy_above"] else ("SELL" if up < r["sell_below"] else "HOLD")

    s52 = quote[target].dropna(); s52 = s52[s52.index >= s52.index.max() - pd.DateOffset(weeks=52)]
    ff = [("52-week trading range (closing prices)", s52.min(), s52.max())]
    if "ev_ebitda" in implied: ff.append((f"Trading comps: EV/EBITDA {implied['ev_ebitda']['multiples']['low']:.1f}x–"
                                          f"{implied['ev_ebitda']['multiples']['high']:.1f}x", implied["ev_ebitda"]["low"], implied["ev_ebitda"]["high"]))
    if "pe" in implied: ff.append((f"Trading comps: P/E {implied['pe']['multiples']['low']:.1f}x–{implied['pe']['multiples']['high']:.1f}x",
                                   implied["pe"]["low"], implied["pe"]["high"]))
    inner = sg.iloc[1:4, 1:4].values
    ff.append(("DCF: perpetuity growth (WACC ±0.5pp, g ±0.25pp)", np.nanmin(inner), np.nanmax(inner)))
    inner_x = sx.iloc[1:4, 1:4].values
    ff.append(("DCF: exit multiple (WACC ±0.5pp, multiple ±1x)", np.nanmin(inner_x), np.nanmax(inner_x)))
    ff.append(("Scenarios: bear to bull (DCF)", scen["Bear"].price_perpetuity, scen["Bull"].price_perpetuity))
    football = pd.DataFrame(ff, columns=["method", "low", "high"])

    return Valuation(price=price, market_cap=caps[target], net_debt=net_debt, shares=shares, assumptions=a, wacc=w, dcf=d,
                     sens_growth=sg, sens_exit=sx, comps=comps, comps_implied=implied, scenarios=scen, football=football,
                     price_target=pt, pt_components={k: (parts[k], wts[k] / tw if k in used else 0) for k in parts},
                     upside=up, rating=rating, week52=(float(s52.min()), float(s52.max())))

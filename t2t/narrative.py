"""Rule-based commentary: turns the numbers into analyst-style sentences.

Everything here is derived from the data, so each sentence can be traced to a number in the model.
Your own view (thesis, catalysts, risks) goes in the run config under `analyst_view` and is shown
alongside, clearly marked as yours.
"""
from __future__ import annotations

import numpy as np

from .analysis import METRICS, Benchmark, fmt


def _pp(x):
    v = round(x * 100, 1)
    return "0.0pp" if v == 0 else f"{v:+.1f}pp"


def _valid(*xs): return all(x is not None and not (isinstance(x, float) and np.isnan(x)) for x in xs)


def financial_insights(b: Benchmark, name: str) -> list[str]:
    s, ind, t = b.scorecard, b.industry, b.target
    out = []
    g = s.loc["revenue_growth"]; c = s.loc["revenue_cagr_3y"]
    if _valid(g.target, g.peer_median):
        out.append(f"Growth: revenue {'grew' if g.target >= 0 else 'fell'} {abs(g.target):.1%} in the latest year vs a peer "
                   f"median of {g.peer_median:+.1%}; the 3-year CAGR is {c.target:+.1%} vs {c.peer_median:+.1%} for peers.")
    m = s.loc["ebit_margin"]
    if _valid(m.target, m.peer_median, ind["target_margin_change_3y"], ind["peer_margin_change_3y"]):
        out.append(f"Profitability: EBIT margin of {m.target:.1%} ranks at the {m.percentile:.0f}th percentile of the group "
                   f"(peer median {m.peer_median:.1%}). Over three years it moved {_pp(ind['target_margin_change_3y'])} "
                   f"while the peer median moved {_pp(ind['peer_margin_change_3y'])}.")
    r = s.loc["roic"]
    if _valid(r.target, r.peer_median):
        out.append(f"Returns: ROIC of {r.target:.1%} vs a peer median of {r.peer_median:.1%} "
                   f"({'above' if r.target > r.peer_median else 'below'} the group).")
    f = s.loc["fcf_conversion"]; fm = s.loc["fcf_margin"]
    if _valid(f.target, fm.target):
        out.append(f"Cash: FCF margin {fm.target:.1%} (peers {fm.peer_median:.1%}); FCF conversion {f.target:.0%} of net income.")
    l = s.loc["net_debt_ebitda"]
    if _valid(l.target):
        desc = lambda x: "net cash" if x < 0 else f"{x:.1f}x net debt / EBITDA"
        peer = desc(l.peer_median) if _valid(l.peer_median) else "n/a"
        out.append(f"Balance sheet: {desc(l.target)} ({l.target:.1f}x) vs a peer median of {peer} ({fmt(l.peer_median, 'x')}).")
    return out


def industry_insights(b: Benchmark) -> list[str]:
    ind = b.industry; out = []
    if _valid(ind["peer_growth_avg_3y"]):
        out.append(f"Peer-group revenue growth has averaged {ind['peer_growth_avg_3y']:+.1%} a year over the last three "
                   f"years (latest {ind['peer_growth_latest']:+.1%}); {b.target} averaged {ind['target_growth_avg_3y']:+.1%}.")
    if _valid(ind["peer_margin_change_3y"]):
        trend = "expanding" if ind["peer_margin_change_3y"] > 0.005 else ("contracting" if ind["peer_margin_change_3y"] < -0.005 else "flat")
        out.append(f"Industry EBIT margins are {trend} ({_pp(ind['peer_margin_change_3y'])} over three years, now "
                   f"{ind['peer_margin_latest']:.1%} at the median).")
    if _valid(ind["peer_gm_change_3y"], ind["target_gm_change_3y"]):
        out.append(f"Gross margin trend: peers {_pp(ind['peer_gm_change_3y'])}, {b.target} {_pp(ind['target_gm_change_3y'])} over three years.")
    return out


def strengths_weaknesses(b: Benchmark, k=3):
    sc = b.scorecard.dropna(subset=["percentile"]).sort_values("percentile", ascending=False)
    line = lambda r: f"{r.label}: {fmt(r.target, r.format)} vs peer median {fmt(r.peer_median, r.format)}"
    strengths = [line(r) for r in sc.itertuples() if r.percentile >= 60][:k]
    weaknesses = [line(r) for r in sc[::-1].itertuples() if r.percentile <= 40][:k]
    return strengths, weaknesses


def risk_flags(b: Benchmark, v) -> list[str]:
    s, ind, out = b.scorecard, b.industry, []
    if _valid(ind["target_margin_change_3y"]) and ind["target_margin_change_3y"] < -0.02:
        out.append(f"Margin compression: EBIT margin down {abs(ind['target_margin_change_3y']) * 100:.1f}pp over three years; "
                   "the forecast assumes a recovery, which is the key swing factor.")
    if _valid(s.loc["revenue_growth", "target"]) and s.loc["revenue_growth", "target"] < 0:
        out.append("Top-line decline in the latest year: the forecast relies on a return to growth.")
    nd = s.loc["net_debt_ebitda"]
    if _valid(nd.target) and (nd.target > 3 or (_valid(nd.peer_q3) and nd.target > nd.peer_q3 and nd.target > 1)):
        out.append(f"Leverage of {nd.target:.1f}x net debt / EBITDA is high relative to peers.")
    ic = s.loc["interest_coverage", "target"]
    if _valid(ic) and ic < 4: out.append(f"Thin interest cover ({ic:.1f}x EBIT / interest).")
    fc = s.loc["fcf_conversion", "target"]
    if _valid(fc) and fc < 0.7: out.append(f"Weak cash conversion: only {fc:.0%} of net income became free cash flow.")
    if v.dcf.tv_share > 0.75:
        out.append(f"Valuation leans on the terminal value ({v.dcf.tv_share:.0%} of DCF enterprise value), "
                   "so it is sensitive to WACC and long-run growth (see sensitivity tables).")
    tgt = v.comps.loc[b.target] if b.target in v.comps.index else None
    peer_med = v.comps.loc[~v.comps["is_target"], "ev_ebitda"].median()
    if tgt is not None and _valid(tgt["ev_ebitda"], peer_med) and tgt["ev_ebitda"] > peer_med * 1.15 \
            and s.loc["revenue_cagr_3y", "percentile"] < 50:
        out.append(f"Premium multiple ({tgt['ev_ebitda']:.1f}x EV/EBITDA vs peers {peer_med:.1f}x) despite below-median growth.")
    a = v.assumptions
    if a.ebit_margin[-1] - s.loc["ebit_margin", "target"] > 0.04:
        out.append(f"Forecast margin expansion of {(a.ebit_margin[-1] - s.loc['ebit_margin', 'target']) * 100:.1f}pp by year "
                   f"{a.years} is ambitious; the bear case removes 2pp.")
    return out


def valuation_summary(v, ticker) -> list[str]:
    d, w = v.dcf, v.wacc
    out = [f"Price target ${v.price_target:,.2f} vs ${v.price:,.2f} current: {v.upside:+.1%} implied return → {v.rating}.",
           f"DCF (perpetuity growth {v.assumptions.terminal_growth:.1%}, WACC {v.assumptions.wacc:.1%}): ${d.price_perpetuity:,.2f} per share; "
           f"exit-multiple method ({v.assumptions.exit_multiple:.1f}x EBITDA): ${d.price_exit:,.2f}."]
    if "ev_ebitda" in v.comps_implied:
        c = v.comps_implied["ev_ebitda"]
        out.append(f"Trading comps (peer median {c['multiples']['median']:.1f}x EV/EBITDA) imply ${c['median']:,.2f} "
                   f"(interquartile ${c['low']:,.2f}–${c['high']:,.2f}).")
    tgt = v.comps.loc[ticker]
    out.append(f"{ticker} trades at {tgt['ev_ebitda']:.1f}x EV/EBITDA and {tgt['pe']:.1f}x P/E (latest fiscal year)."
               if _valid(tgt["ev_ebitda"], tgt["pe"]) else "")
    return [o for o in out if o]


def thesis(b: Benchmark, v, name: str) -> list[str]:
    """Three data-driven thesis points: valuation gap, operating trajectory, what the price implies."""
    s, ind = b.scorecard, b.industry
    pts = []
    gap = "upside" if v.upside > 0 else "downside"
    pts.append(f"Valuation: our blended price target of ${v.price_target:,.2f} implies {abs(v.upside):.0%} {gap}; "
               f"the DCF range under ±0.5pp WACC is ${v.sens_growth.iloc[3, 2]:,.2f}–${v.sens_growth.iloc[1, 2]:,.2f}.")
    a = v.assumptions
    pts.append(f"Operating path: the base case has revenue growth moving from {a.revenue_growth[0]:+.1%} to {a.revenue_growth[-1]:+.1%} "
               f"and EBIT margin from {s.loc['ebit_margin', 'target']:.1%} to {a.ebit_margin[-1]:.1%} by year {a.years}, "
               f"against a peer median margin of {s.loc['ebit_margin', 'peer_median']:.1%}.")
    implied = (v.price - v.scenarios["Bear"].price_perpetuity) / max(v.scenarios["Bull"].price_perpetuity - v.scenarios["Bear"].price_perpetuity, 1e-9)
    pts.append(f"Market-implied expectations: the current price sits {np.clip(implied, 0, 1):.0%} of the way from our bear "
               f"(${v.scenarios['Bear'].price_perpetuity:,.2f}) to bull (${v.scenarios['Bull'].price_perpetuity:,.2f}) DCF value.")
    return pts

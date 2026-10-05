"""Data-quality check: one row per company so you can eyeball the standardised numbers before trusting them."""
import numpy as np
import pandas as pd


def quality_table(rep) -> pd.DataFrame:
    rows = []
    for t, f in rep.fins.items():
        a = f.annual.iloc[-1]; L = f.latest
        flags = []
        m = a["operating_income"] / a["revenue"] if a["revenue"] else np.nan
        if np.isnan(m) or not -0.5 < m < 0.6: flags.append("check EBIT")
        for k in ["revenue", "net_income", "cfo", "capex", "d_and_a", "cash"]:
            if pd.isna(a[k]): flags.append(f"missing {k}")
        if not L["shares"] or np.isnan(L["shares"]): flags.append("missing shares")
        if any(str(w).startswith(t + ":") for w in rep.warnings): flags.append("see notes")
        rows.append({"ticker": t, "FY": int(f.annual.index[-1]), "revenue $m": a["revenue"] / 1e6, "EBIT margin": m,
                     "net income $m": a["net_income"] / 1e6, "FCF $m": a["fcf"] / 1e6, "cash $m": L["cash"] / 1e6,
                     "debt $m": L["debt"] / 1e6, "shares m": L["shares"] / 1e6,
                     "price $": rep.val.comps.loc[t, "price"] if t in rep.val.comps.index else np.nan,
                     "flags": ", ".join(flags) or "ok"})
    return pd.DataFrame(rows).set_index("ticker")

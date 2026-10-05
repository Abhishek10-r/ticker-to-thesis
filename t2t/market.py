"""Market data: share prices (Yahoo Finance, falling back to Stooq) and the 10-year Treasury yield (FRED)."""
from __future__ import annotations

import io
import json
from pathlib import Path

import numpy as np
import pandas as pd
import requests

MARKET_PROXY = "SPY"   # S&P 500 ETF as the market portfolio for beta regressions


class MarketData:
    def prices(self, tickers: list[str], years: int = 3) -> pd.DataFrame:
        """Daily adjusted close, one column per ticker."""
        raise NotImplementedError

    def risk_free(self) -> tuple[float, str]:
        raise NotImplementedError


class LiveMarketData(MarketData):
    def __init__(self, default_rf: float = 0.042):
        self.default_rf = default_rf

    def prices(self, tickers, years=3):
        out = {}
        try:
            import yfinance as yf
            df = yf.download(tickers, period=f"{years}y", interval="1d", auto_adjust=True, progress=False,
                             group_by="column", threads=True)
            close = df["Close"] if isinstance(df.columns, pd.MultiIndex) else df[["Close"]].rename(columns={"Close": tickers[0]})
            for t in tickers:
                if t in close and close[t].notna().sum() > 50: out[t] = close[t]
        except Exception as e:                                   # network, rate limit or API change
            print(f"[market] yfinance failed ({e}); trying Stooq")
        for t in [t for t in tickers if t not in out]:
            s = _stooq(t, years)
            if s is not None: out[t] = s
        missing = [t for t in tickers if t not in out]
        if missing: print(f"[market] no price history for: {missing}")
        return pd.DataFrame(out).sort_index()

    def risk_free(self):
        try:
            r = requests.get("https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10", timeout=20)
            df = pd.read_csv(io.StringIO(r.text))
            v = pd.to_numeric(df.iloc[:, 1], errors="coerce").dropna()
            return float(v.iloc[-1]) / 100, f"FRED DGS10 (10-year Treasury), {df.iloc[v.index[-1], 0]}"
        except Exception:
            return self.default_rf, "Default assumption (FRED unavailable)"


def _stooq(ticker: str, years: int):
    try:
        url = f"https://stooq.com/q/d/l/?s={ticker.lower().replace('.', '-')}.us&i=d"
        df = pd.read_csv(url, parse_dates=["Date"]).set_index("Date")
        s = df["Close"].dropna()
        return s[s.index >= s.index.max() - pd.DateOffset(years=years)] if len(s) > 50 else None
    except Exception:
        return None


class FixtureMarketData(MarketData):
    def __init__(self, folder):
        self.folder = Path(folder)

    def prices(self, tickers, years=3):
        df = pd.read_csv(self.folder / "prices.csv", index_col=0, parse_dates=True)
        return df[[t for t in tickers if t in df.columns]]

    def risk_free(self):
        meta = json.loads((self.folder / "market.json").read_text())
        return meta["risk_free"], "Fixture value"


def regression_beta(prices: pd.DataFrame, ticker: str, market: str = MARKET_PROXY, years: int = 2) -> tuple[float, float, int]:
    """Weekly-return OLS beta vs the market over `years`. Returns (raw beta, R², observations)."""
    p = prices[[ticker, market]].dropna()
    p = p[p.index >= p.index.max() - pd.DateOffset(years=years)]
    w = p.resample("W-FRI").last().pct_change().dropna()
    if len(w) < 30: return np.nan, np.nan, len(w)
    x, y = w[market].values, w[ticker].values
    beta = np.cov(y, x, ddof=1)[0, 1] / np.var(x, ddof=1)
    r2 = np.corrcoef(x, y)[0, 1] ** 2
    return float(beta), float(r2), len(w)

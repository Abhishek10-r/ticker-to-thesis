"""SEC EDGAR client: ticker -> CIK, company metadata and XBRL company facts.

Free, no API key. The SEC requires a descriptive User-Agent (name + email) and allows ~10 requests/second.
Responses are cached on disk so repeated runs are instant and polite.
"""
from __future__ import annotations

import json
import os
import time
from pathlib import Path

import requests

TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
FACTS_URL = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json"
SUBMISSIONS_URL = "https://data.sec.gov/submissions/CIK{cik:010d}.json"
FILING_INDEX_URL = "https://www.sec.gov/Archives/edgar/data/{cik}/{accn_nodash}/{accn}-index.htm"


class SECError(RuntimeError):
    pass


class SECClient:
    """Fetches and caches EDGAR JSON. Swap in `FixtureSource` (tests) with the same three methods."""

    def __init__(self, user_agent: str | None = None, cache_dir: str | Path = ".cache/sec", max_age_hours: float = 24):
        ua = user_agent or os.environ.get("SEC_USER_AGENT")
        if not ua or "@" not in ua:
            raise SECError("SEC requires a User-Agent with your name and email, e.g. "
                           "SECClient('Jane Doe jane@example.com') or set SEC_USER_AGENT.")
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": ua, "Accept-Encoding": "gzip, deflate"})
        self.cache = Path(cache_dir); self.cache.mkdir(parents=True, exist_ok=True)
        self.max_age = max_age_hours * 3600
        self._last = 0.0
        self._tickers: dict | None = None

    # ---------- low level ----------
    def _get_json(self, url: str, cache_name: str) -> dict:
        path = self.cache / cache_name
        if path.exists() and time.time() - path.stat().st_mtime < self.max_age:
            return json.loads(path.read_text())
        for attempt in range(4):
            wait = 0.12 - (time.time() - self._last)          # stay under 10 req/s
            if wait > 0: time.sleep(wait)
            self._last = time.time()
            r = self.session.get(url, timeout=30)
            if r.status_code == 200:
                path.write_text(r.text)
                return r.json()
            if r.status_code == 404:
                raise SECError(f"Not found on EDGAR: {url}")
            time.sleep(1.5 * (attempt + 1))                     # 403/429/5xx: back off and retry
        raise SECError(f"EDGAR request failed ({r.status_code}) for {url}")

    # ---------- public ----------
    def cik(self, ticker: str) -> int:
        if self._tickers is None:
            raw = self._get_json(TICKERS_URL, "company_tickers.json")
            self._tickers = {v["ticker"].upper(): int(v["cik_str"]) for v in raw.values()}
        t = ticker.upper().replace(".", "-")
        if t not in self._tickers:
            raise SECError(f"Ticker {ticker} not found in SEC's ticker list (US-listed SEC filers only).")
        return self._tickers[t]

    def company_facts(self, cik: int) -> dict:
        return self._get_json(FACTS_URL.format(cik=cik), f"facts_{cik}.json")

    def submissions(self, cik: int) -> dict:
        return self._get_json(SUBMISSIONS_URL.format(cik=cik), f"subs_{cik}.json")


class FixtureSource:
    """Offline stand-in for SECClient, reading JSON files from a folder (used by tests and demo mode)."""

    def __init__(self, folder: str | Path):
        self.folder = Path(folder)
        self._map = json.loads((self.folder / "tickers.json").read_text())

    def cik(self, ticker: str) -> int:
        try: return int(self._map[ticker.upper()])
        except KeyError: raise SECError(f"Ticker {ticker} not in fixtures")

    def company_facts(self, cik: int) -> dict:
        return json.loads((self.folder / f"facts_{cik}.json").read_text())

    def submissions(self, cik: int) -> dict:
        return json.loads((self.folder / f"subs_{cik}.json").read_text())


def filing_url(cik: int, accn: str) -> str:
    return FILING_INDEX_URL.format(cik=int(cik), accn_nodash=accn.replace("-", ""), accn=accn)

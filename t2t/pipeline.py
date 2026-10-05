"""One call from ticker to finished research: data -> analysis -> valuation -> narrative -> Excel + deck."""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from pathlib import Path

import yaml

from . import narrative
from .analysis import Benchmark, benchmark
from .market import LiveMarketData, MarketData
from .standardize import Financials, standardize
from .valuation import Valuation, value_company

ROOT = Path(__file__).resolve().parents[1]


def load_config(path: str | Path | None = None, extra: dict | None = None) -> dict:
    cfg = yaml.safe_load(Path(path or ROOT / "config" / "default.yaml").read_text())
    for k, v in (extra or {}).items():
        if isinstance(v, dict) and isinstance(cfg.get(k), dict): cfg[k].update(v)
        else: cfg[k] = v
    return cfg


@dataclass
class Report:
    target: str
    peers: list
    fins: dict
    bench: Benchmark
    val: Valuation
    cfg: dict
    analyst: str = ""
    analyst_view: dict = field(default_factory=dict)
    run_date: str = ""
    data_note: str = ""
    warnings: list = field(default_factory=list)

    @property
    def fin(self) -> Financials: return self.fins[self.target]

    @property
    def name(self) -> str: return self.fin.name

    def insights(self):
        s, w = narrative.strengths_weaknesses(self.bench)
        return {"financial": narrative.financial_insights(self.bench, self.name),
                "industry": narrative.industry_insights(self.bench),
                "strengths": s, "weaknesses": w,
                "risks": narrative.risk_flags(self.bench, self.val),
                "valuation": narrative.valuation_summary(self.val, self.target),
                "thesis": narrative.thesis(self.bench, self.val, self.name)}


@dataclass
class Bundle:
    """Everything fetched from the outside world; analysis can be re-run on it instantly (e.g. new assumptions)."""
    target: str
    fins: dict
    prices: object
    rf: float
    rf_source: str
    warnings: list


def fetch(target: str, peers: list[str], source, market: MarketData | None = None, cfg: dict | None = None, log=print) -> Bundle:
    cfg = cfg or load_config()
    market = market or LiveMarketData(cfg["market"]["default_risk_free"])
    target = target.upper().strip(); peers = [p.upper().strip() for p in peers if p.strip() and p.upper().strip() != target]
    fins, warnings = {}, []
    for t in [target] + peers:
        try:
            log(f"[sec] {t}: fetching filings")
            fins[t] = standardize(t, source)
            warnings += fins[t].warnings
        except Exception as e:
            if t == target: raise
            warnings.append(f"Peer {t} skipped: {e}")
    log("[market] prices and risk-free rate")
    prices = market.prices(list(fins) + [cfg["market"]["market_proxy"]], years=3)
    rf, rf_src = market.risk_free()
    if target not in prices.columns: raise ValueError(f"No price history for {target}.")
    for t in [t for t in list(fins) if t not in prices.columns]:
        warnings.append(f"Peer {t} has no price data: excluded from comps, beta and benchmarking.")
        fins.pop(t)
    if len(fins) < 3: raise ValueError("Need at least two usable peers (with filings and prices).")
    return Bundle(target, fins, prices, rf, rf_src, warnings)


def analyse(bundle: Bundle, cfg: dict | None = None, overrides: dict | None = None, analyst: str = "",
            analyst_view: dict | None = None, data_note: str = "") -> Report:
    cfg = cfg or load_config()
    b = benchmark(bundle.fins, bundle.target)
    v = value_company(bundle.target, bundle.fins, b, bundle.prices, bundle.rf, bundle.rf_source, cfg, overrides)
    return Report(target=bundle.target, peers=[t for t in bundle.fins if t != bundle.target], fins=bundle.fins, bench=b, val=v,
                  cfg=cfg, analyst=analyst, analyst_view=analyst_view or {}, run_date=dt.date.today().isoformat(),
                  data_note=data_note, warnings=list(bundle.warnings))


def run(target: str, peers: list[str], source, market: MarketData | None = None, cfg: dict | None = None,
        overrides: dict | None = None, analyst: str = "", analyst_view: dict | None = None, data_note: str = "",
        log=print) -> Report:
    cfg = cfg or load_config()
    bundle = fetch(target, peers, source, market, cfg, log)
    log("[analysis] benchmarking, DCF, comps, scenarios")
    return analyse(bundle, cfg, overrides, analyst, analyst_view, data_note)


def write_outputs(rep: Report, out_dir: str | Path = "output") -> dict:
    from .deck import build_deck
    from .excel import build_workbook
    out = Path(out_dir); out.mkdir(parents=True, exist_ok=True)
    stem = f"{rep.target}_{rep.run_date}"
    xlsx = build_workbook(rep, out / f"{stem}_valuation_model.xlsx")
    pptx = build_deck(rep, out / f"{stem}_pitch_deck.pptx")
    return {"excel": xlsx, "deck": pptx}

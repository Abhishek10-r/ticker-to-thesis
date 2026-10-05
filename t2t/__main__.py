"""Command line: python -m t2t --run config/runs/nike.yaml   or   python -m t2t NKE --peers LULU DECK CROX ...

Needs a SEC User-Agent with your name and email: --ua "Jane Doe jane@example.com" or env SEC_USER_AGENT.
"""
import argparse
import os
from pathlib import Path

import yaml

from .pipeline import ROOT, load_config, run, write_outputs
from .sec import FixtureSource, SECClient


def main():
    ap = argparse.ArgumentParser(prog="t2t", description="Ticker-to-Thesis: SEC filings in, investment pitch out.")
    ap.add_argument("ticker", nargs="?", help="Target ticker, e.g. NKE")
    ap.add_argument("--peers", nargs="*", default=[], help="Peer tickers")
    ap.add_argument("--run", help="YAML run file (ticker, peers, analyst, analyst_view)")
    ap.add_argument("--ua", help="SEC User-Agent: 'Your Name your@email.com'")
    ap.add_argument("--out", default="output", help="Output folder")
    ap.add_argument("--config", help="Alternative methodology config (default: config/default.yaml)")
    ap.add_argument("--demo", action="store_true", help="Offline demo on fictional fixture companies")
    args = ap.parse_args()

    spec = yaml.safe_load(Path(args.run).read_text()) if args.run else {}
    ticker = args.ticker or spec.get("ticker"); peers = args.peers or spec.get("peers", [])
    if args.demo:
        from .market import FixtureMarketData
        fx = ROOT / "tests" / "fixtures"
        source, market, note = FixtureSource(fx), FixtureMarketData(fx), "DEMO DATA: fictional companies, for testing only."
        ticker, peers = ticker or "DEMO", peers or ["PEERA", "PEERB", "PEERC", "PEERD", "PEERE"]
    else:
        if not ticker or not peers: ap.error("give a ticker and --peers, or --run file.yaml")
        source, market, note = SECClient(args.ua or spec.get("sec_user_agent") or os.environ.get("SEC_USER_AGENT")), None, ""
    cfg = load_config(args.config, spec.get("config_overrides"))
    rep = run(ticker, peers, source, market, cfg, overrides=spec.get("overrides"), analyst=spec.get("analyst", ""),
              analyst_view={k: v for k, v in (spec.get("analyst_view") or {}).items() if v}, data_note=note)
    files = write_outputs(rep, args.out)
    v = rep.val
    print(f"\n{rep.name} ({rep.target}) · {v.rating} · price target ${v.price_target:,.2f} vs ${v.price:,.2f} ({v.upside:+.1%})")
    for w in rep.warnings: print("  note:", w)
    for k, p in files.items(): print(f"  {k}: {p}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:                      # show a clean message instead of a traceback
        raise SystemExit(f"error: {e}")

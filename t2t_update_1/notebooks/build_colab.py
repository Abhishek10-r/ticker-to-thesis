"""Builds notebooks/run_in_colab.ipynb (kept as a script so the notebook stays reviewable in git)."""
import nbformat as nbf

c = []
md = lambda s: c.append(nbf.v4.new_markdown_cell(s.strip()))
code = lambda s: c.append(nbf.v4.new_code_cell(s.strip()))

md("""
# Ticker-to-Thesis: run the engine on real data

This notebook pulls **10-K/10-Q data from SEC EDGAR** for a company and its peers, benchmarks it, values it
(DCF + comps + scenarios) and writes a **live-formula Excel model** and a **pitch deck**.

1. Set `SEC_USER_AGENT` below to your name and email (the SEC requires this; no API key needed).
2. **Runtime → Run all**. It takes about a minute.
3. The last cell downloads the Excel model and the deck.

Set `DEMO = True` to run offline on fictional companies instead.
""")
code("""
REPO = "https://github.com/Abhishek10-r/ticker-to-thesis"
SEC_USER_AGENT = "Your Name your.email@example.com"                  # <- required by the SEC: your name and email
RUN_FILE = "config/runs/nike.yaml"                                   # ticker, peers, analyst, your own view
DEMO = False

import os, sys, subprocess
if not os.path.exists("t2t") and not os.path.exists("ticker-to-thesis"):
    subprocess.run(["git", "clone", "-q", REPO], check=True)
if os.path.exists("ticker-to-thesis"): os.chdir("ticker-to-thesis")
subprocess.run(["git", "pull", "-q"])          # pick up the latest version of the code
subprocess.run([sys.executable, "-m", "pip", "install", "-q", "-r", "requirements.txt"], check=True)
sys.path.insert(0, os.getcwd())
print("ready in", os.getcwd())
""")
md("## 1. Pull filings and prices, run the analysis")
code("""
import yaml, pandas as pd
from t2t.pipeline import load_config, run, write_outputs
from t2t.sec import SECClient, FixtureSource
from t2t.market import FixtureMarketData
from t2t.quality import quality_table
pd.set_option("display.float_format", lambda v: f"{v:,.2f}")

spec = yaml.safe_load(open(RUN_FILE))
cfg = load_config(extra=spec.get("config_overrides"))
if DEMO:
    src, mkt, note = FixtureSource("tests/fixtures"), FixtureMarketData("tests/fixtures"), "DEMO DATA: fictional companies, for testing only."
    spec.update(ticker="DEMO", peers=["PEERA", "PEERB", "PEERC", "PEERD", "PEERE"])
else:
    src, mkt, note = SECClient(SEC_USER_AGENT), None, ""
rep = run(spec["ticker"], spec["peers"], src, mkt, cfg, overrides=spec.get("overrides"), analyst=spec.get("analyst", ""),
          analyst_view={k: v for k, v in (spec.get("analyst_view") or {}).items() if v}, data_note=note)
v = rep.val
print(f"\\n{rep.name} ({rep.target}): {v.rating} | price target ${v.price_target:,.2f} vs ${v.price:,.2f} ({v.upside:+.1%})")
for w in rep.warnings: print("note:", w)
""")
md("""
## 2. Data-quality check (look at this before using the output)
One row per company, latest fiscal year. Anything flagged deserves a look in the **Sources** sheet of the Excel model,
which shows the exact XBRL tag and SEC filing behind every number.
""")
code("quality_table(rep)")
md("## 3. Benchmarking and valuation")
code("""
sc = rep.bench.scorecard[["label", "target", "peer_median", "percentile", "verdict"]]
display(sc)
for k, lines in rep.insights().items():
    if k in ("financial", "industry", "valuation", "risks"):
        print(f"\\n{k.upper()}"); [print(" -", x) for x in lines]
display(v.football)
display(v.sens_growth.round(2))
""")
md("## 4. Write and download the Excel model and pitch deck")
code("""
files = write_outputs(rep, "output")
print(files)
try:
    from google.colab import files as colab_files
    for p in files.values(): colab_files.download(str(p))
except ImportError:
    pass
""")

nb = nbf.v4.new_notebook(); nb["cells"] = c
nb["metadata"] = {"kernelspec": {"name": "python3", "display_name": "Python 3"}, "colab": {"provenance": []}}
nbf.write(nb, "notebooks/run_in_colab.ipynb")
print("written")

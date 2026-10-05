"""Track a call over time: a SUMMARY.md for the latest run and a history.csv with one row per run date."""
from __future__ import annotations

import csv
from pathlib import Path


def track(rep, folder, files: dict, my_rating: str | None = None) -> dict:
    folder = Path(folder); v = rep.val
    row = {"date": rep.run_date, "price": round(v.price, 2), "price_target": round(v.price_target, 2),
           "upside": round(v.upside, 4), "model_rating": v.rating,
           "dcf_perpetuity": round(v.dcf.price_perpetuity, 2), "dcf_exit": round(v.dcf.price_exit, 2),
           "comps_ev_ebitda": round(v.comps_implied.get("ev_ebitda", {}).get("median", float("nan")), 2),
           "wacc": round(v.assumptions.wacc, 4), "risk_free": round(v.wacc.rf, 4),
           "latest_fy": int(rep.fin.annual.index[-1]), "balance_sheet_asof": str(rep.fin.latest.get("asof") or "")}

    hist = folder / "history.csv"
    rows = list(csv.DictReader(hist.open())) if hist.exists() else []
    rows = [r for r in rows if r["date"] != row["date"]] + [{k: str(x) for k, x in row.items()}]
    with hist.open("w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=list(row)); w.writeheader(); w.writerows(rows)

    first = rows[0]
    alert = ""
    if my_rating and my_rating.upper() != v.rating:
        alert = (f"> ⚠️ **The model now says {v.rating}, but the published view is {my_rating.upper()}.** "
                 "Time to revisit the thesis and assumptions in the run file.\n\n")
    lines = [f"# {rep.name} ({rep.target}): automated update", "",
             f"*Refreshed {rep.run_date} from SEC EDGAR filings and market data. Assumptions and thesis: "
             f"`config/runs/` (changed only by the analyst).*", "", alert,
             "| | Latest | First tracked |", "|---|---|---|",
             f"| Date | {row['date']} | {first['date']} |",
             f"| Share price | ${row['price']:,.2f} | ${float(first['price']):,.2f} |",
             f"| Price target | ${row['price_target']:,.2f} | ${float(first['price_target']):,.2f} |",
             f"| Upside / (downside) | {row['upside']:+.1%} | {float(first['upside']):+.1%} |",
             f"| Model rating | **{row['model_rating']}** | {first['model_rating']} |",
             f"| WACC (risk-free) | {row['wacc']:.1%} ({row['risk_free']:.2%}) | {float(first['wacc']):.1%} ({float(first['risk_free']):.2%}) |",
             f"| Latest fiscal year / balance sheet | FY{row['latest_fy']} / {row['balance_sheet_asof']} | FY{first['latest_fy']} |",
             "", f"Valuation: DCF ${row['dcf_perpetuity']:,.2f} (perpetuity) · ${row['dcf_exit']:,.2f} (exit multiple) · "
             f"comps ${row['comps_ev_ebitda']:,.2f} (EV/EBITDA median).", "",
             "Files: " + " · ".join(f"[{n}]({n})" for n in [Path(p).name for p in files.values()] + [f"{rep.target}_pitch_deck.pdf", "history.csv"]), ""]
    if rep.warnings:
        lines += ["Data notes:", ""] + [f"- {w}" for w in rep.warnings] + [""]
    lines += ["*Educational project, not investment advice.*", ""]
    (folder / "SUMMARY.md").write_text("\n".join(lines))
    return row

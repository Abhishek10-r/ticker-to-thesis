"""Ticker-to-Thesis web app: type a ticker and its peers, get the analysis, tweak assumptions, download the
Excel model and pitch deck."""
import io
import os
import tempfile
from pathlib import Path

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

from t2t.analysis import METRICS, TREND_METRICS, fmt
from t2t.market import FixtureMarketData
from t2t.pipeline import ROOT, analyse, fetch, load_config
from t2t.sec import FixtureSource, SECClient, filing_url

st.set_page_config(page_title="Ticker-to-Thesis", page_icon="📈", layout="wide")
TARGET_C, PEER_C, BAND_C, GREY = "#eb6834", "#2a78d6", "#c9d6e8", "#8a8983"
RATING_ICON = {"BUY": "🟢", "HOLD": "🟡", "SELL": "🔴"}


def md(text):  # stop Streamlit reading "$" as LaTeX
    return text.replace("$", "\\$")
CFG = load_config()


def _secret(key):
    try: return st.secrets[key]
    except Exception: return os.environ.get(key, "")


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def get_bundle(mode, ticker, peers, ua):
    if mode == "demo":
        fx = ROOT / "tests" / "fixtures"
        return fetch(ticker, list(peers), FixtureSource(fx), FixtureMarketData(fx), CFG, log=lambda *a: None)
    return fetch(ticker, list(peers), SECClient(ua, cache_dir=Path(tempfile.gettempdir()) / "t2t_sec"), None, CFG, log=lambda *a: None)


@st.cache_data(show_spinner=False)
def build_files(_rep, key):
    from t2t.deck import build_deck
    from t2t.excel import build_workbook
    with tempfile.TemporaryDirectory() as d:
        x = build_workbook(_rep, Path(d) / "model.xlsx").read_bytes()
        p = build_deck(_rep, Path(d) / "deck.pptx").read_bytes()
    return x, p


# ---------------- sidebar ----------------
with st.sidebar:
    st.title("📈 Ticker-to-Thesis")
    st.caption("SEC filings in, investment pitch out.")
    mode = st.radio("Data", ["Live: SEC EDGAR + market data", "Demo: fictional companies"], index=0)
    demo = mode.startswith("Demo")
    ticker = st.text_input("Target ticker", "DEMO" if demo else "NKE").upper().strip()
    peers_s = st.text_input("Peers (comma separated)", "PEERA, PEERB, PEERC, PEERD, PEERE" if demo else "LULU, DECK, CROX, UAA, VFC, COLM")
    ua = _secret("SEC_USER_AGENT")
    if not demo and not ua:
        ua = st.text_input("Your name and email (SEC requires it)", placeholder="Jane Doe jane@example.com")
    go = st.button("Run analysis", type="primary", use_container_width=True)
    st.divider()
    st.subheader("Tweak assumptions")
    g_shift = st.slider("Revenue growth shift (pp, every year)", -5.0, 5.0, 0.0, 0.5) / 100
    m_shift = st.slider("EBIT margin shift (pp, every year)", -5.0, 5.0, 0.0, 0.5) / 100
    tg = st.slider("Terminal growth", 0.0, 4.0, CFG["dcf"]["terminal_growth"] * 100, 0.25) / 100
    wacc_on = st.checkbox("Override WACC")
    wacc_v = st.slider("WACC", 5.0, 14.0, 9.0, 0.25, disabled=not wacc_on) / 100
    exit_on = st.checkbox("Override exit multiple")
    exit_v = st.slider("Exit EV / EBITDA", 4.0, 30.0, 12.0, 0.5, disabled=not exit_on)

peers = tuple(p.strip().upper() for p in peers_s.split(",") if p.strip())
if go:
    st.session_state["req"] = ("demo" if demo else "live", ticker, peers, ua)

if "req" not in st.session_state:
    st.markdown("## Ticker-to-Thesis\n**An automated equity research engine.** Enter a US-listed ticker and a peer group; the app pulls "
                "10-K/10-Q data from SEC EDGAR, benchmarks the company against its peers and the industry trend, values it "
                "(DCF with bottom-up WACC, trading comps, scenarios), and produces a live-formula **Excel model** and a "
                "**pitch deck**.\n\nChoose **Run analysis** in the sidebar. Try *Demo* mode for an instant offline example.")
    st.stop()

req = st.session_state["req"]
if req[0] == "live" and (not req[3] or "@" not in req[3]):
    st.error("SEC EDGAR requires a User-Agent with your name and email. Enter it in the sidebar."); st.stop()
try:
    with st.spinner(f"Pulling filings for {req[1]} and {len(req[2])} peers from SEC EDGAR..."):
        bundle = get_bundle(*req)
except Exception as e:
    st.error(f"Couldn't fetch data: {e}"); st.stop()

note = "DEMO DATA: fictional companies, for testing only." if req[0] == "demo" else ""
base = analyse(bundle, CFG, data_note=note)
a0 = base.val.assumptions
ov = {"revenue_growth": [g + g_shift for g in a0.revenue_growth], "ebit_margin": [m + m_shift for m in a0.ebit_margin],
      "terminal_growth": tg}
if wacc_on: ov["wacc"] = wacc_v
if exit_on: ov["exit_multiple"] = exit_v
changed = g_shift or m_shift or abs(tg - a0.terminal_growth) > 1e-9 or wacc_on or exit_on
try:
    rep = analyse(bundle, CFG, overrides=ov if changed else None, data_note=note)
except ValueError as e:
    st.error(str(e)); st.stop()
v, b, f = rep.val, rep.bench, rep.fin
ins = rep.insights()

# ---------------- header ----------------
st.markdown(f"## {f.name} ({rep.target})")
st.caption(f"{f.sic_description.title()} · peers: {', '.join(rep.peers)} · data as of {rep.run_date}"
           + (" · **custom assumptions**" if changed else ""))
if note: st.warning("Demo mode: fictional companies. Switch to Live for real SEC data.", icon="⚠️")
c = st.columns(5)
c[0].metric("Rating", f"{RATING_ICON[v.rating]} {v.rating}")
c[1].metric("Price target", f"${v.price_target:,.2f}")
c[2].metric("Current price", f"${v.price:,.2f}")
c[3].metric("Implied return", f"{v.upside:+.1%}")
c[4].metric("WACC", f"{v.assumptions.wacc:.1%}")

x_bytes, p_bytes = build_files(rep, (req, tuple(map(float, ov["revenue_growth"])), tuple(map(float, ov["ebit_margin"])), tg,
                                     wacc_on and wacc_v, exit_on and exit_v))
d1, d2, _ = st.columns([1.3, 1.3, 2])
d1.download_button("⬇ Excel model (live formulas)", x_bytes, f"{rep.target}_valuation_model.xlsx", use_container_width=True)
d2.download_button("⬇ Pitch deck (PowerPoint)", p_bytes, f"{rep.target}_pitch_deck.pptx", use_container_width=True)

t1, t2, t3, t4 = st.tabs(["Summary", "Benchmarking & industry", "Valuation", "Data & sources"])

with t1:
    l, r = st.columns([1, 1.15])
    with l:
        st.subheader("Thesis (model-derived)")
        for x in ins["thesis"]: st.markdown(md(f"- {x}"))
        st.subheader("Risks flagged by the data")
        for x in ins["risks"] or ["None triggered."]: st.markdown(md(f"- {x}"))
    with r:
        st.subheader("Valuation football field")
        ff = v.football.copy()
        bars = alt.Chart(ff).mark_bar(size=18, cornerRadius=3, color=PEER_C).encode(
            x=alt.X("low:Q", title="$ per share", scale=alt.Scale(zero=False)), x2="high:Q",
            y=alt.Y("method:N", sort=None, title=None, axis=alt.Axis(labelLimit=320)),
            tooltip=["method", alt.Tooltip("low:Q", format="$,.2f"), alt.Tooltip("high:Q", format="$,.2f")])
        lines = pd.DataFrame({"x": [v.price, v.price_target], "what": [f"Price ${v.price:,.2f}", f"Target ${v.price_target:,.2f}"]})
        rule = alt.Chart(lines).mark_rule(size=2, strokeDash=[4, 3]).encode(
            x="x:Q", color=alt.Color("what:N", scale=alt.Scale(range=["#52514e", TARGET_C]), legend=alt.Legend(title=None, orient="bottom")),
            tooltip=["what"])
        st.altair_chart((bars + rule).properties(height=320).configure_legend(labelLimit=200), use_container_width=True)
        sc = pd.DataFrame({k: [r.price_perpetuity, r.price_perpetuity / v.price - 1] for k, r in v.scenarios.items()},
                          index=["DCF value per share", "vs current price"])[["Bear", "Base", "Bull"]]
        st.dataframe(sc.style.format("{:,.2f}", subset=pd.IndexSlice["DCF value per share", :])
                     .format("{:+.0%}", subset=pd.IndexSlice["vs current price", :]), use_container_width=True)

with t2:
    sc = b.scorecard.copy()
    show = pd.DataFrame({"Metric": sc["label"], rep.target: [fmt(x, k) for x, k in zip(sc["target"], sc["format"])],
                         "Peer median": [fmt(x, k) for x, k in zip(sc["peer_median"], sc["format"])],
                         "Peer Q1–Q3": [f"{fmt(a, k)} – {fmt(c_, k)}" for a, c_, k in zip(sc["peer_q1"], sc["peer_q3"], sc["format"])],
                         "Percentile (100 = best)": sc["percentile"].round(0), "Verdict": sc["verdict"]})
    st.dataframe(show, hide_index=True, use_container_width=True, height=36 * (len(show) + 1) + 4,
                 column_config={"Percentile (100 = best)": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%d")})
    for x in ins["industry"]: st.markdown(md(f"- {x}"))
    m = st.selectbox("Trend", TREND_METRICS, format_func=lambda k: METRICS[k][0])
    tr = b.trends[m].reset_index().rename(columns={"index": "fy", "fy": "fy"})
    tr["fy"] = tr.iloc[:, 0].astype(int).astype(str)
    kind = METRICS[m][1]; axfmt = "%" if kind == "pct" else ".1f"
    band = alt.Chart(tr).mark_area(color=BAND_C, opacity=0.6).encode(x=alt.X("fy:N", title=None), y=alt.Y("peer_q1:Q", title=METRICS[m][0], axis=alt.Axis(format=axfmt)), y2="peer_q3:Q")
    long = tr.melt(id_vars="fy", value_vars=[rep.target, "peer_median"], var_name="series", value_name="value")
    long["series"] = long["series"].replace({"peer_median": "Peer median"})
    ln = alt.Chart(long).mark_line(point=True, strokeWidth=2).encode(
        x="fy:N", y="value:Q", color=alt.Color("series:N", scale=alt.Scale(domain=[rep.target, "Peer median"], range=[TARGET_C, PEER_C]),
                                                 legend=alt.Legend(title=None, orient="bottom")),
        tooltip=["fy", "series", alt.Tooltip("value:Q", format=".1%" if kind == "pct" else ".2f")])
    st.altair_chart((band + ln).properties(height=320), use_container_width=True)
    st.caption("Shaded band: peer interquartile range (25th–75th percentile).")

with t3:
    a = v.assumptions; d = v.dcf
    l, r = st.columns([1.4, 1])
    with l:
        st.subheader("DCF forecast ($m)")
        t_ = d.table.copy()
        tbl = pd.DataFrame({"Revenue": t_["revenue"] / 1e6, "Growth": t_["growth"], "EBIT margin": t_["ebit_margin"], "EBIT": t_["ebit"] / 1e6,
                            "NOPAT": t_["nopat"] / 1e6, "D&A": t_["d_and_a"] / 1e6, "Capex": t_["capex"] / 1e6,
                            "Δ Working capital": t_["change_nwc"] / 1e6, "Unlevered FCF": t_["ufcf"] / 1e6, "PV of UFCF": t_["pv_ufcf"] / 1e6}).T
        tbl.columns = [f"Year {i}" for i in tbl.columns]
        st.dataframe(tbl.style.format(lambda x: f"{x:.1%}" if abs(x) < 1 else f"{x:,.0f}"), use_container_width=True)
        st.subheader("Sensitivity: value per share (WACC × terminal growth)")
        sg = v.sens_growth.copy(); sg.index = [f"{x:.2%}" for x in sg.index]; sg.columns = [f"{x:.2%}" for x in sg.columns]
        st.dataframe(sg.style.format("${:,.2f}").background_gradient(cmap="RdYlGn", axis=None), use_container_width=True)
    with r:
        st.subheader("Enterprise to equity value")
        st.table(pd.DataFrame({"$m": [d.pv_fcf / 1e6, d.pv_tv_perpetuity / 1e6, d.ev_perpetuity / 1e6, -v.net_debt / 1e6,
                                      (d.ev_perpetuity - v.net_debt) / 1e6]},
                              index=["PV of forecast UFCF", "PV of terminal value", "Enterprise value", "Less net debt", "Equity value"]).style.format("{:,.0f}"))
        st.markdown(md(f"**Value per share:** ${d.price_perpetuity:,.2f} (perpetuity) · ${d.price_exit:,.2f} (exit {a.exit_multiple:.1f}x) · "
                       f"terminal value {d.tv_share:.0%} of EV"))
        w = v.wacc
        st.subheader("WACC")
        st.table(pd.DataFrame({"Value": [f"{w.rf:.2%}", f"{w.erp:.2%}", f"{w.beta_relevered:.2f}", f"{w.cost_of_equity:.2%}",
                                         f"{w.cost_of_debt_pre:.2%} ({w.rating})", f"{w.weight_equity:.0%} / {w.weight_debt:.0%}", f"{a.wacc:.2%}"]},
                              index=["Risk-free", "Equity risk premium", "Relevered beta (bottom-up)", "Cost of equity",
                                     "Pre-tax cost of debt", "Weights E / D", "WACC used"]))
    st.subheader("Trading comparables (latest fiscal year)")
    cp = v.comps.copy()
    st.dataframe(pd.DataFrame({"Company": cp["name"], "Market cap ($bn)": cp["market_cap"] / 1e9, "EV ($bn)": cp["enterprise_value"] / 1e9,
                               "EV/Sales": cp["ev_revenue"], "EV/EBITDA": cp["ev_ebitda"], "P/E": cp["pe"], "EBIT margin": cp["ebit_margin"]})
                 .style.format({"Market cap ($bn)": "{:,.1f}", "EV ($bn)": "{:,.1f}", "EV/Sales": "{:.1f}x", "EV/EBITDA": "{:.1f}x",
                                "P/E": "{:.1f}x", "EBIT margin": "{:.1%}"}, na_rep="n/a"), use_container_width=True)

with t4:
    for wmsg in rep.warnings: st.info(wmsg, icon="ℹ️")
    st.subheader(f"Standardised financials: {rep.target} ($m)")
    st.dataframe((f.annual.T / 1e6).style.format("{:,.0f}", na_rep="–"), use_container_width=True)
    st.subheader("Provenance: XBRL tag and SEC filing behind each value")
    prov = f.provenance.copy()
    prov["filing"] = [filing_url(f.cik, x) if x else "" for x in prov["accn"]]
    st.dataframe(prov, hide_index=True, use_container_width=True, column_config={"filing": st.column_config.LinkColumn("filing")})

st.caption("Educational project built on public SEC filings and market data. Not investment advice.")

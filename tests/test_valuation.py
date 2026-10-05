import numpy as np
from t2t.valuation import Assumptions, run_dcf, synthetic_rating


def test_dcf_matches_hand_calculation():
    a = Assumptions(revenue_growth=[0.1, 0.1], ebit_margin=[0.2, 0.2], tax_rate=0.25, da_pct=0.05, capex_pct=0.05,
                    nwc_pct=0.0, terminal_growth=0.02, wacc=0.10, exit_multiple=10)
    d = run_dcf(a, 100.0, net_debt=10.0, shares=10.0, mid_year=False)
    f1, f2 = 110 * 0.2 * 0.75, 121 * 0.2 * 0.75                 # D&A = capex and no NWC, so UFCF = NOPAT
    tv = f2 * 1.02 / 0.08
    ev = f1 / 1.1 + f2 / 1.1 ** 2 + tv / 1.1 ** 2
    assert np.isclose(d.ev_perpetuity, ev)
    assert np.isclose(d.price_perpetuity, (ev - 10) / 10)
    assert np.isclose(d.tv_exit, (121 * 0.2 + 121 * 0.05) * 10)


def test_mid_year_convention_raises_value():
    a = Assumptions([0.05] * 5, [0.15] * 5, 0.25, 0.03, 0.03, 0.1, 0.025, 0.09, 12)
    assert run_dcf(a, 1000, 0, 100, True).ev_perpetuity > run_dcf(a, 1000, 0, 100, False).ev_perpetuity


def test_synthetic_rating_table():
    table = [{"min_coverage": 8.5, "rating": "AAA", "spread": 0.006}, {"min_coverage": 3, "rating": "A-", "spread": 0.012},
             {"min_coverage": -1e9, "rating": "C", "spread": 0.12}]
    assert synthetic_rating(20, table)[0] == "AAA"
    assert synthetic_rating(4, table)[0] == "A-"
    assert synthetic_rating(0.1, table)[0] == "C"
    assert synthetic_rating(float("inf"), table)[0] == "AAA"


def test_report_is_internally_consistent(report):
    v = report.val
    assert v.rating in ("BUY", "HOLD", "SELL")
    assert v.scenarios["Bear"].price_perpetuity < v.scenarios["Base"].price_perpetuity < v.scenarios["Bull"].price_perpetuity
    assert np.isclose(v.sens_growth.iloc[2, 2], v.dcf.price_perpetuity)
    assert v.wacc.wacc > v.assumptions.terminal_growth
    sc = report.bench.scorecard
    assert sc["percentile"].dropna().between(0, 100).all()

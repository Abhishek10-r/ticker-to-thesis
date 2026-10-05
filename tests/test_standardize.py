import numpy as np
from conftest import FX
from t2t.sec import FixtureSource
from t2t.standardize import standardize

SRC = FixtureSource(FX)


def test_tag_switch_is_merged_across_years():
    f = standardize("DEMO", SRC)
    tags = set(f.provenance.loc[f.provenance["item"] == "revenue", "tag"])
    assert not f.annual["revenue"].isna().any()
    assert "RevenueFromContractWithCustomerExcludingAssessedTax" in tags


def test_restated_value_uses_latest_filing():
    f = standardize("DEMO", SRC)
    p = f.provenance[(f.provenance["item"] == "revenue") & (f.provenance["fy"] == 2024)].iloc[0]
    assert p["filed"] > "2025-01-01"                      # the FY2025 10-K restated FY2024


def test_missing_operating_income_is_derived_and_flagged():
    f = standardize("DEMO", SRC)
    assert f.annual["operating_income"].notna().all()
    assert any("EBIT derived" in w for w in f.warnings)
    margin = f.annual["operating_income"] / f.annual["revenue"]
    assert abs(margin.loc[2025] - 0.07) < 1e-6                 # fixture: 7% EBIT margin, tagged only as GP − SG&A


def test_gross_profit_fallback():
    f = standardize("PEERB", SRC)                           # no GrossProfit tag
    np.testing.assert_allclose(f.annual["gross_profit"], f.annual["revenue"] - f.annual["cogs"])


def test_january_year_end_labels_prior_year():
    f = standardize("PEERA", SRC)                           # FY ends 31 Jan 2026 -> FY2025
    assert f.annual.index[-1] == 2025


def test_multi_class_shares_are_summed():
    f = standardize("PEERD", SRC)
    assert 400e6 < f.latest["shares"] < 410e6


def test_share_count_falls_back_to_diluted_when_cover_page_missing():
    f = standardize("DEMO", SRC)
    assert "Diluted" in f.latest["shares_source"]


def test_latest_balance_sheet_uses_10q():
    f = standardize("DEMO", SRC)
    assert str(f.latest["asof"]) == "2025-08-31"

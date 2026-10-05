import sys
from pathlib import Path
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
FX = ROOT / "tests" / "fixtures"


@pytest.fixture(scope="session")
def report():
    if not (FX / "tickers.json").exists():
        import subprocess; subprocess.run([sys.executable, str(ROOT / "tests" / "make_fixtures.py")], check=True)
    from t2t.market import FixtureMarketData
    from t2t.pipeline import load_config, run
    from t2t.sec import FixtureSource
    return run("DEMO", ["PEERA", "PEERB", "PEERC", "PEERD", "PEERE"], FixtureSource(FX), FixtureMarketData(FX), load_config(),
               analyst="Test", data_note="DEMO", log=lambda *a: None)

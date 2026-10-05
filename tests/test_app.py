from pathlib import Path
from streamlit.testing.v1 import AppTest

APP = str(Path(__file__).resolve().parents[1] / "app.py")


def test_app_demo_mode_renders():
    at = AppTest.from_file(APP, default_timeout=120).run()
    assert not at.exception
    at.radio[0].set_value("Demo: fictional companies").run()
    at.button[0].click().run()
    assert not at.exception
    assert any("Price target" == m.label for m in at.metric)
    at.slider[1].set_value(2.0).run()                          # margin +2pp re-values the company
    assert not at.exception

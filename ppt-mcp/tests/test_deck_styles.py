"""Task 2 tests: IB deck style defaults, override precedence, ALLOWED_ROOTS."""

import json

import pytest

from config import ALLOWED_ROOTS, load_custom_overrides, resolve_style
from styles.ib_deck import IB_DECK_DEFAULTS, get_deck_style


def test_deck_defaults():
    assert get_deck_style("header_fill") == "002060"
    assert any("IBsuite" in r or "Office plugin MCP" in r for r in ALLOWED_ROOTS)


def test_ib_deck_defaults_are_spec_verbatim():
    assert IB_DECK_DEFAULTS["header_fill"] == "002060"
    assert IB_DECK_DEFAULTS["header_font_color"] == "FFFFFF"
    assert IB_DECK_DEFAULTS["header_font_bold"] is True
    assert IB_DECK_DEFAULTS["header_font_size"] == 18
    assert IB_DECK_DEFAULTS["body_font_name"] == "Arial"
    assert IB_DECK_DEFAULTS["body_font_size"] == 12
    assert IB_DECK_DEFAULTS["body_font_color"] == "000000"
    assert IB_DECK_DEFAULTS["table_grid_color"] == "D9D9D9"


def test_ib_deck_slide_size_defaults():
    assert IB_DECK_DEFAULTS["slide_width_16x9_in"] == (13.33, 7.5)
    assert IB_DECK_DEFAULTS["slide_size_16x9"] == "16:9"


def test_get_deck_style_unknown_key_raises():
    with pytest.raises(KeyError):
        get_deck_style("definitely_not_a_deck_style_key")


def test_resolve_style_override_wins():
    assert resolve_style("header_fill", {"header_fill": "123456"}) == "123456"
    assert resolve_style("header_fill", None) == "002060"


def test_load_custom_overrides_returns_dict():
    assert isinstance(load_custom_overrides(), dict)


def test_custom_json_override_wins_over_default(tmp_path, monkeypatch):
    import config

    cfg = tmp_path / "custom.json"
    cfg.write_text(json.dumps({"header_fill": "ABCDEF"}), encoding="utf-8")
    monkeypatch.setattr(config, "CUSTOM_STYLES_PATH", cfg)

    assert resolve_style("header_fill", None) == "ABCDEF"
    assert resolve_style("body_font_name", None) == "Arial"


def test_missing_custom_json_falls_back_to_defaults(tmp_path, monkeypatch):
    import config

    monkeypatch.setattr(config, "CUSTOM_STYLES_PATH", tmp_path / "absent.json")

    assert load_custom_overrides() == {}
    assert resolve_style("header_fill", None) == "002060"


def test_malformed_custom_json_falls_back_to_defaults(tmp_path, monkeypatch):
    import config

    cfg = tmp_path / "custom.json"
    cfg.write_text("{ not json", encoding="utf-8")
    monkeypatch.setattr(config, "CUSTOM_STYLES_PATH", cfg)

    assert load_custom_overrides() == {}
    assert resolve_style("header_fill", None) == "002060"

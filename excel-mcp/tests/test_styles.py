"""Task 1 tests: IB style defaults, override precedence, ALLOWED_ROOTS."""

import json

import pytest

from config import ALLOWED_ROOTS, load_custom_overrides, resolve_style
from styles.ib import IB_DEFAULTS, get_style

INT_FMT = '_-* #,##0_-;-* #,##0_-;-* "-"_-;_-@_-'


def test_ib_defaults_and_override():
    assert get_style("header_fill") == "002060"
    assert get_style("body_font") == "Arial 10"
    assert resolve_style("header_fill", {"header_fill": "123456"}) == "123456"
    assert resolve_style("header_fill", None) == "002060"


def test_allowed_roots_has_project_default():
    assert any("Office plugin MCP" in r for r in ALLOWED_ROOTS)


def test_ib_defaults_are_spec_verbatim():
    assert IB_DEFAULTS["header_fill"] == "002060"
    assert IB_DEFAULTS["input_fill"] == "D9E1F2"
    assert IB_DEFAULTS["number_format_int"] == INT_FMT
    assert IB_DEFAULTS["number_format_pct"] == "%_0.0%"
    assert IB_DEFAULTS["border_grid"] == "D9D9D9"


def test_get_style_unknown_key_raises():
    with pytest.raises(KeyError):
        get_style("definitely_not_a_style_key")


def test_load_custom_overrides_returns_dict():
    assert isinstance(load_custom_overrides(), dict)


def test_custom_json_override_wins_over_default(tmp_path, monkeypatch):
    import config

    cfg = tmp_path / "custom.json"
    cfg.write_text(json.dumps({"header_fill": "ABCDEF"}), encoding="utf-8")
    monkeypatch.setattr(config, "CUSTOM_STYLES_PATH", cfg)

    assert resolve_style("header_fill", None) == "ABCDEF"
    assert resolve_style("body_font", None) == "Arial 10"


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

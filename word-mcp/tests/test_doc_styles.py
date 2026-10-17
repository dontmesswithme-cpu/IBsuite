"""Task 1 tests: IB doc style defaults, override precedence, DENIED_PREFIXES."""

import json
import pytest

from config import DENIED_PREFIXES, load_custom_overrides, resolve_style
from styles.ib_doc import IB_DOC_DEFAULTS, get_doc_style


def test_doc_defaults():
    assert get_doc_style("heading_fill") == "002060"
    assert get_doc_style("body_font") == "Arial 10"
    lowered = [d.lower() for d in DENIED_PREFIXES]
    assert any("windows" in d for d in lowered)


def test_ib_doc_defaults_are_spec_verbatim():
    assert IB_DOC_DEFAULTS["heading_fill"] == "002060"
    assert IB_DOC_DEFAULTS["heading_font_color"] == "FFFFFF"
    assert IB_DOC_DEFAULTS["heading_font_bold"] is True
    assert IB_DOC_DEFAULTS["heading_font_name"] == "Arial"
    assert IB_DOC_DEFAULTS["heading_font_size"] == 14
    assert IB_DOC_DEFAULTS["body_font"] == "Arial 10"
    assert IB_DOC_DEFAULTS["body_font_name"] == "Arial"
    assert IB_DOC_DEFAULTS["body_font_size"] == 10
    assert IB_DOC_DEFAULTS["body_font_color"] == "000000"
    assert IB_DOC_DEFAULTS["table_grid_color"] == "D9D9D9"
    assert IB_DOC_DEFAULTS["table_header_fill"] == "002060"
    assert IB_DOC_DEFAULTS["table_header_color"] == "FFFFFF"
    assert IB_DOC_DEFAULTS["page_size"] == "Letter"


def test_get_doc_style_unknown_key_raises():
    with pytest.raises(KeyError):
        get_doc_style("definitely_not_a_doc_style_key")


def test_resolve_style_override_wins():
    assert resolve_style("heading_fill", {"heading_fill": "123456"}) == "123456"
    assert resolve_style("heading_fill", None) == "002060"


def test_load_custom_overrides_returns_dict():
    assert isinstance(load_custom_overrides(), dict)


def test_custom_json_override_wins_over_default(tmp_path, monkeypatch):
    import config

    cfg = tmp_path / "custom.json"
    cfg.write_text(json.dumps({"heading_fill": "ABCDEF"}), encoding="utf-8")
    monkeypatch.setattr(config, "CUSTOM_STYLES_PATH", cfg)

    assert resolve_style("heading_fill", None) == "ABCDEF"
    assert resolve_style("body_font_name", None) == "Arial"


def test_missing_custom_json_falls_back_to_defaults(tmp_path, monkeypatch):
    import config

    monkeypatch.setattr(config, "CUSTOM_STYLES_PATH", tmp_path / "absent.json")

    assert load_custom_overrides() == {}
    assert resolve_style("heading_fill", None) == "002060"


def test_malformed_custom_json_falls_back_to_defaults(tmp_path, monkeypatch):
    import config

    cfg = tmp_path / "custom.json"
    cfg.write_text("{ not json", encoding="utf-8")
    monkeypatch.setattr(config, "CUSTOM_STYLES_PATH", cfg)

    assert load_custom_overrides() == {}
    assert resolve_style("heading_fill", None) == "002060"

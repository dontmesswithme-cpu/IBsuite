"""Task 7 tests: engine dispatcher (headless default, headed fallback)."""

import pytest
from openpyxl import Workbook

import core.engine as engine_module
from core.engine import dispatch
from core.headed import HeadedUnavailable


def make_minimal_xlsx(path: str) -> None:
    wb = Workbook()
    wb.active.title = "Sheet1"
    wb.save(path)
    wb.close()


def test_dispatch_defaults_headless(tmp_path):
    f = tmp_path / "m.xlsx"
    make_minimal_xlsx(str(f))
    out = dispatch("read_range", {"path": str(f), "sheet": "Sheet1", "cell_range": "A1"})
    assert "values" in out


def test_dispatch_unknown_mode_rejected(tmp_path):
    f = tmp_path / "m.xlsx"
    make_minimal_xlsx(str(f))

    with pytest.raises(ValueError, match="mode"):
        dispatch(
            "read_range",
            {"path": str(f), "sheet": "Sheet1", "cell_range": "A1", "mode": "hyperdrive"},
        )


def test_dispatch_unknown_tool_rejected(tmp_path):
    f = tmp_path / "m.xlsx"
    make_minimal_xlsx(str(f))

    with pytest.raises(ValueError, match="unknown tool"):
        dispatch("teleport", {"path": str(f)})


def test_dispatch_write_then_read(tmp_path):
    f = tmp_path / "m.xlsx"
    make_minimal_xlsx(str(f))

    wrote = dispatch(
        "write_range",
        {"path": str(f), "sheet": "Sheet1", "cell_range": "A1", "values": [[7]]},
    )
    assert wrote["ok"] is True
    assert dispatch(
        "read_range", {"path": str(f), "sheet": "Sheet1", "cell_range": "A1"}
    )["values"] == [[7]]


def test_dispatch_format_chart_and_versions(tmp_path):
    f = tmp_path / "m.xlsx"
    make_minimal_xlsx(str(f))

    assert dispatch(
        "apply_ib_table",
        {"path": str(f), "sheet": "Sheet1", "cell_range": "A1:B2", "has_totals": False},
    )["ok"] is True
    assert dispatch(
        "add_sheet", {"path": str(f), "name": "S2"}
    )["ok"] is True
    listed = dispatch("list_versions", {"path": str(f)})
    assert isinstance(listed, dict) and isinstance(listed["versions"], list)


def test_dispatch_headed_falls_back_with_warning(monkeypatch, tmp_path):
    f = tmp_path / "m.xlsx"
    make_minimal_xlsx(str(f))

    def boom(*args, **kwargs):
        raise HeadedUnavailable("simulated detached COM")

    monkeypatch.setattr(engine_module.headed, "headed_read", boom)

    out = dispatch(
        "read_range",
        {"path": str(f), "sheet": "Sheet1", "cell_range": "A1", "mode": "headed"},
    )
    assert out["warning"] == "headed_unavailable_fell_back"
    assert "values" in out


def test_dispatch_macro_is_server_only(tmp_path):
    with pytest.raises(ValueError, match="server"):
        dispatch("macro_run_gated", {"path": "x", "macro": "y"})


def test_dispatch_missing_param_rejected(tmp_path):
    f = tmp_path / "m.xlsx"
    make_minimal_xlsx(str(f))

    with pytest.raises(ValueError, match="read_range"):
        dispatch("read_range", {"path": str(f), "sheet": "Sheet1"})

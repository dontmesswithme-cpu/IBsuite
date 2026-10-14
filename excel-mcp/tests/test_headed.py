"""Task 7 tests: headed COM engine (graceful when Excel is absent)."""

import pytest
from openpyxl import Workbook

from core.engine import dispatch
from core.headed import HeadedUnavailable, headed_apply, headed_read, is_excel_available


def make_minimal_xlsx(path: str) -> None:
    wb = Workbook()
    wb.active.title = "Sheet1"
    wb.save(path)
    wb.close()


@pytest.mark.needs_excel
def test_headed_detached_falls_back(tmp_path):
    f = tmp_path / "m.xlsx"
    make_minimal_xlsx(str(f))
    out = dispatch(
        "read_range",
        {"path": str(f), "sheet": "Sheet1", "cell_range": "A1", "mode": "headed"},
    )
    assert out.get("values") is not None or out.get("warning") == "headed_unavailable_fell_back"


@pytest.mark.needs_excel
def test_is_excel_available_returns_bool():
    assert isinstance(is_excel_available(), bool)


@pytest.mark.needs_excel
def test_headed_read_write_roundtrip(tmp_path):
    if not is_excel_available():
        pytest.skip("Excel Desktop COM unavailable on this box")
    f = tmp_path / "m.xlsx"
    make_minimal_xlsx(str(f))

    headed_apply(
        str(f),
        {"tool": "write_range", "sheet": "Sheet1", "cell_range": "A1", "values": [[42]]},
    )
    out = headed_read(str(f), "Sheet1", "A1")
    assert out["values"] == [[42]]


def test_headed_read_bad_range_rejected_before_com(tmp_path):
    f = tmp_path / "m.xlsx"
    make_minimal_xlsx(str(f))

    with pytest.raises(ValueError, match="bad .*range"):
        headed_read(str(f), "Sheet1", "nope")


def test_headed_apply_unsupported_tool_signals_fallback(tmp_path):
    f = tmp_path / "m.xlsx"
    make_minimal_xlsx(str(f))

    with pytest.raises(HeadedUnavailable):
        headed_apply(
            str(f),
            {"tool": "format_range", "sheet": "Sheet1", "cell_range": "A1", "style": {}},
        )


def test_headed_read_missing_file_before_com(tmp_path):
    f = tmp_path / "ghost.xlsx"

    with pytest.raises(FileNotFoundError):
        headed_read(str(f), "Sheet1", "A1")

    assert not f.exists()

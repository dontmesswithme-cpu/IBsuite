"""Tests for Real-Model Golden fixture, formulas verbatim, regression pins, and shots (R1, R2, R4, R5)."""

import base64
import hashlib
import json
from pathlib import Path

import pytest
from openpyxl import load_workbook

from core.advanced_headless import add_chart, add_image, add_pivot
from core.format_headless import apply_ib_table
from core.headless import read_range, write_range
from core.shots import full_shot, thumbnail

FIXTURE_PATH = Path(__file__).parent / "fixtures" / "ib_model.xlsx"
GOLDEN_PATH = Path(__file__).parent / "golden_ib.json"
GOLDEN = json.loads(GOLDEN_PATH.read_text(encoding="utf-8"))


def _rgb(color) -> str | None:
    if color is None or color.rgb is None:
        return None
    rgb = str(color.rgb)
    return rgb[2:] if len(rgb) == 8 else rgb


def test_full_ib_grade_vs_golden(tmp_path):
    assert FIXTURE_PATH.is_file()
    model_copy = tmp_path / "model.xlsx"
    model_copy.write_bytes(FIXTURE_PATH.read_bytes())

    wb = load_workbook(str(model_copy))
    try:
        assert set(wb.sheetnames) >= {"IS", "BS", "CF", "Comps"}

        # Assert IS formatting matches golden
        ws_is = wb["IS"]

        # Header fill and font
        assert _rgb(ws_is["A1"].fill.start_color) == GOLDEN["fills"]["header"]
        assert _rgb(ws_is["B1"].fill.start_color) == GOLDEN["fills"]["header"]
        assert ws_is["A1"].font.bold is GOLDEN["fonts"]["header"]["bold"]
        assert ws_is["A1"].font.size == GOLDEN["fonts"]["header"]["size"]
        assert _rgb(ws_is["A1"].font.color) == GOLDEN["fonts"]["header"]["color"]
        assert ws_is["A1"].font.name == GOLDEN["fonts"]["header"]["name"]

        # Body font and number formats
        assert ws_is["A2"].font.name == GOLDEN["fonts"]["body"]["name"]
        assert ws_is["A2"].font.size == GOLDEN["fonts"]["body"]["size"]
        assert ws_is["B2"].number_format == GOLDEN["number_formats"]["int"]

        # Input cells fill
        assert _rgb(ws_is["B2"].fill.start_color) == GOLDEN["fills"]["input"]

        # Borders: grid and total
        assert ws_is["A2"].border.bottom.style == GOLDEN["borders"]["grid"]["style"]
        assert _rgb(ws_is["A2"].border.bottom.color) == GOLDEN["borders"]["grid"]["color"]
        assert ws_is["A12"].border.bottom.style == GOLDEN["borders"]["total"]["style"]
        assert _rgb(ws_is["A12"].border.bottom.color) == GOLDEN["borders"]["total"]["color"]

        # Layout: freeze panes, col width, row height, print setup
        assert ws_is.freeze_panes == GOLDEN["layout"]["freeze_panes"]
        assert 14 <= ws_is.column_dimensions["A"].width <= 18
        assert ws_is.row_dimensions[1].height == GOLDEN["layout"]["row_height"]
        assert ws_is.page_setup.orientation == GOLDEN["layout"]["print"]["orientation"]
        assert ws_is.page_setup.fitToWidth == GOLDEN["layout"]["print"]["fit_to_pages_wide"]
        assert ws_is.page_setup.fitToHeight == GOLDEN["layout"]["print"]["fit_to_pages_tall"]
        assert ws_is.oddFooter.center.text == GOLDEN["layout"]["print"]["footer"]
        assert ws_is.oddFooter.left.text == ""
        assert ws_is.oddFooter.right.text == ""
        assert ws_is.oddFooter.center.text == "&F | &A"
        assert f"{ws_is.oddFooter.left.text}{ws_is.oddFooter.center.text}{ws_is.oddFooter.right.text}" == "&F | &A"

        # Chart present on Comps
        assert len(wb["Comps"]._charts) == 1
    finally:
        wb.close()

    # Verify apply_ib_table on a range in the model yields golden properties
    apply_res = apply_ib_table(str(model_copy), "IS", "A1:D12", has_totals=True)
    assert apply_res["ok"] is True


def test_formulas_verbatim_and_circular_safe(tmp_path):
    assert FIXTURE_PATH.is_file()
    model_copy = tmp_path / "model.xlsx"
    model_copy.write_bytes(FIXTURE_PATH.read_bytes())

    # Read range without evaluation - verify verbatim formula strings
    res_is = read_range(str(model_copy), "IS", "B4")
    assert res_is["values"][0][0] == "=B2-B3"

    # Cross-sheet links
    res_bs_cash = read_range(str(model_copy), "BS", "B2")
    assert res_bs_cash["values"][0][0] == "=CF!B10"

    res_cf_ni = read_range(str(model_copy), "CF", "B2")
    assert res_cf_ni["values"][0][0] == "=IS!B12"

    res_bs_re = read_range(str(model_copy), "BS", "C12")
    assert res_bs_re["values"][0][0] == "=B12+IS!C12"

    # SUMIFS pivot-summary
    res_comps = read_range(str(model_copy), "Comps", "B11")
    assert "SUMIFS" in res_comps["values"][0][0]

    # Intentional circular cell flagged CIRCULAR_INTENTIONAL
    res_flag = read_range(str(model_copy), "CF", "W1")
    assert res_flag["values"][0][0] == "CIRCULAR_INTENTIONAL"

    res_circ = read_range(str(model_copy), "CF", "X1")
    assert res_circ["values"][0][0] == "=CF!X1+1"

    # Loading with openpyxl does not crash or hang
    wb = load_workbook(str(model_copy), data_only=False)
    try:
        assert wb["CF"]["X1"].value == "=CF!X1+1"
    finally:
        wb.close()

    # Write and roundtrip custom formula verbatim
    write_range(str(model_copy), "IS", "D15", [["=IS!B2-BS!B2"]])
    verify_write = read_range(str(model_copy), "IS", "D15")
    assert verify_write["values"][0][0] == "=IS!B2-BS!B2"


def test_regression_pins(tmp_path):
    assert FIXTURE_PATH.is_file()
    model_copy = tmp_path / "model.xlsx"
    model_copy.write_bytes(FIXTURE_PATH.read_bytes())
    orig_sha = hashlib.sha256(model_copy.read_bytes()).hexdigest()

    # 1. Single-range headless error reads "bad cell range"
    with pytest.raises(ValueError, match="bad cell range"):
        read_range(str(model_copy), "IS", "BAD_RANGE_123")
    with pytest.raises(ValueError, match="bad cell range"):
        write_range(str(model_copy), "IS", "BAD_RANGE_123", [[1]])

    # 2. Multi-range advanced errors read "bad data range" / "bad source range" / "bad image anchor"
    with pytest.raises(ValueError, match="bad data range"):
        add_chart(str(model_copy), "IS", "BAD_DATA", "bar", "E2")
    with pytest.raises(ValueError, match="bad source range"):
        add_pivot(str(model_copy), "IS", "BAD_SRC", "NewPivot", ["Line Item"], ["2023A"])
    with pytest.raises(ValueError, match="bad image anchor"):
        add_image(str(model_copy), "IS", str(FIXTURE_PATH), "BAD_ANCHOR")

    # 3. add_pivot REFUSES an existing dest sheet
    with pytest.raises(ValueError, match="already exists"):
        add_pivot(str(model_copy), "IS", "A1:D12", "IS", ["Line Item"], ["2023A"])

    # 4. Oversize > 100000 cells raises "split into batches" with file bytes untouched
    oversize_rows = [[1] * 1001 for _ in range(101)]  # 101 * 1001 = 101,101 cells
    with pytest.raises(ValueError, match="split into batches"):
        write_range(str(model_copy), "IS", "A1", oversize_rows)
    assert hashlib.sha256(model_copy.read_bytes()).hexdigest() == orig_sha

    # 5. golden_ib.json top-level keys frozen to exactly fills, number_formats, fonts, borders, layout
    assert set(GOLDEN.keys()) == {"fills", "number_formats", "fonts", "borders", "layout"}


def test_shots_on_real_model(tmp_path):
    assert FIXTURE_PATH.is_file()
    model_copy = tmp_path / "model.xlsx"
    model_copy.write_bytes(FIXTURE_PATH.read_bytes())

    # Range thumbnail on real model
    thumb = thumbnail(str(model_copy), "IS", "A1:D12")
    assert isinstance(thumb, str)
    raw = base64.b64decode(thumb)
    assert raw.startswith(b"\x89PNG\r\n\x1a\n")
    assert len(thumb) < 40000

    # Full shot on same range is larger than thumbnail
    full = full_shot(str(model_copy), "IS", "A1:D12", mode="headless")
    assert isinstance(full, str)
    assert len(thumb) < len(full)

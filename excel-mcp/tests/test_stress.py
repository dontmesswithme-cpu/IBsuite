"""Task P1.2 tests: headed COM stress, retry/needs_retry handling, and detached fallback (R3)."""

import hashlib
from pathlib import Path

import pytest

import core.engine as engine_module
from core.engine import dispatch
from core.headed import HeadedUnavailable, headed_apply, headed_read, is_excel_available
from core.versions import backup

FIXTURE_PATH = Path(__file__).parent / "fixtures" / "ib_model.xlsx"


@pytest.mark.needs_excel
def test_headed_realmodel_roundtrip(tmp_path):
    if not is_excel_available():
        pytest.skip("Excel Desktop COM unavailable on this box")

    assert FIXTURE_PATH.is_file()
    f = tmp_path / "model.xlsx"
    f.write_bytes(FIXTURE_PATH.read_bytes())

    res_apply = headed_apply(
        str(f),
        {
            "tool": "write_range",
            "sheet": "IS",
            "cell_range": "A15:D15",
            "values": [["Advisory", 100, 200, "=B15+C15"]],
        },
    )
    assert res_apply["ok"] is True

    res_read = headed_read(str(f), "IS", "A15:D15")
    assert res_read["values"][0][0] == "Advisory"
    assert res_read["values"][0][1] == 100
    assert res_read["values"][0][2] == 200
    assert res_read["formulas"][0][3] == "=B15+C15"


def test_busy_locked_needs_retry_no_halfwrite(monkeypatch, tmp_path):
    assert FIXTURE_PATH.is_file()
    f = tmp_path / "model.xlsx"
    f.write_bytes(FIXTURE_PATH.read_bytes())
    orig_sha = hashlib.sha256(f.read_bytes()).hexdigest()

    # Create backup before write attempt
    bid = backup(str(f), "write_range", "A15:D15")
    assert bid

    # Track delays slept during retry backoff
    slept = []
    monkeypatch.setattr(
        "core.headed.time.sleep",
        lambda d: slept.append(d),
    )

    # Simulate Excel busy/locked on each open/write
    attempts = [0]

    def _simulated_locked_excel():
        attempts[0] += 1
        # Corrupt file temporarily to simulate a partial/half-write failure mid-operation
        f.write_bytes(b"HALF_WRITTEN_CORRUPTED_DATA")
        raise PermissionError("Permission denied: file is locked by Excel process")

    monkeypatch.setattr(
        "core.headed._excel",
        _simulated_locked_excel,
    )

    result = headed_apply(
        str(f),
        {
            "tool": "write_range",
            "sheet": "IS",
            "cell_range": "A15:D15",
            "values": [["Clobber", 1, 2, 3]],
            "backup_id": bid,
        },
    )

    # 1. Output signals needs_retry
    assert result["ok"] is False
    assert result["status"] == "needs_retry"
    assert "locked" in result["reason"].lower() or "busy" in result["reason"].lower()

    # 2. Retries with backoff 200/500/1000ms took place (3 retries after initial)
    assert slept == [0.2, 0.5, 1.0]

    # 3. .bak restored, file is untouched/not half-written
    assert hashlib.sha256(f.read_bytes()).hexdigest() == orig_sha


def test_detached_falls_back(monkeypatch, tmp_path):
    assert FIXTURE_PATH.is_file()
    f = tmp_path / "model.xlsx"
    f.write_bytes(FIXTURE_PATH.read_bytes())

    def boom(*args, **kwargs):
        raise HeadedUnavailable("simulated detached COM session")

    monkeypatch.setattr(engine_module.headed, "headed_read", boom)

    out = dispatch(
        "read_range",
        {"path": str(f), "sheet": "IS", "cell_range": "A1:B2", "mode": "headed"},
    )

    # Falls back gracefully to headless with fallback warning, never crashes
    assert out.get("warning") == "headed_unavailable_fell_back"
    assert "values" in out
    assert out["values"][0][0] == "Line Item"

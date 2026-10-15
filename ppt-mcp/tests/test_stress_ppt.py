import hashlib
from pathlib import Path
import pytest

import core.engine as engine_module
from core.engine import dispatch
from core.headed import HeadedUnavailable, headed_apply, headed_read, is_powerpoint_available
from core.versions import backup

FIXTURE_PATH = Path(__file__).parent / "fixtures" / "ib_deck.pptx"


@pytest.mark.needs_powerpoint
def test_headed_deck_roundtrip(tmp_path):
    if not is_powerpoint_available():
        pytest.skip("PowerPoint Desktop COM unavailable on this box")

    assert FIXTURE_PATH.is_file()
    f = tmp_path / "deck.pptx"
    f.write_bytes(FIXTURE_PATH.read_bytes())

    res_read = headed_read(str(f), slide=1)
    assert res_read["ok"] is True
    assert "Pitch Book" in res_read["title"]

    res_apply = headed_apply(
        str(f),
        {"tool": "set_text", "slide": 1, "shape": "title", "text": "Headed Roundtrip Title"},
    )
    assert res_apply["ok"] is True

    res_read_after = headed_read(str(f), slide=1)
    assert res_read_after["title"] == "Headed Roundtrip Title"


def test_ppt_busy_locked_needs_retry_no_halfwrite(monkeypatch, tmp_path):
    assert FIXTURE_PATH.is_file()
    f = tmp_path / "deck.pptx"
    f.write_bytes(FIXTURE_PATH.read_bytes())
    orig_sha = hashlib.sha256(f.read_bytes()).hexdigest()

    bid = backup(str(f), "set_text", "title")
    assert bid

    slept = []
    monkeypatch.setattr(
        "core.headed.time.sleep",
        lambda d: slept.append(d),
    )

    def _simulated_locked_ppt():
        f.write_bytes(b"CORRUPTED_HALF_WRITE_PPT_DATA")
        raise PermissionError("Permission denied: file is locked by PowerPoint process")

    monkeypatch.setattr(
        "core.headed._powerpoint",
        _simulated_locked_ppt,
    )

    result = headed_apply(
        str(f),
        {
            "tool": "set_text",
            "slide": 1,
            "shape": "title",
            "text": "Clobber Text",
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


def test_ppt_detached_falls_back(monkeypatch, tmp_path):
    assert FIXTURE_PATH.is_file()
    f = tmp_path / "deck.pptx"
    f.write_bytes(FIXTURE_PATH.read_bytes())

    def boom(*args, **kwargs):
        raise HeadedUnavailable("simulated detached PowerPoint COM session")

    monkeypatch.setattr(engine_module.headed, "headed_read", boom)

    out = dispatch(
        "read_text",
        {"path": str(f), "slide": 1, "mode": "headed"},
    )

    assert out.get("warning") == "headed_unavailable_fell_back"
    assert "Pitch Book" in out["title"]

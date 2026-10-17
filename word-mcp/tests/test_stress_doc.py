import hashlib
from pathlib import Path
import pytest
from docx import Document
from core.engine import dispatch
from core.headed import HeadedUnavailable, headed_apply, headed_read, is_word_available
from core.versions import backup

FIXTURE_PATH = Path(__file__).parent / "fixtures" / "ib_memo.docx"


@pytest.mark.needs_word
def test_headed_memo_roundtrip(tmp_path):
    if not is_word_available():
        pytest.skip("Word Desktop COM unavailable")
    assert FIXTURE_PATH.is_file()
    f = tmp_path / "memo.docx"
    f.write_bytes(FIXTURE_PATH.read_bytes())
    out = dispatch("read_text", {"path": str(f), "start": 0, "end": 0, "mode": "headed"})
    assert "paragraphs" in out or out.get("warning") == "headed_unavailable_fell_back"


def test_word_busy_locked_needs_retry_no_halfwrite(monkeypatch, tmp_path):
    assert FIXTURE_PATH.is_file()
    f = tmp_path / "memo.docx"
    f.write_bytes(FIXTURE_PATH.read_bytes())
    orig_sha = hashlib.sha256(f.read_bytes()).hexdigest()
    bid = backup(str(f), "add_paragraph", "p")
    assert bid
    slept = []
    monkeypatch.setattr("core.headed.time.sleep", lambda d: slept.append(d))

    def _simulated_locked():
        f.write_bytes(b"CORRUPTED")
        raise PermissionError("file is locked by Word process")

    monkeypatch.setattr("core.headed._word", _simulated_locked)
    result = headed_apply(str(f), {"tool": "add_paragraph", "text": "X", "backup_id": bid})
    assert result["ok"] is False
    assert result["status"] == "needs_retry"
    assert slept == [0.2, 0.5, 1.0]
    assert hashlib.sha256(f.read_bytes()).hexdigest() == orig_sha


def test_word_detached_falls_back(monkeypatch, tmp_path):
    assert FIXTURE_PATH.is_file()
    f = tmp_path / "memo.docx"
    f.write_bytes(FIXTURE_PATH.read_bytes())

    def boom(*args, **kwargs):
        raise HeadedUnavailable("simulated detached")

    import core.engine as engine_module
    monkeypatch.setattr(engine_module.headed, "headed_read", boom)
    out = dispatch("read_text", {"path": str(f), "start": 0, "end": 0, "mode": "headed"})
    assert out.get("warning") == "headed_unavailable_fell_back"

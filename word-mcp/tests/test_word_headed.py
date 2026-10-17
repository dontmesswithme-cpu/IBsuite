import pytest
from docx import Document
from core.engine import dispatch
from core.headed import HeadedUnavailable, _word, is_word_available


def make_minimal_docx(path: str) -> None:
    doc = Document()
    doc.add_paragraph("Sample Title")
    doc.save(path)


def test_headed_visible_default(monkeypatch):
    import core.headed as headed_mod

    seen = {}

    class FakeApp:
        def __init__(self, seen):
            self._seen = seen

        def __setattr__(self, name, value):
            if name == "Visible":
                self._seen[name] = value
            else:
                object.__setattr__(self, name, value)

        def Quit(self):
            pass

    monkeypatch.setattr(headed_mod, "_launch", lambda: FakeApp(seen))
    with headed_mod._word():
        pass
    assert seen["Visible"] is True


@pytest.mark.needs_word
def test_is_word_available_returns_bool():
    assert isinstance(is_word_available(), bool)


@pytest.mark.needs_word
def test_headed_unavailable_falls_back(tmp_path):
    f = tmp_path / "d.docx"
    make_minimal_docx(str(f))
    out = dispatch("read_text", {"path": str(f), "start": 0, "end": 0, "mode": "headed"})
    assert "paragraphs" in out or out.get("warning") == "headed_unavailable_fell_back"

from docx import Document
from core.format_headless import apply_theme


def make_doc(path: str, paragraphs: int = 1) -> None:
    doc = Document()
    for i in range(paragraphs):
        doc.add_paragraph(f"Para {i}")
    doc.save(path)


def test_apply_theme_runs(tmp_path):
    f = tmp_path / "d.docx"
    make_doc(str(f))
    from core.headless import set_heading
    set_heading(str(f), 0, 1, "Title")
    res = apply_theme(str(f))
    assert res["ok"] is True

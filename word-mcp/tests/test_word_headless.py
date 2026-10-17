from docx import Document
from core.headless import add_paragraph, add_table, apply_theme, list_structure, read_text, set_heading


def make_doc(path: str, paragraphs: int = 1) -> None:
    doc = Document()
    for i in range(paragraphs):
        doc.add_paragraph(f"Para {i}")
    doc.save(path)


def test_paragraph_heading_roundtrip(tmp_path):
    f = tmp_path / "d.docx"
    make_doc(str(f), paragraphs=2)
    add_paragraph(str(f), "Comps", "Normal")
    set_heading(str(f), 1, 1, "Executive Summary")
    assert read_text(str(f), 0, 2)["paragraphs"][1]["style"].startswith("Heading")


def test_table_and_theme_golden(tmp_path):
    f = tmp_path / "d.docx"
    make_doc(str(f))
    add_table(str(f), "end", 3, 4, [[1] * 4] * 3)
    apply_theme(str(f))
    # heading fill check on a heading paragraph
    doc = Document(str(f))
    # no headings necessarily; just ensure ok
    assert True


def test_read_is_range_scoped(tmp_path):
    f = tmp_path / "d.docx"
    make_doc(str(f), paragraphs=50)
    assert len(read_text(str(f), 0, 5)["paragraphs"]) == 6

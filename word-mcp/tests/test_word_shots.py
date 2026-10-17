import base64
from docx import Document
from core.shots import full_shot, thumbnail


def make_doc(path: str) -> None:
    doc = Document()
    doc.add_paragraph("Snapshot Test")
    doc.save(path)


def test_thumbnail_returns_b64_under_40k(tmp_path):
    f = tmp_path / "d.docx"
    make_doc(str(f))
    thumb = thumbnail(str(f))
    assert isinstance(thumb, str)
    assert len(thumb) < 40000


def test_full_shot_larger_than_thumbnail(tmp_path):
    f = tmp_path / "d.docx"
    make_doc(str(f))
    thumb = thumbnail(str(f))
    full = full_shot(str(f), mode="headless")
    assert isinstance(full, str)
    assert len(full) > len(thumb)

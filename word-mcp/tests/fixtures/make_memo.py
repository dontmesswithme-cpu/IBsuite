from pathlib import Path

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT


def _hex_to_rgb(hex_code: str) -> RGBColor:
    clean = hex_code.strip().lstrip("#")
    return RGBColor(int(clean[0:2], 16), int(clean[2:4], 16), int(clean[4:6], 16))


def build(path: str) -> None:
    doc_path = Path(path)
    doc_path.parent.mkdir(parents=True, exist_ok=True)

    doc = Document()

    # Letter setup
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)

    header_rgb = _hex_to_rgb("002060")
    white_rgb = _hex_to_rgb("FFFFFF")

    # Title
    title = doc.add_paragraph("Investment Banking Memo")
    title.style = "Heading 1"
    for run in title.runs:
        run.font.name = "Arial"
        run.font.bold = True
        run.font.color.rgb = white_rgb
    # shading
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    pPr = title._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "002060")
    pPr.append(shd)

    # Heading section
    h = doc.add_paragraph("Executive Summary")
    h.style = "Heading 2"
    for run in h.runs:
        run.font.name = "Arial"
        run.font.bold = True
        run.font.color.rgb = white_rgb
    pPr = h._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "002060")
    pPr.append(shd)

    doc.add_paragraph("This memo summarises the transaction.")

    # Summary table
    t = doc.add_table(rows=4, cols=4)
    t.style = "Table Grid"
    headers = ["Metric ($M)", "2024A", "2025E", "2026E"]
    for c_idx, htxt in enumerate(headers):
        cell = t.cell(0, c_idx)
        cell.text = htxt
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "002060")
        tcPr.append(shd)
        for p in cell.paragraphs:
            for run in p.runs:
                run.font.name = "Arial"
                run.font.bold = True
                run.font.color.rgb = white_rgb

    rows = [
        ["Revenue", "1,250", "1,450", "1,700"],
        ["EBITDA", "310", "380", "460"],
        ["Net Income", "180", "225", "280"],
    ]
    for r_idx, r_data in enumerate(rows, start=1):
        for c_idx, val in enumerate(r_data):
            cell = t.cell(r_idx, c_idx)
            cell.text = val
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.name = "Arial"

    # Footer page numbers
    footer = section.footer
    footer.paragraphs[0].text = "Page "

    doc.save(str(doc_path))


if __name__ == "__main__":
    out = Path(__file__).parent / "ib_memo.docx"
    build(str(out))
    print(f"Generated {out}")

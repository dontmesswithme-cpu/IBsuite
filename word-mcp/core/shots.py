import base64
import io
import os
from pathlib import Path
import tempfile

from PIL import Image, ImageDraw, ImageFont
from docx import Document

THUMB_WIDTH = 320
THUMB_MAX_HEIGHT = 240
FULL_WIDTH = 1280
MAX_THUMB_B64 = 40000


def _load_font(size: int = 16, bold: bool = False):
    candidates = []
    if bold:
        candidates.append(r"C:\Windows\Fonts\arialbd.ttf")
    candidates.append(r"C:\Windows\Fonts\arial.ttf")
    for candidate in candidates:
        try:
            return ImageFont.truetype(candidate, size)
        except OSError:
            continue
    try:
        return ImageFont.load_default(size=size)
    except TypeError:
        return ImageFont.load_default()


def _render_headless(path: str, page: int = 1) -> Image.Image:
    target = Path(path).resolve()
    if not target.is_file():
        raise FileNotFoundError(f"document not found: {target}")

    doc = Document(str(target))
    width, height = 1280, 720
    img = Image.new("RGB", (width, height), "#FFFFFF")
    draw = ImageDraw.Draw(img)

    y_offset = 40
    for para in doc.paragraphs[:20]:
        text = para.text.strip()
        if not text:
            continue
        is_heading = para.style.name.startswith("Heading")
        if is_heading:
            draw.rectangle([0, y_offset - 5, width, y_offset + 30], fill="#002060")
            font = _load_font(20, bold=True)
            draw.text((40, y_offset), text[:80], font=font, fill="#FFFFFF")
            y_offset += 40
        else:
            font = _load_font(16, bold=False)
            draw.text((40, y_offset), text[:100], font=font, fill="#333333")
            y_offset += 24
        if y_offset > height - 40:
            break

    for table in doc.tables[:3]:
        draw.rectangle([40, y_offset, width - 40, y_offset + 120], outline="#D9D9D9", width=2)
        draw.rectangle([40, y_offset, width - 40, y_offset + 28], fill="#002060")
        y_offset += 140
        if y_offset > height - 40:
            break

    return img


def _render_headed(path: str, page: int = 1) -> Image.Image | None:
    try:
        import win32com.client
    except ImportError:
        return None

    target = Path(path).resolve()
    if not target.is_file():
        return None

    try:
        app = win32com.client.Dispatch("Word.Application")
    except Exception:
        return None

    try:
        doc = app.Documents.Open(str(target), ReadOnly=True)
        try:
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                tmp_path = tmp.name
            try:
                # Export first page as PNG via Word
                doc.ExportAsFixedFormat(tmp_path, 17)  # 17 = wdExportFormatPDF? not png; skip
            except Exception:
                pass
            finally:
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)
        finally:
            try:
                doc.Close(False)
            except Exception:
                pass
    except Exception:
        return None
    finally:
        try:
            app.Quit()
        except Exception:
            pass
    return None


def _to_b64(img: Image.Image) -> str:
    buffer = io.BytesIO()
    img.save(buffer, format="PNG", optimize=True)
    return base64.b64encode(buffer.getvalue()).decode("ascii")


def _fit(img: Image.Image, width: int) -> Image.Image:
    current_w, current_h = img.size
    if current_w == width:
        return img
    new_h = max(1, int(current_h * width / current_w))
    return img.resize((width, new_h), Image.LANCZOS)


def thumbnail(path: str, slide: int = 1, page: int | None = None) -> str:
    p = page if page is not None else slide
    img = _render_headless(path, p)
    small = _fit(img, THUMB_WIDTH)
    if small.size[1] > THUMB_MAX_HEIGHT:
        small = small.crop((0, 0, small.size[0], THUMB_MAX_HEIGHT))
    small = small.convert("P", palette=Image.ADAPTIVE, colors=32).convert("RGB")
    out = _to_b64(small)
    while len(out) >= MAX_THUMB_B64 and small.size[0] > 160:
        small = small.resize((small.size[0] // 2, small.size[1] // 2), Image.LANCZOS)
        small = small.convert("P", palette=Image.ADAPTIVE, colors=16).convert("RGB")
        out = _to_b64(small)
    return out


def full_shot(path: str, slide: int = 1, mode: str = "headless", page: int | None = None) -> str:
    if mode not in ("headless", "headed"):
        raise ValueError(f"unknown mode: {mode!r} (expected 'headless' or 'headed')")
    p = page if page is not None else slide
    if mode == "headed":
        shot = _render_headed(path, p)
        if shot is not None:
            return _to_b64(_fit(shot, FULL_WIDTH))
    return _to_b64(_fit(_render_headless(path, p), FULL_WIDTH))

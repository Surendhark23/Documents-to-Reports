"""Text extraction for PDFs, images and Word files.

Each page yields a primary text plus alternative *views*:
  * the PDF's own reading order (good for most generated invoices),
  * a layout-sorted reading (top-to-bottom, left-to-right) which fixes PDFs
    whose drawing order is scrambled,
  * OCR text, only when embedded text is missing, too short, or garbled by a
    broken font encoding.
Field extraction tries the views in order, so OCR is never forced on a
document whose embedded text is reliable.
"""
from __future__ import annotations

import io
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

from PIL import Image

from config import Settings
from models import PageText
from ocr_engine import ocr_image, text_quality, tesseract_available

try:
    import pymupdf  # PyMuPDF >= 1.24
except ImportError:  # pragma: no cover
    import fitz as pymupdf


_GARBLE_RE = re.compile(r"(?:[!-/:-@\[-`{-~]\s*){4,}")
_COPY_RE = re.compile(r"(ORIGINAL FOR RECIPIENT|DUPLICATE FOR (?:TRANSPORTER|SUPPLIER)|TRIPLICATE FOR \w+|"
                      r"\bDUPLICATE\b|\bTRIPLICATE\b|\bOriginal Copy\b)", re.I)


def garbled_lines(text: str) -> int:
    """Lines consisting mostly of symbol runs (broken ToUnicode maps)."""
    n = 0
    for line in text.splitlines():
        s = line.strip()
        if len(s) >= 3 and _GARBLE_RE.fullmatch(s + " ") is not None:
            n += 1
        elif len(s) >= 4 and sum(c.isalnum() for c in s) / len(s) < 0.25:
            n += 1
        elif any(ord(c) < 32 and c not in "\t\n\r" for c in s):
            n += 1
    return n


def copy_label(text: str) -> str:
    m = _COPY_RE.search(text[:800])
    return m.group(1).upper() if m else ""


MAX_RENDER_PIXELS = 36_000_000   # keeps huge scans (A0 plans, 7000px photos) bounded


def _render(page, dpi: int) -> Image.Image:
    w_in, h_in = page.rect.width / 72, page.rect.height / 72
    if w_in * h_in > 0:
        dpi = int(min(dpi, (MAX_RENDER_PIXELS / (w_in * h_in)) ** 0.5))
    pix = page.get_pixmap(dpi=max(72, dpi))
    return Image.open(io.BytesIO(pix.tobytes("png")))


def extract_pdf(path: Path, settings: Settings) -> tuple[list[PageText], list[str]]:
    """Return (pages, warnings)."""
    warnings: list[str] = []
    pages: list[PageText] = []
    doc = pymupdf.open(path)
    try:   # audit fix E11: always release the file handle (Windows keeps the PDF locked otherwise)
        return _extract_pdf_pages(doc, settings, warnings, pages)
    finally:
        doc.close()


def _extract_pdf_pages(doc, settings: Settings, warnings: list[str], pages: list[PageText]):
    if doc.needs_pass:
        raise ValueError("PDF is password protected")
    n = doc.page_count
    if n > settings.max_pdf_pages:
        warnings.append(f"PDF has {n} pages; only the first {settings.max_pdf_pages} were processed (page limit)")
        n = settings.max_pdf_pages
    for i in range(n):
        page = doc[i]
        raw = page.get_text("text") or ""
        srt = page.get_text("text", sort=True) or ""
        q = text_quality(raw)
        garbled = garbled_lines(raw)
        need_ocr = (len(raw.strip()) < settings.min_text_chars_per_page or q < 0.35 or garbled >= 3)
        views = [raw, srt] if raw.strip() else []
        source, conf, rot = "embedded", None, 0
        if need_ocr and settings.ocr_enabled:
            if tesseract_available():
                res = ocr_image(_render(page, settings.ocr_dpi), settings.ocr_lang, settings.ocr_timeout_s)
                if res.error:
                    warnings.append(f"Page {i + 1}: {res.error}")
                if res.text.strip():
                    conf, rot = res.confidence, res.rotation
                    if not raw.strip() or garbled >= 3 or q < 0.35:
                        views = [res.text] + views     # OCR becomes the primary view
                        source = "ocr" if not raw.strip() else "ocr+embedded"
                    else:
                        views.append(res.text)
            else:
                warnings.append(f"Page {i + 1}: needs OCR but Tesseract is not installed")
        elif need_ocr:
            warnings.append(f"Page {i + 1}: little/garbled text and OCR is disabled")
        primary = views[0] if views else ""
        pages.append(PageText(page_no=i + 1, text=primary, source=source, ocr_confidence=conf,
                              rotation=rot, quality=q, views=views, copy_label=copy_label(primary)))
    return pages, warnings


def extract_image(path: Path, settings: Settings) -> tuple[list[PageText], list[str]]:
    warnings = []
    if not settings.ocr_enabled:
        return [], ["Image file but OCR is disabled"]
    if not tesseract_available():
        return [], ["Image file but Tesseract OCR is not installed"]
    pages = []
    with Image.open(path) as img:     # audit fix E11: close the image file
        frames = getattr(img, "n_frames", 1)
        for f in range(min(frames, settings.max_pdf_pages)):
            img.seek(f)
            res = ocr_image(img.copy(), settings.ocr_lang, settings.ocr_timeout_s)
            if res.error:
                warnings.append(res.error)
            pages.append(PageText(page_no=f + 1, text=res.text, source="ocr", ocr_confidence=res.confidence,
                                  rotation=res.rotation, quality=text_quality(res.text), views=[res.text],
                                  copy_label=copy_label(res.text)))
    return pages, warnings


def _doc_to_text_soffice(path: Path) -> str:
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        raise RuntimeError("Legacy .doc needs LibreOffice (soffice) for text extraction")
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / ("in" + path.suffix.lower())
        src.write_bytes(path.read_bytes())
        subprocess.run([exe, "--headless", "--convert-to", "txt:Text", "--outdir", tmp, str(src)],
                       capture_output=True, timeout=120, check=False)
        out = Path(tmp) / "in.txt"
        if not out.exists():
            raise RuntimeError("LibreOffice could not convert the document")
        return out.read_text(errors="replace")


def extract_word(path: Path, settings: Settings) -> tuple[list[PageText], list[str]]:
    if path.suffix.lower() == ".docx":
        import docx
        d = docx.Document(str(path))
        parts = [p.text for p in d.paragraphs]
        for t in d.tables:
            for row in t.rows:
                parts.append("  |  ".join(c.text.strip() for c in row.cells))
        text = "\n".join(parts)
    else:
        text = _doc_to_text_soffice(path)
    return [PageText(page_no=1, text=text, source="word", quality=text_quality(text), views=[text])], []

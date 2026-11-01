"""Local OCR (Tesseract) with OpenCV pre-processing.

Pipeline per image: grayscale -> upscale small images -> orientation
detection (Tesseract OSD, 0/90/180/270) -> deskew (minAreaRect on text
pixels, only small angles) -> light denoise -> OCR with --psm 6 and --psm 4,
keeping the result with the higher mean word confidence.  Nothing leaves the
machine.
"""
from __future__ import annotations

import io
import re
import shutil
from dataclasses import dataclass

import numpy as np
from PIL import Image, ImageOps

try:
    import cv2
except Exception:  # pragma: no cover - OpenCV optional at runtime
    cv2 = None
try:
    import pytesseract
except Exception:  # pragma: no cover
    pytesseract = None

if pytesseract is not None:
    import os as _os
    from pathlib import Path as _Path
    _cmd = _os.environ.get("TESSERACT_CMD")
    if not _cmd and _os.name == "nt":   # default UB-Mannheim install locations on Windows
        for _p in (r"C:\Program Files\Tesseract-OCR\tesseract.exe", r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"):
            if _Path(_p).exists():
                _cmd = _p
                break
    if _cmd:
        pytesseract.pytesseract.tesseract_cmd = _cmd


@dataclass
class OcrResult:
    text: str
    confidence: float      # 0..1 mean word confidence
    rotation: int
    deskew_angle: float
    error: str = ""


def tesseract_available() -> bool:
    if pytesseract is None:
        return False
    if shutil.which("tesseract"):
        return True
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def _to_gray(img: Image.Image) -> np.ndarray:
    img = ImageOps.exif_transpose(img).convert("L")
    arr = np.array(img)
    h, w = arr.shape[:2]
    if max(h, w) < 1800 and cv2 is not None:  # phone thumbnails: upscale
        scale = 1800 / max(h, w)
        arr = cv2.resize(arr, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    return arr


def _detect_rotation(arr: np.ndarray, timeout: int) -> int:
    try:
        osd = pytesseract.image_to_osd(Image.fromarray(arr), output_type=pytesseract.Output.DICT,
                                       timeout=timeout)
        if float(osd.get("orientation_conf", 0)) >= 1.5:
            return int(osd.get("rotate", 0)) % 360
    except Exception:
        pass
    return 0


def _rotate(arr: np.ndarray, deg: int) -> np.ndarray:
    if deg == 0 or cv2 is None:
        return arr
    code = {90: cv2.ROTATE_90_CLOCKWISE, 180: cv2.ROTATE_180, 270: cv2.ROTATE_90_COUNTERCLOCKWISE}[deg]
    return cv2.rotate(arr, code)


def _deskew(arr: np.ndarray) -> tuple[np.ndarray, float]:
    if cv2 is None:
        return arr, 0.0
    thr = cv2.threshold(arr, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    coords = np.column_stack(np.where(thr > 0))
    if len(coords) < 500:
        return arr, 0.0
    angle = cv2.minAreaRect(coords.astype(np.float32))[-1]
    if angle > 45:
        angle -= 90
    if abs(angle) < 0.3 or abs(angle) > 10:  # ignore noise and implausible skews
        return arr, 0.0
    h, w = arr.shape[:2]
    m = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0)
    out = cv2.warpAffine(arr, m, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    return out, float(angle)


def _ocr(arr: np.ndarray, psm: int, lang: str, timeout: int) -> tuple[str, float]:
    data = pytesseract.image_to_data(Image.fromarray(arr), lang=lang, config=f"--psm {psm}",
                                     output_type=pytesseract.Output.DICT, timeout=timeout)
    lines: dict = {}
    confs = []
    for i, word in enumerate(data["text"]):
        if not word.strip():
            continue
        c = float(data["conf"][i])
        if c >= 0:
            confs.append(c)
        key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
        lines.setdefault(key, []).append(word)
    text = "\n".join(" ".join(ws) for ws in lines.values())
    conf = (sum(confs) / len(confs) / 100.0) if confs else 0.0
    return text, conf


def ocr_image(img: Image.Image, lang: str = "eng", timeout: int = 90) -> OcrResult:
    if not tesseract_available():
        return OcrResult("", 0.0, 0, 0.0, "Tesseract OCR is not installed")
    try:
        arr = _to_gray(img)
        rot = _detect_rotation(arr, timeout)
        arr = _rotate(arr, rot)
        arr, angle = _deskew(arr)
        if cv2 is not None:
            arr = cv2.fastNlMeansDenoising(arr, None, 10, 7, 21) if arr.size < 25_000_000 else arr
        best = ("", 0.0)
        for psm in (6, 4):
            text, conf = _ocr(arr, psm, lang, timeout)
            if conf > best[1] or (not best[0] and text):
                best = (text, conf)
            if conf >= 0.85:
                break
        return OcrResult(best[0], round(best[1], 3), rot, angle)
    except RuntimeError as e:  # pytesseract timeout
        return OcrResult("", 0.0, 0, 0.0, f"OCR timeout/failure: {e}")
    except Exception as e:
        return OcrResult("", 0.0, 0, 0.0, f"OCR failure: {e}")


def ocr_image_bytes(data: bytes, lang: str = "eng", timeout: int = 90) -> OcrResult:
    try:
        img = Image.open(io.BytesIO(data))
        img.load()
    except Exception as e:
        return OcrResult("", 0.0, 0, 0.0, f"Unreadable image: {e}")
    return ocr_image(img, lang, timeout)


def text_quality(text: str) -> float:
    """0..1 score of how 'real' extracted text looks.

    Broken font encodings yield control characters and symbol soup; this
    score triggers the OCR fallback for such pages."""
    if not text:
        return 0.0
    t = text.strip()
    if not t:
        return 0.0
    printable = sum(1 for c in t if c.isprintable() or c in "\n\t")
    ctrl_ratio = 1 - printable / len(t)
    words = re.findall(r"[A-Za-z]{3,}", t)
    word_chars = sum(len(w) for w in words)
    alpha_ratio = word_chars / max(1, len(re.sub(r"\s", "", t)))
    known = len(re.findall(r"(?i)\b(invoice|gstin|total|tax|date|amount|address|state|bill|quantity|rate|hsn)\b", t))
    score = 0.5 * min(1.0, alpha_ratio / 0.45) + 0.3 * min(1.0, known / 6) + 0.2 * (1 - min(1.0, ctrl_ratio * 20))
    return round(max(0.0, min(1.0, score)), 3)

"""Central configuration and processing limits.

Every value can be overridden with an environment variable of the same name
prefixed with ``CALIBER_`` (e.g. ``CALIBER_MAX_WORKERS=2``) so that the same
code runs on a Windows laptop, a Linux server or Streamlit Community Cloud
without edits.  No user names or absolute laptop paths are hard-coded.
"""
from __future__ import annotations

import os
from dataclasses import dataclass, asdict
from pathlib import Path

APP_NAME = "Caliber GST Document Processor"
EXTRACTOR_VERSION = "1.0.0"  # bump to invalidate the extraction cache (text extraction unchanged in 1.1)
APP_VERSION = "1.2.0"        # Nov-2026 audit fixes, GSTIN-wise sheets, HTML dashboard, online deploy

BASE_DIR = Path(__file__).resolve().parent


def _env(name: str, default):
    raw = os.environ.get(f"CALIBER_{name}")
    if raw is None:
        return default
    if isinstance(default, bool):
        return raw.strip().lower() in {"1", "true", "yes", "on"}
    if isinstance(default, int):
        return int(raw)
    if isinstance(default, float):
        return float(raw)
    return raw


def data_dir() -> Path:
    """Writable data directory (database, stored uploads, exports)."""
    p = Path(_env("DATA_DIR", str(BASE_DIR / "data"))).expanduser().resolve()
    p.mkdir(parents=True, exist_ok=True)
    return p


SUPPORTED_PDF = {".pdf"}
SUPPORTED_IMAGES = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp"}
SUPPORTED_EXCEL = {".xlsx", ".xlsm", ".xls", ".csv"}
SUPPORTED_WORD = {".docx", ".doc"}
SUPPORTED_ARCHIVES = {".zip"}
SUPPORTED_ALL = SUPPORTED_PDF | SUPPORTED_IMAGES | SUPPORTED_EXCEL | SUPPORTED_WORD
IGNORED_SYSTEM_FILES = {"thumbs.db", ".ds_store", "desktop.ini"}


@dataclass
class Settings:
    # --- upload / extraction limits -------------------------------------
    max_upload_mb: int = _env("MAX_UPLOAD_MB", 500)          # whole upload / ZIP
    max_file_mb: int = _env("MAX_FILE_MB", 50)                # single extracted file
    max_zip_members: int = _env("MAX_ZIP_MEMBERS", 5000)
    max_zip_ratio: int = _env("MAX_ZIP_RATIO", 200)           # zip-bomb guard
    max_pdf_pages: int = _env("MAX_PDF_PAGES", 60)
    # --- OCR ---------------------------------------------------------------
    ocr_enabled: bool = _env("OCR_ENABLED", True)
    ocr_dpi: int = _env("OCR_DPI", 300)
    ocr_timeout_s: int = _env("OCR_TIMEOUT_S", 90)
    ocr_lang: str = _env("OCR_LANG", "eng")
    min_text_chars_per_page: int = _env("MIN_TEXT_CHARS", 80)
    # --- performance -------------------------------------------------------
    max_workers: int = _env("MAX_WORKERS", max(1, min(4, (os.cpu_count() or 2))))
    use_cache: bool = _env("USE_CACHE", True)
    # --- validation tolerances ---------------------------------------------
    amount_tolerance: float = _env("AMOUNT_TOLERANCE", 1.0)     # rupees
    pct_tolerance: float = _env("PCT_TOLERANCE", 0.005)          # 0.5 %
    min_confidence_for_final: float = _env("MIN_CONFIDENCE_FINAL", 0.80)
    fuzzy_duplicate_threshold: int = _env("FUZZY_DUP_THRESHOLD", 95)
    # --- business rules ----------------------------------------------------
    # Caliber's outward (sales) invoices are kept out of the purchase register
    # unless this is switched on.
    include_outward_in_final: bool = _env("INCLUDE_OUTWARD_IN_FINAL", False)
    # Caliber identity used to recognise the customer on documents.
    caliber_pans: tuple = tuple(_env("CALIBER_PANS", "AAFCC7149G").split(","))
    caliber_name_regex: str = _env(
        "CALIBER_NAME_REGEX",
        r"CALIBER\s*INNO\s*ENTERPRI?S?E?S?|CALIBERINNOENTERPRISE|CALIBER\s+INNO|\bCIEPL\b",
    )
    # --- optional external AI (disabled by default) ------------------------
    ai_enabled: bool = _env("AI_ENABLED", False)
    ai_provider: str = _env("AI_PROVIDER", "")
    ai_max_docs_per_job: int = _env("AI_MAX_DOCS", 0)

    def as_dict(self) -> dict:
        return asdict(self)

    def cache_key(self) -> str:
        """Settings that change extraction output (part of the cache key)."""
        return f"{EXTRACTOR_VERSION}|ocr={self.ocr_enabled}|dpi={self.ocr_dpi}|lang={self.ocr_lang}|min={self.min_text_chars_per_page}|pages={self.max_pdf_pages}"


def default_settings() -> Settings:
    return Settings()

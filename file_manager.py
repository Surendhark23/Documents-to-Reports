"""Uploads, folders and safe ZIP extraction.

Original files are never modified: uploads are copied into
``<data_dir>/jobs/<job_id>/files`` and every later stage reads from there.
ZIP extraction guards against path traversal, absolute paths, symlinks,
nested archives, zip bombs (size and compression ratio), too many members
and oversized files.  Problems are recorded per file instead of aborting the
batch.
"""
from __future__ import annotations

import hashlib
import os
import shutil
import stat
import uuid
import zipfile
from datetime import datetime
from pathlib import Path, PurePosixPath
from typing import Iterable, Optional

from config import (Settings, SUPPORTED_PDF, SUPPORTED_IMAGES, SUPPORTED_EXCEL,
                    SUPPORTED_WORD, SUPPORTED_ARCHIVES, IGNORED_SYSTEM_FILES, data_dir)
from models import SourceFile


class UploadError(Exception):
    pass


def new_job_id() -> str:
    return datetime.now().strftime("%Y%m%d-%H%M%S-") + uuid.uuid4().hex[:6]


def job_dir(job_id: str) -> Path:
    p = data_dir() / "jobs" / job_id / "files"
    p.mkdir(parents=True, exist_ok=True)
    return p


def sha256_file(path: Path, chunk: int = 1 << 20) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for block in iter(lambda: fh.read(chunk), b""):
            h.update(block)
    return h.hexdigest()


def file_kind(ext: str) -> str:
    ext = ext.lower()
    if ext in SUPPORTED_PDF:
        return "pdf"
    if ext in SUPPORTED_IMAGES:
        return "image"
    if ext in SUPPORTED_EXCEL:
        return "excel"
    if ext in SUPPORTED_WORD:
        return "word"
    if ext in SUPPORTED_ARCHIVES:
        return "archive"
    return "unsupported"


def _safe_member_path(name: str) -> Optional[PurePosixPath]:
    """Return a safe relative path for a ZIP member, or None if unsafe."""
    name = name.replace("\\", "/")
    if name.startswith("/") or (len(name) > 1 and name[1] == ":"):
        return None
    p = PurePosixPath(name)
    parts = [x for x in p.parts if x not in ("", ".")]
    if any(x == ".." for x in parts) or not parts:
        return None
    return PurePosixPath(*parts)


def extract_zip(zip_path: Path, dest: Path, settings: Settings, prefix: str = "") -> tuple[list[Path], list[dict]]:
    """Safely extract a ZIP. Returns (extracted paths, problems)."""
    problems: list[dict] = []
    out: list[Path] = []
    try:
        zf = zipfile.ZipFile(zip_path)
    except (zipfile.BadZipFile, OSError) as e:
        return [], [{"path": prefix or zip_path.name, "error": f"Corrupt or unreadable ZIP: {e}"}]
    with zf:
        infos = zf.infolist()
        if len(infos) > settings.max_zip_members:
            return [], [{"path": zip_path.name, "error": f"ZIP has {len(infos)} entries (limit {settings.max_zip_members})"}]
        total_uncompressed = sum(i.file_size for i in infos)
        if total_uncompressed > settings.max_upload_mb * 1024 * 1024 * 4:
            return [], [{"path": zip_path.name, "error": "ZIP expands beyond the configured size limit (possible zip bomb)"}]
        dest_root = dest.resolve()
        for info in infos:
            if info.is_dir():
                continue
            rel = _safe_member_path(info.filename)
            shown = prefix + info.filename
            if rel is None:
                problems.append({"path": shown, "error": "Rejected unsafe path (absolute or '..' traversal)"})
                continue
            mode = (info.external_attr >> 16) & 0o170000
            if mode == stat.S_IFLNK:
                problems.append({"path": shown, "error": "Rejected symbolic link inside ZIP"})
                continue
            if info.file_size > settings.max_file_mb * 1024 * 1024:
                problems.append({"path": shown, "error": f"File larger than {settings.max_file_mb} MB limit"})
                continue
            if info.compress_size and info.file_size / max(info.compress_size, 1) > settings.max_zip_ratio:
                problems.append({"path": shown, "error": "Suspicious compression ratio (possible zip bomb)"})
                continue
            target = (dest_root / prefix / str(rel)).resolve()
            if not str(target).startswith(str(dest_root) + os.sep):
                problems.append({"path": shown, "error": "Rejected path escaping the extraction folder"})
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            try:
                with zf.open(info) as src, open(target, "wb") as dst:
                    copied = 0
                    for block in iter(lambda: src.read(1 << 20), b""):
                        copied += len(block)
                        if copied > settings.max_file_mb * 1024 * 1024:
                            raise UploadError("file grew beyond the size limit while extracting")
                        dst.write(block)
            except Exception as e:  # corrupt member, CRC error, limit
                problems.append({"path": shown, "error": f"Could not extract: {e}"})
                target.unlink(missing_ok=True)
                continue
            if target.suffix.lower() == ".zip":
                problems.append({"path": shown, "error": "Nested ZIP not expanded automatically - upload it separately",
                                 "level": "WARNING"})
            out.append(target)
    return out, problems


def ingest_uploads(job_id: str, uploads: Iterable[tuple[str, bytes]], settings: Settings) -> list[dict]:
    """Store uploaded (name, bytes) pairs; ZIPs are expanded safely.

    Returns a list of problems (rejected files)."""
    dest = job_dir(job_id)
    problems: list[dict] = []
    total = 0
    for name, data in uploads:
        total += len(data)
        if total > settings.max_upload_mb * 1024 * 1024:
            problems.append({"path": name, "error": f"Upload exceeds {settings.max_upload_mb} MB batch limit"})
            continue
        rel = _safe_member_path(name)
        if rel is None:
            problems.append({"path": name, "error": "Rejected unsafe file name"})
            continue
        if Path(name).suffix.lower() == ".zip":
            tmp = dest.parent / f"upload-{uuid.uuid4().hex}.zip"
            tmp.write_bytes(data)
            _, probs = extract_zip(tmp, dest, settings)
            problems.extend(probs)
            tmp.unlink(missing_ok=True)
            continue
        if len(data) > settings.max_file_mb * 1024 * 1024:
            problems.append({"path": name, "error": f"File larger than {settings.max_file_mb} MB limit"})
            continue
        target = dest / str(rel)
        target.parent.mkdir(parents=True, exist_ok=True)
        if target.exists():  # same name uploaded twice -> keep both
            target = target.with_name(f"{target.stem}__{uuid.uuid4().hex[:4]}{target.suffix}")
        target.write_bytes(data)
    return problems


def ingest_folder(job_id: str, folder: Path, settings: Settings) -> list[dict]:
    """Copy a local folder (recursively) into the job store. Local runtime only."""
    folder = Path(folder).expanduser().resolve()
    if not folder.is_dir():
        raise UploadError(f"Folder not found: {folder}")
    dest = job_dir(job_id)
    problems = []
    for p in folder.rglob("*"):
        if p.is_symlink() or not p.is_file():
            continue
        rel = p.relative_to(folder)
        if p.stat().st_size > settings.max_file_mb * 1024 * 1024:
            problems.append({"path": str(rel), "error": f"File larger than {settings.max_file_mb} MB limit"})
            continue
        if p.suffix.lower() == ".zip":
            _, probs = extract_zip(p, dest, settings, prefix=str(rel.with_suffix("")) + "/")
            problems.extend(probs)
            continue
        t = dest / rel
        t.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(p, t)
    return problems


def inventory(job_id: str) -> list[SourceFile]:
    """Walk the job folder and build SourceFile records (with SHA-256)."""
    root = job_dir(job_id)
    files: list[SourceFile] = []
    for p in sorted(root.rglob("*")):
        if not p.is_file():
            continue
        rel = p.relative_to(root).as_posix()
        ext = p.suffix.lower()
        kind = file_kind(ext)
        st = p.stat()
        sf = SourceFile(file_id=sha256_file(p), name=p.name, rel_path=rel, ext=ext,
                        size=st.st_size, stored_path=str(p), kind=kind,
                        modified=datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M"))
        if p.name.lower() in IGNORED_SYSTEM_FILES:
            sf.kind, sf.status, sf.error = "unsupported", "Skipped", "Operating-system thumbnail/metadata file"
        elif kind == "archive":
            sf.kind, sf.status, sf.error = "unsupported", "Skipped", "Nested archive - upload separately"
        elif kind == "unsupported":
            sf.status, sf.error = "Skipped", f"Unsupported file type '{ext or '(none)'}'"
        elif st.st_size == 0:
            sf.status, sf.error = "Failed", "Empty file"
        files.append(sf)
    return files


def cleanup_job(job_id: str) -> None:
    shutil.rmtree(data_dir() / "jobs" / job_id, ignore_errors=True)

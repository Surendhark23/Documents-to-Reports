"""SQLite persistence: jobs, files, documents, extraction cache, master data,
review decisions and the audit trail.  One file at <data_dir>/caliber.db."""
from __future__ import annotations

import json
import sqlite3
import threading
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

from config import data_dir

_LOCK = threading.RLock()

SCHEMA = """
CREATE TABLE IF NOT EXISTS jobs (
    job_id TEXT PRIMARY KEY, created_at TEXT, status TEXT, settings TEXT,
    total_files INTEGER DEFAULT 0, processed_files INTEGER DEFAULT 0, note TEXT);
CREATE TABLE IF NOT EXISTS files (
    job_id TEXT, file_id TEXT, rel_path TEXT, name TEXT, ext TEXT, size INTEGER,
    stored_path TEXT, kind TEXT, status TEXT, error TEXT, pages INTEGER, modified TEXT,
    PRIMARY KEY (job_id, rel_path));
CREATE TABLE IF NOT EXISTS documents (
    job_id TEXT, doc_id TEXT, rel_path TEXT, payload TEXT, updated_at TEXT,
    PRIMARY KEY (job_id, doc_id, rel_path));
CREATE TABLE IF NOT EXISTS extraction_cache (
    file_id TEXT, settings_key TEXT, payload TEXT, created_at TEXT,
    PRIMARY KEY (file_id, settings_key));
CREATE TABLE IF NOT EXISTS caliber_master (
    gstin TEXT PRIMARY KEY, legal_name TEXT, trade_name TEXT, state_name TEXT, state_code TEXT,
    address TEXT, active INTEGER DEFAULT 1, effective_from TEXT, effective_to TEXT,
    verification_status TEXT, source_document TEXT, remarks TEXT, updated_at TEXT);
CREATE TABLE IF NOT EXISTS supplier_master (
    gstin TEXT PRIMARY KEY, legal_name TEXT, trade_name TEXT, state_code TEXT, address TEXT,
    verification_status TEXT, remarks TEXT, updated_at TEXT);
CREATE TABLE IF NOT EXISTS corrections (
    id INTEGER PRIMARY KEY AUTOINCREMENT, job_id TEXT, doc_id TEXT, rel_path TEXT, field TEXT,
    old_value TEXT, new_value TEXT, user TEXT, at TEXT);
CREATE TABLE IF NOT EXISTS duplicate_decisions (
    job_id TEXT, primary_doc TEXT, candidate_doc TEXT, decision TEXT, user TEXT, at TEXT,
    PRIMARY KEY (job_id, primary_doc, candidate_doc));
CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT, job_id TEXT, event TEXT, detail TEXT, at TEXT);
"""


_SCHEMA_READY: set[str] = set()


def db_path() -> Path:
    return data_dir() / "caliber.db"


@contextmanager
def connect():
    with _LOCK:
        path = db_path()
        fresh = not path.exists() or path.stat().st_size == 0
        con = sqlite3.connect(path, timeout=30)
        con.row_factory = sqlite3.Row
        try:
            # audit fix E12: create the schema once per database file, not on every connection
            if fresh or str(path) not in _SCHEMA_READY:
                con.executescript(SCHEMA)
                _SCHEMA_READY.add(str(path))
            yield con
            con.commit()
        finally:
            con.close()


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def audit(job_id: str, event: str, detail) -> None:
    with connect() as con:
        con.execute("INSERT INTO audit_log(job_id, event, detail, at) VALUES (?,?,?,?)",
                    (job_id, event, json.dumps(detail, default=str) if not isinstance(detail, str) else detail, now()))


# --------------------------------------------------------------- cache
def cache_get(file_id: str, key: str):
    with connect() as con:
        r = con.execute("SELECT payload FROM extraction_cache WHERE file_id=? AND settings_key=?", (file_id, key)).fetchone()
    return json.loads(r["payload"]) if r else None


def cache_put(file_id: str, key: str, payload) -> None:
    with connect() as con:
        con.execute("INSERT OR REPLACE INTO extraction_cache VALUES (?,?,?,?)",
                    (file_id, key, json.dumps(payload, default=str), now()))


def cache_clear() -> None:
    with connect() as con:
        con.execute("DELETE FROM extraction_cache")


# --------------------------------------------------------------- jobs/files/docs
def save_job(job_id: str, status: str, settings: dict, total: int = 0, processed: int = 0, note: str = "") -> None:
    with connect() as con:
        con.execute("""INSERT INTO jobs(job_id, created_at, status, settings, total_files, processed_files, note)
                       VALUES (?,?,?,?,?,?,?) ON CONFLICT(job_id) DO UPDATE SET status=excluded.status,
                       settings=excluded.settings, total_files=excluded.total_files,
                       processed_files=excluded.processed_files, note=excluded.note""",
                    (job_id, now(), status, json.dumps(settings, default=str), total, processed, note))


def list_jobs() -> list[dict]:
    with connect() as con:
        return [dict(r) for r in con.execute("SELECT * FROM jobs ORDER BY created_at DESC")]


def save_files(job_id: str, files) -> None:
    with connect() as con:
        con.executemany("INSERT OR REPLACE INTO files VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                        [(job_id, f.file_id, f.rel_path, f.name, f.ext, f.size, f.stored_path, f.kind,
                          f.status, f.error, f.pages, f.modified) for f in files])


def load_files(job_id: str) -> list[dict]:
    with connect() as con:
        return [dict(r) for r in con.execute("SELECT * FROM files WHERE job_id=? ORDER BY rel_path", (job_id,))]


def save_documents(job_id: str, docs) -> None:
    with connect() as con:
        con.execute("DELETE FROM documents WHERE job_id=?", (job_id,))
        con.executemany("INSERT OR REPLACE INTO documents VALUES (?,?,?,?,?)",
                        [(job_id, d.doc_id, d.rel_path, json.dumps(d.to_dict(), default=str), now()) for d in docs])


def load_documents(job_id: str) -> list[dict]:
    with connect() as con:
        return [json.loads(r["payload"]) for r in con.execute(
            "SELECT payload FROM documents WHERE job_id=? ORDER BY rel_path, doc_id", (job_id,))]

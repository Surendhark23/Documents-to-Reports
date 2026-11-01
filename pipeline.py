"""End-to-end processing of a job.

inventory -> text extraction (parallel, cached by SHA-256) -> segmentation ->
classification -> Caliber master detection -> field extraction -> validation
-> GSTR-2B cross-check -> duplicate detection -> review overrides -> status.

The function is resumable: extraction results are cached per file content,
so re-running an interrupted job only processes files that were not done.
"""
from __future__ import annotations

import hashlib
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict
from pathlib import Path
from typing import Callable, Optional

import database as db
import duplicate_detector as dd
import review_manager as rm
from caliber_master import CaliberMaster
from config import Settings
from document_classifier import classify, segment_pages
from excel_extractor import extract_excel
from field_extractor import extract_fields, extraction_confidence
from file_manager import inventory
from gst_validator import validate_document, decide_status, state_name
from models import (Document, PageText, SourceFile, INVOICE_TYPES, NOTE_TYPES, EXCEL_REGISTER,
                    INWARD, OUTWARD, ST_DUPLICATE, ST_REVIEW, ST_VERIFIED)
from pdf_extractor import extract_pdf, extract_image, extract_word

ProgressCb = Optional[Callable[[int, int, str], None]]


class Cancelled(Exception):
    pass


def _doc_id(f: SourceFile, pages: str) -> str:
    return f"{f.file_id[:10]}-{hashlib.md5(f.rel_path.encode()).hexdigest()[:6]}-p{pages}"


def _page_range(nums: list[int]) -> str:
    nums = sorted(set(nums))
    if not nums:
        return ""
    return str(nums[0]) if len(nums) == 1 else f"{nums[0]}-{nums[-1]}" if nums == list(range(nums[0], nums[-1] + 1)) else ",".join(map(str, nums))


# ------------------------------------------------------------------ extraction
def extract_file(f: SourceFile, settings: Settings) -> dict:
    """Extract text for one file (uses the cache). Never raises."""
    key = settings.cache_key()
    if settings.use_cache:
        hit = db.cache_get(f.file_id, key)
        if hit is not None:
            hit["cached"] = True
            return hit
    out = {"pages": [], "warnings": [], "excel": None, "error": "", "cached": False}
    path = Path(f.stored_path)
    try:
        if f.size > settings.max_file_mb * 1024 * 1024:
            raise ValueError(f"File larger than {settings.max_file_mb} MB limit")
        if f.kind == "pdf":
            pages, warn = extract_pdf(path, settings)
        elif f.kind == "image":
            pages, warn = extract_image(path, settings)
        elif f.kind == "word":
            pages, warn = extract_word(path, settings)
        elif f.kind == "excel":
            ex = extract_excel(path)
            pages, warn = ex.pop("pages"), []
            out["excel"] = ex
        else:
            raise ValueError("Unsupported file type")
        out["pages"] = [asdict(p) for p in pages]
        out["warnings"] = warn
        if not any((p.text or "").strip() for p in pages):
            out["error"] = "No readable text could be extracted" + (f" ({'; '.join(warn)})" if warn else "")
    except Exception as e:
        out["error"] = f"{type(e).__name__}: {e}"
        out["trace"] = traceback.format_exc(limit=3)
    if settings.use_cache and not out["error"]:
        db.cache_put(f.file_id, key, out)
    return out


# ------------------------------------------------------------------ documents
def build_documents(f: SourceFile, ext: dict) -> list[Document]:
    pages = [PageText(**p) for p in ext.get("pages", [])]
    f.pages = len(pages)
    docs = []
    excel = ext.get("excel")
    if excel and excel.get("kind") in ("gstr2b", "register"):
        text = "\n".join(p.text for p in pages)
        d = Document(doc_id=_doc_id(f, "all"), file_id=f.file_id, file_name=f.name, rel_path=f.rel_path,
                     pages="all sheets", text=text, text_source="excel", views=[text])
        d.register_rows = excel.get("register_rows", [])
        d.fields["register_kind"] = "GSTR-2B" if excel["kind"] == "gstr2b" else "Invoice register"
        d.fields["register_owner_gstin"] = excel.get("owner_gstin", "")
        d.fields["register_period"] = excel.get("period", "")
        docs.append(d)
        return docs
    if f.kind in ("excel", "word", "image") or len(pages) <= 1:
        segs = [{"pages": [p.page_no for p in pages] or [1], "text_pages": pages, "copies": 1}]
    else:
        segs = segment_pages(pages)
    for s in segs:
        tp = s["text_pages"]
        n_views = max((len(p.views) for p in tp), default=0)
        views = []
        for vi in range(n_views):
            views.append("\n".join(p.views[vi] if vi < len(p.views) else p.text for p in tp))
        text = views[0] if views else ""
        ocr = [p.ocr_confidence for p in tp if p.ocr_confidence is not None]
        src = ",".join(sorted({p.source for p in tp}))
        rng = _page_range(s["pages"]) if f.kind == "pdf" else ("1" if f.kind != "excel" else "sheet(s)")
        d = Document(doc_id=_doc_id(f, rng), file_id=f.file_id, file_name=f.name, rel_path=f.rel_path,
                     pages=rng, text=text, text_source=src, ocr_confidence=(min(ocr) if ocr else None),
                     views=views, copies_in_file=s.get("copies", 1))
        if s.get("copies", 1) > 1:
            d.remarks.append(f"{s['copies']} printed copies (original/duplicate/triplicate) in this file - counted once")
        docs.append(d)
    return docs


def _match_2b(doc: Document, index: dict) -> str:
    f = doc.fields
    key = (f.get("supplier_gstin"), dd.norm_no(f.get("invoice_number")))
    row = index.get(key)
    if not row:
        return ""
    diffs = []
    for k in ("taxable_value", "igst", "cgst", "sgst"):
        a, b = f.get(k), row.get(k)
        if a is not None and b is not None and abs((a or 0) - (b or 0)) > 1:
            diffs.append(f"{k} {a:,.2f} vs 2B {b:,.2f}")
    where = f"{row.get('sheet')} row {row.get('row')} ({row.get('period', '')})"
    return f"Found in GSTR-2B {where}" + (f" - differences: {'; '.join(diffs)}" if diffs else " - values agree")


# ------------------------------------------------------------------ main
def run_job(job_id: str, settings: Settings, progress: ProgressCb = None,
            should_cancel: Optional[Callable[[], bool]] = None) -> dict:
    files = inventory(job_id)
    db.save_files(job_id, files)
    db.save_job(job_id, "Running", settings.as_dict(), total=len(files))
    db.audit(job_id, "job_start", {"files": len(files)})
    todo = [f for f in files if f.status == "Pending"]
    extracted: dict[str, dict] = {}
    done = 0

    def tick(msg):
        if progress:
            progress(done, len(todo), msg)

    ex = ThreadPoolExecutor(max_workers=max(1, settings.max_workers))
    futs = {ex.submit(extract_file, f, settings): f for f in todo}
    try:
        for fut in as_completed(futs):
            f = futs[fut]
            res = fut.result()
            extracted[f.rel_path] = res
            f.status = "Failed" if res.get("error") else ("Processed (cached)" if res.get("cached") else "Processed")
            f.error = res.get("error") or "; ".join(res.get("warnings", []))
            done += 1
            tick(f"{f.rel_path}: {f.status}")
            if should_cancel and should_cancel() and done < len(todo):
                raise Cancelled(f"Cancelled after {done} of {len(todo)} files; completed files are cached - press Resume")
    except BaseException:
        # cancel button, Streamlit rerun (StopException) or crash: drop queued work, keep cache
        ex.shutdown(wait=False, cancel_futures=True)
        db.save_job(job_id, "Interrupted", settings.as_dict(), total=len(files), processed=done)
        db.audit(job_id, "job_interrupted", {"processed": done})
        raise
    ex.shutdown(wait=True)

    # ---- documents + classification -----------------------------------------------
    docs: list[Document] = []
    for f in files:
        res = extracted.get(f.rel_path)
        if not res or res.get("error"):
            continue
        for d in build_documents(f, res):
            is_reg = d.fields.get("register_kind") is not None
            d.doc_type, d.type_confidence, d.type_reason = classify(d.views, f.name, is_excel_register=is_reg)
            docs.append(d)
    db.save_files(job_id, files)

    # ---- Caliber master (evidence based, never overwrites user edits) ---------------
    master = CaliberMaster(settings)
    owners = {d.fields["register_owner_gstin"]: f"{d.rel_path} (GSTR-2B owner)"
              for d in docs if d.fields.get("register_owner_gstin")}
    proposals = master.detect_from_documents(docs, owners)
    added = master.apply_proposals(proposals)
    if added:
        db.audit(job_id, "master_detected", {"added": added})

    # ---- fields --------------------------------------------------------------------------
    for d in docs:
        if d.doc_type == EXCEL_REGISTER:
            d.extraction_confidence = 1.0
            continue
        fields, conf, breakdown = extract_fields(d.views, d.doc_type, master, settings)
        d.fields.update(fields)
        d.field_confidence = conf
        d.line_items = breakdown
        d.direction = fields.get("direction", d.direction)
        d.extraction_confidence = extraction_confidence(conf, d.ocr_confidence, fields.get("amounts_balanced", False))

    # second pass: master proposals that need bill-to names (non-PAN registrations)
    more = master.apply_proposals(master.detect_from_documents(docs, owners))
    if more:
        db.audit(job_id, "master_detected", {"added": more})

    dup_rows = finalize(job_id, docs, master, settings)
    db.save_documents(job_id, docs)
    failed = [f for f in files if f.status in ("Failed", "Skipped")]
    db.save_job(job_id, "Completed", settings.as_dict(), total=len(files), processed=len(files) - len(failed),
                note=f"{len(docs)} documents")
    db.audit(job_id, "job_complete", {"documents": len(docs), "failed_or_skipped": len(failed)})
    return {"job_id": job_id, "files": files, "documents": docs, "duplicates": dup_rows, "master": master}


def finalize(job_id: str, docs: list[Document], master: CaliberMaster, settings: Settings) -> list[dict]:
    """Overrides -> validation -> GSTR-2B match -> duplicates -> status.

    A pure function of the extracted documents plus stored review decisions,
    so it is re-run after every correction without touching the source files."""
    for d in docs:
        d.remarks, d.validations = [], []
        d.duplicate_group = d.duplicate_role = ""
        d.caliber_gstin = d.caliber_state = ""
        if d.copies_in_file > 1:
            d.remarks.append(f"{d.copies_in_file} printed copies (original/duplicate/triplicate) in this file - counted once")
        if d.fields.get("amounts_derived"):
            d.remarks.append(f"Derived value: {d.fields['amounts_derived']} (not printed) - verify")
    rm.apply_overrides(docs, rm.load_overrides(job_id))

    for d in docs:
        cg = d.fields.get("customer_gstin") if d.direction == INWARD else d.fields.get("supplier_gstin") if d.direction == OUTWARD else ""
        if cg and master.is_caliber_gstin(cg):
            d.caliber_gstin, d.caliber_state = cg, state_name(cg)
        validate_document(d, master, settings)

    index = {}
    for d in docs:
        for r in d.register_rows:
            index[(r.get("supplier_gstin"), dd.norm_no(r.get("invoice_number")))] = r
    for d in docs:
        d.gstr2b_match = ""
        if d.doc_type in INVOICE_TYPES | NOTE_TYPES and d.direction == INWARD:
            d.gstr2b_match = _match_2b(d, index) or ("Not in uploaded GSTR-2B files" if index else "")

    dd.link_notes_and_revisions(docs)
    dup_rows = dd.detect(docs, rm.load_duplicate_decisions(job_id), settings.fuzzy_duplicate_threshold)

    for d in docs:
        d.status = decide_status(d, settings)
        if d.duplicate_role in ("Duplicate", "Revised/conflicting version") and d.doc_type != EXCEL_REGISTER:
            d.status = ST_DUPLICATE
        if d.fields.get("_approved") and d.status == ST_REVIEW:
            if not any(v["level"] == "ERROR" for v in d.validations):
                d.status = ST_VERIFIED
                d.remarks.append("Verified by reviewer")
            else:
                d.remarks.append("Reviewer approval recorded but ERROR-level issues remain - correct the fields first")
        for v in d.validations:
            if v["level"] in ("ERROR", "WARNING"):
                d.remarks.append(f"[{v['level']}] {v['message']}")
    return dup_rows


def revalidate(job_id: str, settings: Settings) -> dict:
    """Re-apply corrections, duplicate decisions and master changes to a finished job (no re-extraction)."""
    files = [SourceFile(**{k: v for k, v in r.items() if k != "job_id"}) for r in db.load_files(job_id)]
    docs = [Document.from_dict(d) for d in db.load_documents(job_id)]
    master = CaliberMaster(settings)
    dup_rows = finalize(job_id, docs, master, settings)
    db.save_documents(job_id, docs)
    return {"job_id": job_id, "files": files, "documents": docs, "duplicates": dup_rows, "master": master}


def load_job(job_id: str, settings: Settings) -> dict:
    """Open a finished job (re-applies the latest master and review decisions)."""
    return revalidate(job_id, settings)

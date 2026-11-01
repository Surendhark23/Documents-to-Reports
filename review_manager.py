"""Manual review: field corrections, classification overrides, duplicate
decisions and approval.  Every change is written to the ``corrections`` /
``duplicate_decisions`` tables with user and timestamp (audit trail) and is
re-applied whenever a job is re-processed, so review work is never lost."""
from __future__ import annotations

import json
from typing import Any

import database as db
from field_extractor import parse_amount, parse_date
from models import AMOUNT_FIELDS, DOC_TYPES, Document

EDITABLE_FIELDS = [
    "doc_type", "direction", "invoice_number", "invoice_date", "original_invoice_number",
    "supplier_legal_name", "supplier_trade_name", "supplier_gstin", "supplier_address",
    "customer_legal_name", "customer_gstin", "customer_address", "bill_to_name", "bill_to_gstin",
    "ship_to_name", "ship_to_gstin", "place_of_supply", "reverse_charge", "currency",
    *AMOUNT_FIELDS, "po_number", "irn", "eway_bill", "review_status",
]
APPROVE = "__approved__"


def record_correction(job_id: str, doc: Document, field: str, new_value: Any, user: str = "reviewer") -> None:
    old = doc.doc_type if field == "doc_type" else doc.fields.get(field)
    with db.connect() as con:
        con.execute("INSERT INTO corrections(job_id, doc_id, rel_path, field, old_value, new_value, user, at) "
                    "VALUES (?,?,?,?,?,?,?,?)",
                    (job_id, doc.doc_id, doc.rel_path, field, json.dumps(old, default=str),
                     json.dumps(new_value, default=str), user, db.now()))


def approve(job_id: str, doc: Document, user: str = "reviewer", note: str = "") -> None:
    record_correction(job_id, doc, APPROVE, note or "approved", user)


def load_overrides(job_id: str) -> dict[str, dict]:
    """{doc_id: {field: latest value}} - later corrections win."""
    out: dict[str, dict] = {}
    with db.connect() as con:
        for r in con.execute("SELECT doc_id, field, new_value FROM corrections WHERE job_id=? ORDER BY id", (job_id,)):
            out.setdefault(r["doc_id"], {})[r["field"]] = json.loads(r["new_value"])
    return out


def correction_history(job_id: str, doc_id: str | None = None) -> list[dict]:
    q = "SELECT * FROM corrections WHERE job_id=?" + (" AND doc_id=?" if doc_id else "") + " ORDER BY id"
    with db.connect() as con:
        return [dict(r) for r in con.execute(q, (job_id, doc_id) if doc_id else (job_id,))]


def coerce(field: str, value: Any) -> Any:
    if value is None or value == "":
        return None if field in AMOUNT_FIELDS else ""
    if field in AMOUNT_FIELDS:
        return parse_amount(str(value)) if not isinstance(value, (int, float)) else float(value)
    if field == "invoice_date":
        iso, _ = parse_date(str(value))
        return iso or str(value)
    if field in ("supplier_gstin", "customer_gstin", "bill_to_gstin", "ship_to_gstin"):
        return str(value).strip().upper()
    if field == "doc_type" and value not in DOC_TYPES:
        raise ValueError(f"Unknown document type {value}")
    return value


def apply_overrides(docs: list[Document], overrides: dict[str, dict]) -> list[Document]:
    """Apply stored corrections; returns the documents that were changed."""
    changed = []
    for d in docs:
        ov = overrides.get(d.doc_id)
        if not ov:
            continue
        for field, value in ov.items():
            if field == APPROVE:
                d.fields["_approved"] = value
                continue
            if field == "doc_type":
                d.doc_type, d.type_confidence = value, 1.0
                d.type_reason = f"Set manually by reviewer (was: {d.type_reason})"
            elif field == "direction":
                d.direction = value
            else:
                d.fields[field] = value
                d.field_confidence[field] = 1.0
        d.remarks.append("Manually corrected: " + ", ".join(k for k in ov if k != APPROVE)
                         if any(k != APPROVE for k in ov) else "Approved by reviewer")
        changed.append(d)
    return changed


def save_duplicate_decision(job_id: str, primary_doc: str, candidate_doc: str, decision: str, user: str = "reviewer") -> None:
    with db.connect() as con:
        con.execute("INSERT OR REPLACE INTO duplicate_decisions VALUES (?,?,?,?,?,?)",
                    (job_id, primary_doc, candidate_doc, decision, user, db.now()))
    db.audit(job_id, "duplicate_decision", {"primary": primary_doc, "candidate": candidate_doc, "decision": decision, "by": user})


def load_duplicate_decisions(job_id: str) -> dict[tuple[str, str], str]:
    with db.connect() as con:
        return {(r["primary_doc"], r["candidate_doc"]): r["decision"]
                for r in con.execute("SELECT * FROM duplicate_decisions WHERE job_id=?", (job_id,))}

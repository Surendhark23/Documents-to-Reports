"""Four-level duplicate and revised-document detection.

Level 1  identical file bytes (SHA-256)
Level 2  same supplier GSTIN + same document number (+ same document class)
Level 3  level 2 + same date, taxable value and tax amounts  -> data duplicate
         level 2 with differing values                        -> revised version / conflicting copy
Level 4  fuzzy text similarity + same total when keys are missing

Nothing is deleted.  Each candidate is reported with the primary document,
match reason, confidence and a suggested action; the user's decision
('Exact duplicate', 'Revised invoice', 'Separate transaction') is stored and
re-applied on every run.  Documents of different types with matching values
(e.g. a quotation and the later invoice) are reported as *related*, never
merged.
"""
from __future__ import annotations

import re
from collections import defaultdict

from rapidfuzz import fuzz

from models import (Document, INVOICE_TYPES, NOTE_TYPES, EXCEL_REGISTER, CREDIT_NOTE, DEBIT_NOTE)

DECISIONS = ["Exact duplicate", "Revised invoice", "Separate transaction"]


def norm_no(no: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", (no or "").upper())


def _doc_class(t: str) -> str:
    if t in INVOICE_TYPES:
        return "INV"
    if t == CREDIT_NOTE:
        return "CN"
    if t == DEBIT_NOTE:
        return "DN"
    return t


def _rank(d: Document, mtimes: dict) -> tuple:
    """Lower is better for choosing the primary document of a group."""
    name = d.file_name.lower()
    copy_marker = bool(re.search(r"\(\d+\)|copy|\.\.\.|111|duplicate", name))
    return (0 if d.doc_type in INVOICE_TYPES | NOTE_TYPES else 1,
            -round(d.extraction_confidence, 2),
            0 if d.fields.get("irn") else 1,
            1 if copy_marker else 0,
            d.rel_path.count("/"),
            d.rel_path)


def _values(d: Document) -> tuple:
    f = d.fields
    return (f.get("invoice_date"), f.get("taxable_value"), f.get("igst"), f.get("cgst"), f.get("sgst"),
            f.get("invoice_total"))


def _diff_fields(a: Document, b: Document) -> list[str]:
    out = []
    for k in ("invoice_date", "taxable_value", "igst", "cgst", "sgst", "invoice_total", "customer_gstin"):
        va, vb = a.fields.get(k), b.fields.get(k)
        if va is None or vb is None:
            continue
        if isinstance(va, float) and isinstance(vb, float):
            if abs(va - vb) > 0.5:
                out.append(f"{k}: {va:,.2f} vs {vb:,.2f}")
        elif va != vb:
            out.append(f"{k}: {va} vs {vb}")
    return out


def _norm_text(t: str) -> str:
    t = re.sub(r"\s+", " ", (t or "").lower())
    return t[:4000]


def detect(docs: list[Document], decisions: dict | None = None, threshold: int = 95,
           mtimes: dict | None = None) -> list[dict]:
    """Annotate docs in place and return report rows."""
    decisions = decisions or {}
    mtimes = mtimes or {}
    rows: list[dict] = []
    grouped: set[str] = set()
    gid = 0

    def add_group(members: list[Document], level: str, reason: str, conf: float, action: str,
                  role: str = "Duplicate") -> None:
        nonlocal gid
        members = sorted(members, key=lambda d: _rank(d, mtimes))
        already = [m for m in members if m.doc_id in grouped]
        if already:
            # v1.1.1 fix E16: a member already belongs to an earlier group (e.g. an L1 identical-file
            # group).  Join that group instead of opening a new one - before, the new 'primary' had no
            # duplicates listed, both copies stayed Verified and SC 0290 was counted twice in Final.
            primary = next((m for m in already if m.duplicate_role == "Primary"), already[0])
            g = primary.duplicate_group
            if not any(m.doc_id not in grouped for m in members):
                return
        else:
            primary = members[0]
            gid += 1
            g = f"D{gid:03d}"
            primary.duplicate_group, primary.duplicate_role = g, "Primary"
        for m in [x for x in members if x is not primary]:
            if m.doc_id in grouped:
                continue
            dec = decisions.get((primary.doc_id, m.doc_id), "")
            diffs = _diff_fields(primary, m)
            rows.append({
                "Group": g, "Level": level, "Primary Document": f"{primary.rel_path} [p{primary.pages}]",
                "Duplicate Candidate": f"{m.rel_path} [p{m.pages}]",
                "Primary Doc ID": primary.doc_id, "Candidate Doc ID": m.doc_id,
                "Document Type": m.doc_type, "Supplier GSTIN": m.fields.get("supplier_gstin", ""),
                "Document Number": m.fields.get("invoice_number", ""),
                "Match Reason": reason, "Differences": "; ".join(diffs),
                "Match Confidence": conf, "Suggested Action": action, "User Decision": dec,
            })
            m.duplicate_group = g
            if dec == "Separate transaction":
                m.duplicate_role = "Separate (user)"
            elif dec == "Revised invoice":
                m.duplicate_role = "Revised version (user)"
            else:
                m.duplicate_role = role
            grouped.add(m.doc_id)
        grouped.add(primary.doc_id)

    # ---- Level 1: identical bytes -----------------------------------------------
    by_hash = defaultdict(list)
    for d in docs:
        by_hash[(d.file_id, d.pages)].append(d)
    for (fid, pages), members in by_hash.items():
        paths = {m.rel_path for m in members}
        if len(paths) > 1:
            add_group(members, "L1 - identical file",
                      "Byte-identical file stored under different names/folders", 1.0,
                      "Keep one copy; exclude the others from totals")

    # ---- Level 2/3: supplier GSTIN + number ---------------------------------------
    money = [d for d in docs if d.doc_type in INVOICE_TYPES | NOTE_TYPES]
    by_key = defaultdict(list)
    for d in money:
        no, sup = norm_no(d.fields.get("invoice_number")), d.fields.get("supplier_gstin")
        if no and sup:
            by_key[(sup, no, _doc_class(d.doc_type))].append(d)
    for key, members in by_key.items():
        # collapse L1 copies to one representative
        reps, seen = [], set()
        for m in sorted(members, key=lambda d: _rank(d, mtimes)):
            if m.file_id in seen:
                continue
            seen.add(m.file_id)
            reps.append(m)
        if len(reps) < 2:
            continue
        reps = sorted(reps, key=lambda d: _rank(d, mtimes))
        base = reps[0]
        same = [base] + [m for m in reps[1:] if _values(m) == _values(base)]
        diff = [m for m in reps[1:] if _values(m) != _values(base)]
        if len(same) > 1:
            add_group(same, "L3 - same supplier, number, date and values",
                      "Same supplier GSTIN, document number, date, taxable value and taxes (different files)", 0.97,
                      "Exact duplicate - keep primary only")
        for m in diff:
            diffs = _diff_fields(base, m)
            if not diffs:   # values missing on one side -> level 2 only
                add_group([base, m], "L2 - same supplier and number",
                          "Same supplier GSTIN and document number; values could not be fully compared", 0.85,
                          "Probable duplicate - confirm")
            else:
                add_group([base, m], "L2 - same number, different values",
                          "Same supplier GSTIN and document number but values differ -> revised version or conflicting copy",
                          0.8, "Review: decide which version is final (Revised invoice) or Separate transaction",
                          role="Revised/conflicting version")

    # ---- Level 4: fuzzy -------------------------------------------------------------
    pool = [d for d in docs if d.doc_type != EXCEL_REGISTER and d.doc_id not in grouped and len(d.text) > 200]
    norm = {d.doc_id: _norm_text(d.text) for d in pool}
    for i, a in enumerate(pool):
        if a.doc_id in grouped:
            continue
        for b in pool[i + 1:]:
            if b.doc_id in grouped or a.file_id == b.file_id:
                continue
            ta, tb = a.fields.get("invoice_total"), b.fields.get("invoice_total")
            if ta is None or tb is None or abs(ta - tb) > 1:
                continue
            # audit fix E9: fuzzy matching is only for documents whose keys are missing.  Two
            # documents that both print a number (or a date) and differ are separate invoices from
            # the same template (Noori Store 26-27/054 of 9-Jul vs 26-27/077 of 14-Aug).
            na, nb = norm_no(a.fields.get("invoice_number", "")), norm_no(b.fields.get("invoice_number", ""))
            da, db_ = a.fields.get("invoice_date", ""), b.fields.get("invoice_date", "")
            if (na and nb and na != nb) or (da and db_ and da != db_):
                continue
            score = fuzz.token_set_ratio(norm[a.doc_id], norm[b.doc_id])
            if score < threshold:
                continue
            if _doc_class(a.doc_type) != _doc_class(b.doc_type):
                rows.append({
                    "Group": "", "Level": "Related (not duplicate)",
                    "Primary Document": f"{a.rel_path} [p{a.pages}]", "Duplicate Candidate": f"{b.rel_path} [p{b.pages}]",
                    "Primary Doc ID": a.doc_id, "Candidate Doc ID": b.doc_id, "Document Type": f"{a.doc_type} / {b.doc_type}",
                    "Supplier GSTIN": a.fields.get("supplier_gstin", ""), "Document Number": "",
                    "Match Reason": f"Same values and {score}% similar text but different document types",
                    "Differences": "", "Match Confidence": round(score / 100, 2),
                    "Suggested Action": "Keep both - e.g. quotation followed by invoice", "User Decision": ""})
                continue
            add_group([a, b], "L4 - fuzzy text match", f"{score}% text similarity and identical total", round(score / 100 * 0.9, 2),
                      "Probable duplicate - confirm")
    return rows


def link_notes_and_revisions(docs: list[Document]) -> None:
    """Attach references between credit/debit notes, revised invoices and originals."""
    by_no = defaultdict(list)
    for d in docs:
        if d.doc_type in INVOICE_TYPES and d.fields.get("invoice_number"):
            by_no[(d.fields.get("supplier_gstin"), norm_no(d.fields["invoice_number"]))].append(d)
    for d in docs:
        ref = d.fields.get("original_invoice_number")
        if d.doc_type in NOTE_TYPES and ref:
            hits = by_no.get((d.fields.get("supplier_gstin"), norm_no(ref)), [])
            if hits:
                d.remarks.append(f"Original invoice {ref} found in batch: {hits[0].rel_path} [p{hits[0].pages}]")
                for h in hits:
                    h.remarks.append(f"{d.doc_type} {d.fields.get('invoice_number')} issued against this invoice ({d.rel_path})")
            else:
                d.remarks.append(f"Original invoice {ref} is not in this batch")

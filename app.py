"""Caliber GST Document Processor - Streamlit entry point.

    streamlit run app.py

Pages: Dashboard | Upload & Processing | Caliber GST Master | Document Review |
Duplicate Review | Reports & Export.  All processing is local (PyMuPDF,
Tesseract, OpenCV); no document leaves the machine.
"""
from __future__ import annotations

import io
import json
import os
from collections import Counter
from pathlib import Path

import pandas as pd
import streamlit as st

import database as db
import review_manager as rm
from caliber_master import CaliberMaster, MASTER_COLUMNS, VERIFIED, DETECTED, REJECTED
from config import APP_NAME, default_settings, data_dir, SUPPORTED_ALL, SUPPORTED_ARCHIVES
from file_manager import new_job_id, ingest_uploads, ingest_folder, UploadError
from gst_validator import gstin_valid, state_name
from models import (DOC_TYPES, INVOICE_TYPES, NOTE_TYPES, INWARD, OUTWARD, THIRD_PARTY, CONSIGNEE_ONLY,
                    DIRECTION_UNKNOWN, ST_VERIFIED, ST_REVIEW, ST_DUPLICATE, ST_NON_TAX, ST_FAILED, ST_OUTWARD)
from ocr_engine import tesseract_available
from pipeline import run_job, revalidate, Cancelled
from report_generator import build_workbook, invoice_row, supplier_master_rows

st.set_page_config(page_title=APP_NAME, page_icon="📑", layout="wide")

CLOUD = os.environ.get("CALIBER_CLOUD", "") == "1"
STATUS_BADGE = {ST_VERIFIED: "🟢", ST_REVIEW: "🟠", ST_DUPLICATE: "🟣", ST_NON_TAX: "⚪", ST_FAILED: "🔴",
                ST_OUTWARD: "🔵"}

st.markdown("""
<style>
  .block-container {padding-top: 1.4rem;}
  div[data-testid="stMetric"] {background: rgba(31,58,95,0.06); border: 1px solid rgba(31,58,95,0.15);
      border-radius: 8px; padding: 10px 14px;}
  .small-note {color: #6b7280; font-size: 0.85rem;}
</style>""", unsafe_allow_html=True)


# ------------------------------------------------------------------ state helpers
def settings():
    if "settings" not in st.session_state:
        st.session_state.settings = default_settings()
    return st.session_state.settings


def current():
    """The loaded job result (re-built from the database when needed)."""
    jid = st.session_state.get("job_id")
    if not jid:
        return None
    if st.session_state.get("res") is None or st.session_state.res["job_id"] != jid:
        if not db.load_documents(jid) and not db.load_files(jid):
            return None
        st.session_state.res = revalidate(jid, settings())
    return st.session_state.res


def refresh():
    jid = st.session_state.get("job_id")
    if jid:
        st.session_state.res = revalidate(jid, settings())
    # v1.1.1: a workbook built before a master / review change is stale - force a rebuild
    st.session_state.pop("wb", None)


def df_from(rows, cols=None):
    df = pd.DataFrame(rows)
    if cols:
        df = df[[c for c in cols if c in df.columns]]
    return df


# ------------------------------------------------------------------ sidebar
def sidebar():
    s = settings()
    st.sidebar.title("📑 Caliber GST")
    st.sidebar.caption("Supplier-invoice identification & state-wise GST register")
    page = st.sidebar.radio("Page", ["Dashboard", "Upload & Processing", "Caliber GST Master", "Document Review",
                                     "Duplicate Review", "Reports & Export"], key="page")
    jobs = db.list_jobs()
    if jobs:
        labels = {j["job_id"]: f"{j['job_id']} · {j['status']} · {j.get('note') or ''}" for j in jobs}
        ids = list(labels)
        cur = st.session_state.get("job_id")
        idx = ids.index(cur) if cur in ids else 0
        sel = st.sidebar.selectbox("Batch (job)", ids, index=idx, format_func=lambda x: labels[x])
        if sel != cur:
            st.session_state.job_id, st.session_state.res = sel, None
    with st.sidebar.expander("Processing settings"):
        s.ocr_enabled = st.checkbox("OCR for scans/photos (local Tesseract)", s.ocr_enabled,
                                    help="Embedded PDF text is always preferred; OCR runs only when needed.")
        if s.ocr_enabled and not tesseract_available():
            st.warning("Tesseract is not installed - scanned pages and images will be listed under Processing_Errors.")
        s.max_workers = st.slider("Parallel workers", 1, 8, s.max_workers)
        s.max_file_mb = st.number_input("Max file size (MB)", 1, 500, s.max_file_mb)
        s.max_pdf_pages = st.number_input("Max pages per PDF", 1, 500, s.max_pdf_pages)
        s.ocr_timeout_s = st.number_input("OCR timeout per page (s)", 10, 600, s.ocr_timeout_s)
        s.amount_tolerance = st.number_input("Rounding tolerance (₹)", 0.0, 10.0, float(s.amount_tolerance), 0.5)
        s.min_confidence_for_final = st.slider("Min. confidence for auto-verify", 0.5, 1.0, float(s.min_confidence_for_final), 0.05)
        s.include_outward_in_final = st.checkbox("Include Caliber's own outward invoices in Final_Tax_Invoices",
                                                 s.include_outward_in_final)
    st.sidebar.markdown("<span class='small-note'>🔒 Local processing only. External AI is disabled "
                        "(see README to enable with approval).</span>", unsafe_allow_html=True)
    return page


# ------------------------------------------------------------------ dashboard
def page_dashboard():
    st.header("Dashboard")
    res = current()
    if not res:
        st.info("No batch loaded yet. Go to **Upload & Processing** to process a ZIP or files.")
        return
    files, docs = res["files"], res["documents"]
    c = Counter(d.status for d in docs)
    ti = [d for d in docs if d.doc_type in INVOICE_TYPES and d.direction == INWARD]
    cols = st.columns(4)
    cols[0].metric("Files received", len(files))
    cols[1].metric("Files processed", sum(f.status.startswith("Processed") for f in files))
    cols[2].metric("Inward tax invoices", len(ti), help="Supplier invoices billed to a Caliber GSTIN (incl. duplicates)")
    cols[3].metric("Verified (final)", c.get(ST_VERIFIED, 0))
    cols = st.columns(4)
    cols[0].metric("Non-tax documents", c.get(ST_NON_TAX, 0))
    cols[1].metric("Credit / debit notes", sum(d.doc_type in NOTE_TYPES for d in docs))
    cols[2].metric("Duplicate candidates", c.get(ST_DUPLICATE, 0))
    cols[3].metric("Pending review", c.get(ST_REVIEW, 0))
    cols = st.columns(4)
    cols[0].metric("Processing errors", sum(f.status == "Failed" for f in files))
    cols[1].metric("Skipped files", sum(f.status == "Skipped" for f in files))
    cols[2].metric("Caliber outward docs", c.get(ST_OUTWARD, 0))
    cols[3].metric("Documents found", len(docs))

    final = [d for d in docs if d.status == ST_VERIFIED and d.doc_type in INVOICE_TYPES and d.direction == INWARD]
    st.subheader("Final register by Caliber registration")
    if final:
        rows = {}
        for d in final:
            k = f"{d.caliber_gstin} ({d.caliber_state})"
            r = rows.setdefault(k, {"Registration": k, "Invoices": 0, "Taxable": 0.0, "IGST": 0.0, "CGST": 0.0, "SGST": 0.0, "Total": 0.0})
            r["Invoices"] += 1
            for col, key in (("Taxable", "taxable_value"), ("IGST", "igst"), ("CGST", "cgst"), ("SGST", "sgst"), ("Total", "invoice_total")):
                r[col] += d.fields.get(key) or 0
        st.dataframe(pd.DataFrame(rows.values()).style.format({c: "{:,.2f}" for c in ("Taxable", "IGST", "CGST", "SGST", "Total")}),
                     hide_index=True, width="stretch")
    else:
        st.caption("No verified invoices yet - verify the Caliber GST master first, then review pending documents.")
    unverified = [e for e in res["master"].entries.values() if e.get("verification_status") == DETECTED]
    if unverified:
        st.warning(f"{len(unverified)} Caliber GST registration(s) were detected but not verified yet. "
                   "Invoices billed to them stay in Pending_Review until you verify them on **Caliber GST Master**.")
    st.subheader("Documents by type and status")
    t = Counter((d.doc_type, d.status) for d in docs)
    st.dataframe(pd.DataFrame([{"Document type": a, "Status": b, "Count": n} for (a, b), n in sorted(t.items())]),
                 hide_index=True, width="stretch")


# ------------------------------------------------------------------ upload
def page_upload():
    st.header("Upload & Processing")
    s = settings()
    st.markdown("Upload a **ZIP** of the month's documents, or individual PDFs, images, Excel/CSV and Word files. "
                "Sub-folders inside a ZIP are processed recursively. Original files are never modified.")
    exts = sorted(e.lstrip(".") for e in SUPPORTED_ALL | SUPPORTED_ARCHIVES)
    up = st.file_uploader("Drag and drop files or a ZIP", type=exts, accept_multiple_files=True)
    folder = ""
    if not CLOUD:
        with st.expander("Or process a folder on this computer (local installation only)"):
            folder = st.text_input("Folder path", placeholder=r"e.g. D:\Caliber\Jun-26  or  /home/user/Caliber/Jun-26")
            st.caption("Browsers cannot hand a whole folder to a web app on every platform; when the app runs on a "
                       "server or Streamlit Cloud, zip the folder and upload the ZIP instead.")
    else:
        st.caption("Folder selection is not available in the hosted version - zip the folder and upload the ZIP.")
    c1, c2, c3 = st.columns([1, 1, 2])
    start = c1.button("▶ Process", type="primary", disabled=not (up or folder))
    jid = st.session_state.get("job_id")
    job = next((j for j in db.list_jobs() if j["job_id"] == jid), None)
    resume = c2.button("⟳ Resume / re-run selected batch", disabled=not job,
                       help="Re-runs the selected batch. Completed files are taken from the cache, so only unfinished work is redone.")
    if job and job["status"] in ("Running", "Interrupted"):
        st.warning(f"Batch {jid} was interrupted after {job['processed_files']} of {job['total_files']} files. Press Resume.")

    if start:
        jid = new_job_id()
        problems = []
        try:
            if up:
                problems += ingest_uploads(jid, [(f.name, f.getvalue()) for f in up], s)
            if folder:
                problems += ingest_folder(jid, Path(folder), s)
        except UploadError as e:
            st.error(str(e))
            return
        (data_dir() / "jobs" / jid / "upload_problems.json").write_text(json.dumps(problems))
        st.session_state.job_id = jid
        _run(jid, problems)
    elif resume and job:
        _run(jid, _problems(jid))
    res = current()
    if res:
        st.subheader("Processing log")
        log = pd.DataFrame([{"File": f.rel_path, "Type": f.kind, "Size KB": round(f.size / 1024, 1), "Pages": f.pages,
                             "Status": f.status, "Message": f.error} for f in res["files"]])
        st.dataframe(log, hide_index=True, width="stretch", height=380)
        probs = _problems(res["job_id"])
        if probs:
            st.subheader("Rejected during upload")
            st.dataframe(pd.DataFrame(probs), hide_index=True, width="stretch")


def _problems(jid):
    p = data_dir() / "jobs" / jid / "upload_problems.json"
    return json.loads(p.read_text()) if p.exists() else []


def _run(jid, problems):
    s = settings()
    cancel_flag = data_dir() / "jobs" / jid / "CANCEL"
    cancel_flag.unlink(missing_ok=True)
    bar = st.progress(0.0, text="Starting...")
    log = st.empty()
    lines = []
    if st.button("■ Cancel safely", key=f"cancel-{jid}"):
        cancel_flag.touch()

    def prog(done, total, msg):
        bar.progress(done / max(total, 1), text=f"{done}/{total} files")
        lines.append(msg)
        log.code("\n".join(lines[-12:]))

    try:
        with st.spinner("Extracting, classifying and validating..."):
            st.session_state.res = run_job(jid, s, prog, should_cancel=lambda: cancel_flag.exists())
        bar.progress(1.0, text="Done")
        res = st.session_state.res
        st.success(f"Processed {len(res['files'])} files -> {len(res['documents'])} documents. "
                   "Next: verify the Caliber GST master, then work through Document Review.")
        if problems:
            st.warning(f"{len(problems)} item(s) were rejected during upload - see below and Processing_Errors.")
    except Cancelled as e:
        st.warning(str(e))


# ------------------------------------------------------------------ master
def page_master():
    st.header("Caliber GST Master")
    s = settings()
    master = CaliberMaster(s)
    st.markdown("State-wise registrations of **Caliber Inno Enterprise Pvt Ltd**. Entries are proposed only from GSTINs "
                "found on the documents (with the documents as evidence) - verify, correct or reject them here. "
                "The state code is always derived from the first two digits of the GSTIN.")
    df = master.to_dataframe()
    if df.empty:
        df = pd.DataFrame(columns=MASTER_COLUMNS)
    edited = st.data_editor(
        df, num_rows="dynamic", width="stretch", hide_index=True, key="master_editor",
        column_config={
            "Verification Status": st.column_config.SelectboxColumn(options=[VERIFIED, DETECTED, REJECTED]),
            "Active/Inactive": st.column_config.SelectboxColumn(options=["Active", "Inactive"]),
            "State Code": st.column_config.TextColumn(disabled=True), "State Name": st.column_config.TextColumn(disabled=True),
        })
    c1, c2, c3 = st.columns(3)
    if c1.button("💾 Save master", type="primary"):
        errs, keep = [], set()
        col_to_db = {"Customer Legal Name": "legal_name", "Customer Trade Name": "trade_name", "Registered Address": "address",
                     "Effective From": "effective_from", "Effective To": "effective_to",
                     "Verification Status": "verification_status", "Source Document": "source_document", "Remarks": "remarks"}
        for _, r in edited.iterrows():
            g = str(r.get("GSTIN") or "").strip().upper()
            if not g:
                continue
            if not gstin_valid(g):
                errs.append(f"{g}: invalid GSTIN (format/checksum) - not saved")
                continue
            e = {k: str(r.get(c) or "") for c, k in col_to_db.items()}
            e.update(gstin=g, active=str(r.get("Active/Inactive", "Active")) != "Inactive", state_code=g[:2],
                     state_name=state_name(g))
            master.upsert(e, user="ui")
            keep.add(g)
        for g in set(master.entries) - keep:
            master.delete(g)
        for e in errs:
            st.error(e)
        refresh()
        st.success("Master saved; open batches were re-validated.")
    if c2.button("✔ Mark all detected as Verified"):
        for g, e in master.entries.items():
            if e.get("verification_status") == DETECTED:
                master.upsert({"gstin": g, "verification_status": VERIFIED}, user="ui")
        refresh()
        st.rerun()
    buf = io.BytesIO()
    master.to_dataframe().to_excel(buf, index=False)
    c3.download_button("⬇ Export master (Excel)", buf.getvalue(), "Caliber_GST_Master.xlsx")
    imp = st.file_uploader("Import master from Excel (columns as above; GSTIN required)", type=["xlsx", "xls"])
    if imp and st.button("Import"):
        n, errs = master.import_excel(imp)
        st.success(f"Imported {n} registration(s).")
        for e in errs:
            st.warning(e)
        refresh()


# ------------------------------------------------------------------ document review
def _preview(d, res):
    f = next((x for x in res["files"] if x.rel_path == d.rel_path), None)
    if not f or not Path(f.stored_path).exists():
        st.caption("Original file not available.")
        return
    path = Path(f.stored_path)
    st.download_button("⬇ Original file", path.read_bytes(), path.name, key=f"dl-{d.doc_id}")
    if f.kind == "pdf":
        import pymupdf
        pages = []
        for part in str(d.pages).split(","):
            if "-" in part:
                a, b = part.split("-")
                pages += list(range(int(a), int(b) + 1))
            elif part.strip().isdigit():
                pages.append(int(part))
        doc = pymupdf.open(path)
        for n in pages[:4]:
            if 1 <= n <= doc.page_count:
                st.image(doc[n - 1].get_pixmap(dpi=110).tobytes("png"), caption=f"{path.name} - page {n}",
                         width="stretch")
    elif f.kind == "image":
        st.image(str(path), width="stretch")
    elif f.kind == "excel":
        try:
            st.dataframe(pd.read_excel(path, header=None, nrows=60) if path.suffix != ".csv" else pd.read_csv(path, nrows=60))
        except Exception as e:
            st.caption(f"Preview not available: {e}")
    with st.expander("Extracted text"):
        st.text(d.text[:8000])


def page_review():
    st.header("Document Review")
    res = current()
    if not res:
        st.info("Process a batch first.")
        return
    docs = res["documents"]
    c1, c2, c3 = st.columns(3)
    sts = c1.multiselect("Status", sorted({d.status for d in docs}), default=[ST_REVIEW] if any(d.status == ST_REVIEW for d in docs) else [])
    tys = c2.multiselect("Document type", sorted({d.doc_type for d in docs}))
    q = c3.text_input("Search (file, supplier, number)")
    sel = [d for d in docs if (not sts or d.status in sts) and (not tys or d.doc_type in tys) and
           (not q or q.lower() in f"{d.rel_path} {d.fields.get('supplier_legal_name', '')} {d.fields.get('invoice_number', '')}".lower())]
    st.caption(f"{len(sel)} document(s)")
    table = pd.DataFrame([{"": STATUS_BADGE.get(d.status, ""), "Status": d.status, "Type": d.doc_type,
                           "File": d.rel_path, "Pages": d.pages, "Number": d.fields.get("invoice_number", ""),
                           "Supplier": d.fields.get("supplier_legal_name", ""), "Total": d.fields.get("invoice_total"),
                           "Conf.": d.extraction_confidence,
                           "Issues": sum(v["level"] in ("ERROR", "WARNING") for v in d.validations)} for d in sel])
    if table.empty:
        return
    ev = st.dataframe(table, hide_index=True, width="stretch", height=260, on_select="rerun",
                      selection_mode="single-row", key="review_table")
    rows = ev.selection.rows if ev and hasattr(ev, "selection") else []
    d = sel[rows[0]] if rows else sel[0]
    st.divider()
    st.subheader(f"{STATUS_BADGE.get(d.status, '')} {d.file_name}  ·  pages {d.pages}")
    st.caption(f"{d.doc_type} (confidence {d.type_confidence:.0%}) - {d.type_reason}")
    left, right = st.columns([1, 1])
    with left:
        _preview(d, res)
    with right:
        for v in d.validations:
            fn = st.error if v["level"] == "ERROR" else st.warning if v["level"] == "WARNING" else st.info
            fn(f"{v['rule']}: {v['message']}")
        if d.gstr2b_match:
            st.info(f"GSTR-2B: {d.gstr2b_match}")
        for r in d.remarks:
            if not r.startswith("["):
                st.caption("• " + r)
        with st.form(f"form-{d.doc_id}"):
            new = {}
            new["doc_type"] = st.selectbox("Document type", DOC_TYPES, index=DOC_TYPES.index(d.doc_type))
            dirs = [INWARD, OUTWARD, THIRD_PARTY, CONSIGNEE_ONLY, DIRECTION_UNKNOWN]
            new["direction"] = st.selectbox("Direction", dirs, index=dirs.index(d.direction) if d.direction in dirs else 4)
            fcols = st.columns(2)
            for i, field in enumerate([x for x in rm.EDITABLE_FIELDS if x not in ("doc_type", "direction", "review_status")]):
                cur = d.fields.get(field)
                cur = "" if cur is None else (", ".join(cur) if isinstance(cur, list) else str(cur))
                conf = d.field_confidence.get(field)
                label = field.replace("_", " ").title() + (f"  ({conf:.0%})" if isinstance(conf, (int, float)) else "")
                new[field] = fcols[i % 2].text_input(label, cur, key=f"{d.doc_id}-{field}")
            note = st.text_input("Reviewer note (optional)")
            b1, b2 = st.columns(2)
            save = b1.form_submit_button("💾 Save corrections", type="primary")
            approve = b2.form_submit_button("✔ Save & mark verified")
        if save or approve:
            changed = 0
            for field, val in new.items():
                old = d.doc_type if field == "doc_type" else d.direction if field == "direction" else d.fields.get(field)
                oldv = "" if old is None else str(old)
                if str(val) != oldv:
                    try:
                        rm.record_correction(res["job_id"], d, field, rm.coerce(field, val), user="reviewer")
                        changed += 1
                    except ValueError as e:
                        st.error(str(e))
            if approve:
                rm.approve(res["job_id"], d, note=note)
            refresh()
            st.success(f"Saved {changed} correction(s)" + (" and approval" if approve else "") + "; document re-validated.")
            st.rerun()
        hist = rm.correction_history(res["job_id"], d.doc_id)
        if hist:
            with st.expander(f"Audit trail ({len(hist)})"):
                st.dataframe(pd.DataFrame(hist)[["at", "user", "field", "old_value", "new_value"]], hide_index=True)


# ------------------------------------------------------------------ duplicates
def page_duplicates():
    st.header("Duplicate Review")
    res = current()
    if not res:
        st.info("Process a batch first.")
        return
    rows = [r for r in res["duplicates"] if r.get("Group")]
    if not rows:
        st.success("No duplicate candidates.")
        return
    df = pd.DataFrame(rows)[["Group", "Level", "Primary Document", "Duplicate Candidate", "Document Number",
                             "Match Reason", "Differences", "Match Confidence", "Suggested Action", "User Decision"]]
    ev = st.dataframe(df, hide_index=True, width="stretch", height=300, on_select="rerun",
                      selection_mode="single-row", key="dup_table")
    sel = ev.selection.rows if ev and hasattr(ev, "selection") else []
    r = rows[sel[0]] if sel else rows[0]
    byid = {d.doc_id: d for d in res["documents"]}
    a, b = byid.get(r["Primary Doc ID"]), byid.get(r["Candidate Doc ID"])
    st.subheader(f"{r['Group']} - {r['Level']}")
    st.caption(r["Match Reason"])
    if a and b:
        keys = ["invoice_number", "invoice_date", "supplier_gstin", "customer_gstin", "taxable_value", "igst", "cgst",
                "sgst", "invoice_total", "irn"]
        cmp = pd.DataFrame({"Field": keys, "Primary": [str(a.fields.get(k, "")) for k in keys],
                            "Candidate": [str(b.fields.get(k, "")) for k in keys]})
        cmp["Same"] = cmp["Primary"] == cmp["Candidate"]
        st.dataframe(cmp, hide_index=True, width="stretch")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f"**Primary:** {a.rel_path} [p{a.pages}]")
            _preview(a, res)
        with c2:
            st.markdown(f"**Candidate:** {b.rel_path} [p{b.pages}]")
            _preview(b, res)
    choice = st.radio("Decision", rm_decisions(), horizontal=True,
                      index=rm_decisions().index(r["User Decision"]) if r.get("User Decision") in rm_decisions() else 0)
    if st.button("Save decision", type="primary"):
        rm.save_duplicate_decision(res["job_id"], r["Primary Doc ID"], r["Candidate Doc ID"], choice)
        refresh()
        st.success("Decision saved. 'Separate transaction' returns the candidate to normal validation; "
                   "'Exact duplicate' / 'Revised invoice' keep it out of the totals.")
        st.rerun()


def rm_decisions():
    from duplicate_detector import DECISIONS
    return DECISIONS


# ------------------------------------------------------------------ reports
def page_reports():
    st.header("Reports & Export")
    res = current()
    if not res:
        st.info("Process a batch first.")
        return
    s = settings()
    probs = _problems(res["job_id"])
    # audit fix E10: the uploaded register was exported while both Caliber GSTINs were still
    # 'Detected - Unverified', so every invoice sat in Pending_Review and Final_Tax_Invoices was empty.
    held = [d for d in res["documents"] if d.status == ST_REVIEW
            and any(v["rule"] == "MASTER_UNVERIFIED" for v in d.validations)]
    only_master = [d for d in held if not any(v["level"] == "ERROR" or (v["level"] == "WARNING" and v["rule"] != "MASTER_UNVERIFIED")
                                              for v in d.validations)]
    if held:
        st.error(f"⚠ {len(held)} invoice(s) are held in Pending_Review because the Caliber GST registration is not "
                 f"verified yet ({len(only_master)} of them have no other issue). Open **Caliber GST Master**, verify the "
                 f"registrations, then build the workbook - otherwise Final_Tax_Invoices will be incomplete.")
    sig = (res["job_id"], tuple(d.status for d in res["documents"]))
    if st.session_state.get("wb_sig") != sig:
        st.session_state.pop("wb", None)     # statuses changed since the last build (e.g. master verified)
    n_final = sum(1 for d in res["documents"] if d.status == ST_VERIFIED and d.doc_type in INVOICE_TYPES
                  and (d.direction == INWARD or (s.include_outward_in_final and d.direction == OUTWARD)))
    st.info(f"Final_Tax_Invoices will contain {n_final} verified invoice(s).")
    if st.button("Build consolidated workbook", type="primary"):
        st.session_state.wb = build_workbook(res, s, upload_problems=probs, as_bytes=True)
        st.session_state.wb_sig = sig
    if st.session_state.get("wb"):
        st.download_button("⬇ Download Caliber_GST_Register.xlsx", st.session_state.wb,
                           f"Caliber_GST_Register_{res['job_id']}.xlsx", type="primary")
        from dashboard import build_data, render_html
        html = render_html(build_data(st.session_state.wb, f"Caliber_GST_Register_{res['job_id']}.xlsx"))
        st.download_button("⬇ Download HTML dashboard (index.html)", html.encode("utf-8"), "index.html", "text/html",
                           help="One self-contained web page: open it in any browser or publish it on GitHub Pages")
    st.subheader("Individual sheets")
    docs = res["documents"]
    sheets = {
        "Final_Tax_Invoices": [invoice_row(d) for d in docs if d.status == ST_VERIFIED and d.doc_type in INVOICE_TYPES
                               and (d.direction == INWARD or (s.include_outward_in_final and d.direction == OUTWARD))],
        "Pending_Review": [invoice_row(d) for d in docs if d.status == ST_REVIEW],
        "Credit_Debit_Notes": [invoice_row(d) for d in docs if d.doc_type in NOTE_TYPES],
        "Outward_Sales_Register": [invoice_row(d) for d in docs if d.doc_type in INVOICE_TYPES and d.direction == OUTWARD],
        "Duplicate_Report": res["duplicates"],
        "Supplier_Master": supplier_master_rows(docs),
    }
    name = st.selectbox("Sheet", list(sheets))
    df = pd.DataFrame(sheets[name])
    st.dataframe(df, hide_index=True, width="stretch", height=300)
    c1, c2 = st.columns(2)
    c1.download_button("⬇ CSV", df.to_csv(index=False).encode("utf-8-sig"), f"{name}.csv")
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    c2.download_button("⬇ Excel", buf.getvalue(), f"{name}.xlsx")
    st.subheader("Logs")
    with db.connect() as con:
        log = pd.DataFrame([dict(r) for r in con.execute("SELECT * FROM audit_log WHERE job_id IN (?, '') ORDER BY id",
                                                         (res["job_id"],))])
    st.download_button("⬇ Processing & audit log (CSV)", log.to_csv(index=False).encode("utf-8-sig"), "audit_log.csv")
    exc = []
    for d in docs:
        for v in d.validations:
            if v["level"] in ("ERROR", "WARNING"):
                exc.append({"File": d.rel_path, "Pages": d.pages, "Type": d.doc_type, "Status": d.status,
                            "Level": v["level"], "Rule": v["rule"], "Message": v["message"]})
    for f in res["files"]:
        if f.status in ("Failed", "Skipped"):
            exc.append({"File": f.rel_path, "Status": f.status, "Level": "ERROR" if f.status == "Failed" else "INFO",
                        "Rule": "FILE", "Message": f.error})
    st.download_button("⬇ Exceptions report (CSV)", pd.DataFrame(exc).to_csv(index=False).encode("utf-8-sig"), "exceptions.csv")


def _app_password() -> str:
    """Optional access password for an online deployment.  Read from the environment or from
    Streamlit secrets (.streamlit/secrets.toml / Streamlit Cloud 'Secrets') - never from the code."""
    if os.environ.get("APP_PASSWORD"):
        return os.environ["APP_PASSWORD"]
    try:
        return str(st.secrets.get("APP_PASSWORD", ""))
    except FileNotFoundError:       # no secrets file at all (normal on a local install)
        return ""


def password_gate() -> None:
    pw = _app_password()
    if not pw or st.session_state.get("authenticated"):
        return
    st.title("📑 Caliber GST")
    entered = st.text_input("Password", type="password", key="pw_input")
    if entered:
        import hmac
        if hmac.compare_digest(entered, pw):
            st.session_state.authenticated = True
            st.session_state.pop("pw_input", None)
            st.rerun()
        st.error("Wrong password.")
    st.stop()


password_gate()

PAGES = {"Dashboard": page_dashboard, "Upload & Processing": page_upload, "Caliber GST Master": page_master,
         "Document Review": page_review, "Duplicate Review": page_duplicates, "Reports & Export": page_reports}
PAGES[sidebar()]()

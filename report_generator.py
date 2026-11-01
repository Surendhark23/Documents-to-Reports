"""Consolidated Excel workbook.

Sheets 1-10 follow the specification; three supporting sheets are appended
(Outward_Sales_Register, Tax_Rate_Breakdown, GSTR2B_Register_Rows).
Dates are written as real Excel dates (DD-MM-YYYY), invoice numbers and
GSTINs as text, amounts as numbers.  Missing values stay blank - never 0.
Only Verified inward tax invoices count toward the final totals.
"""
from __future__ import annotations

import io
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from config import data_dir, APP_VERSION
from gst_validator import state_name
from models import (INVOICE_COLUMNS, EXTRA_COLUMNS, INVOICE_TYPES, NOTE_TYPES, NON_TAX_TYPES, EXCEL_REGISTER,
                    UNRELATED, UNKNOWN, INWARD, OUTWARD, ST_VERIFIED, ST_REVIEW, ST_DUPLICATE, ST_OUTWARD, Document)

HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")
HEADER_FONT = Font(bold=True, color="FFFFFF")
REVIEW_FILL = PatternFill("solid", fgColor="FFF2CC")
ERROR_FILL = PatternFill("solid", fgColor="F8D7DA")
OK_FILL = PatternFill("solid", fgColor="E2F0D9")
DUP_FILL = PatternFill("solid", fgColor="E7E6F7")
TOTAL_FONT = Font(bold=True)
THIN = Side(style="thin", color="D0D0D0")
AMOUNT_COLS = {"Taxable Value", "IGST", "CGST", "SGST", "UTGST", "Cess", "Other Charges", "Discount",
               "Round-Off", "Invoice Total", "Invoice Value", "Tax", "Total Tax", "Note Value"}
TEXT_COLS = {"Invoice Number", "Original Invoice Number", "Supplier GSTIN", "Customer GSTIN", "Bill To GSTIN",
             "Ship To GSTIN", "GSTIN", "HSN/SAC", "PO Number", "IRN", "E-Way Bill Number", "Document Number",
             "State Code", "Supplier State Code", "Customer State Code", "File SHA-256", "Document ID"}


def _date(iso: str):
    if not iso:
        return None
    try:
        return datetime.strptime(iso, "%Y-%m-%d")
    except ValueError:
        return None


def invoice_row(d: Document) -> dict:
    f = d.fields
    issues = [v["message"] for v in d.validations if v["level"] in ("ERROR", "WARNING")]
    vstatus = ("Errors: " + str(sum(v["level"] == "ERROR" for v in d.validations)) + ", Warnings: " +
               str(sum(v["level"] == "WARNING" for v in d.validations))) if issues else "Passed"
    rates = f.get("tax_rates") or []
    row = {
        "Source File Name": d.file_name, "Source File Path / Relative Path": d.rel_path,
        "Page Number or Page Range": d.pages, "Document Type": d.doc_type, "Document Status": d.status,
        "Invoice Date": _date(f.get("invoice_date")), "Invoice Number": f.get("invoice_number", ""),
        "Original Invoice Number": f.get("original_invoice_number", ""),
        "Supplier Legal Name": f.get("supplier_legal_name", ""), "Supplier Trade Name": f.get("supplier_trade_name", ""),
        "Supplier GSTIN": f.get("supplier_gstin", ""), "Supplier State": f.get("supplier_state", ""),
        "Supplier State Code": f.get("supplier_state_code", ""), "Supplier Address": f.get("supplier_address", ""),
        "Customer Legal Name": f.get("customer_legal_name", ""), "Customer Trade Name": f.get("customer_trade_name", ""),
        "Customer GSTIN": f.get("customer_gstin", ""), "Customer State": f.get("customer_state", ""),
        "Customer State Code": f.get("customer_state_code", ""), "Customer Address": f.get("customer_address", ""),
        "Bill To Name": f.get("bill_to_name", ""), "Bill To GSTIN": f.get("bill_to_gstin", ""),
        "Ship To Name": f.get("ship_to_name", ""), "Ship To GSTIN": f.get("ship_to_gstin", ""),
        "Place of Supply": f.get("place_of_supply", ""), "Reverse Charge Indicator": f.get("reverse_charge", ""),
        "Currency": f.get("currency", ""), "Taxable Value": f.get("taxable_value"), "IGST": f.get("igst"),
        "CGST": f.get("cgst"), "SGST": f.get("sgst"), "UTGST": f.get("utgst"), "Cess": f.get("cess"),
        "Other Charges": f.get("other_charges"), "Discount": f.get("discount"), "Round-Off": f.get("round_off"),
        "Invoice Total": f.get("invoice_total"), "HSN/SAC": ", ".join(f.get("hsn") or []),
        "Tax Rate or Tax Rates": ", ".join(f"{r:g}%" for r in rates), "PO Number": f.get("po_number", ""),
        "IRN": f.get("irn", ""), "E-Way Bill Number": f.get("eway_bill", ""),
        "OCR / Extraction Confidence": round(d.extraction_confidence, 2), "Validation Status": vstatus,
        "Review Remarks": " | ".join(dict.fromkeys(d.remarks)),
        "Transaction Direction": d.direction,
        "Caliber Registration (State)": f"{d.caliber_gstin} ({d.caliber_state})" if d.caliber_gstin else "",
        "GSTR-2B Match": d.gstr2b_match, "Duplicate Group": f"{d.duplicate_group} {d.duplicate_role}".strip(),
        "Document ID": d.doc_id, "File SHA-256": d.file_id,
    }
    return row


def _write_sheet(wb: Workbook, title: str, columns: list[str], rows: list[dict], totals: Optional[list[str]] = None,
                 status_col: Optional[str] = None, note: str = "") -> None:
    ws = wb.create_sheet(title)
    r0 = 1
    if note:
        ws.cell(row=1, column=1, value=note).font = Font(italic=True, color="555555")
        r0 = 2
    for c, name in enumerate(columns, start=1):
        cell = ws.cell(row=r0, column=c, value=name)
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for i, row in enumerate(rows, start=r0 + 1):
        for c, name in enumerate(columns, start=1):
            v = row.get(name)
            if isinstance(v, (list, tuple)):
                v = ", ".join(map(str, v))
            if name in TEXT_COLS and v is not None and not isinstance(v, str):
                v = str(v)
            cell = ws.cell(row=i, column=c, value=v)
            if name in AMOUNT_COLS:
                cell.number_format = "#,##0.00"
            elif isinstance(v, datetime):
                cell.number_format = "DD-MM-YYYY"
            elif name in TEXT_COLS:
                cell.number_format = "@"
        if status_col:
            st = str(row.get(status_col, ""))
            fill = OK_FILL if st.startswith(ST_VERIFIED) else ERROR_FILL if ("Failed" in st or "Error" in st) \
                else DUP_FILL if "Duplicate" in st else REVIEW_FILL if ("Review" in st or "Unknown" in st) else None
            if fill:
                ws.cell(row=i, column=columns.index(status_col) + 1).fill = fill
    last = r0 + len(rows)
    if totals and rows:
        tr = last + 1
        ws.cell(row=tr, column=1, value="TOTAL").font = TOTAL_FONT
        for name in totals:
            if name in columns:
                c = columns.index(name) + 1
                col = get_column_letter(c)
                cell = ws.cell(row=tr, column=c, value=f"=SUBTOTAL(9,{col}{r0 + 1}:{col}{last})")
                cell.font, cell.number_format = TOTAL_FONT, "#,##0.00"
    ws.freeze_panes = ws.cell(row=r0 + 1, column=min(3, len(columns)))
    if rows:
        ws.auto_filter.ref = f"A{r0}:{get_column_letter(len(columns))}{last}"
    for c, name in enumerate(columns, start=1):
        sample = [len(str(r.get(name) or "")) for r in rows[:200]]
        width = min(60, max(10, len(name) * 0.9, (max(sample) if sample else 0) * 0.95))
        ws.column_dimensions[get_column_letter(c)].width = width
    ws.row_dimensions[r0].height = 32


def supplier_master_rows(docs: list[Document]) -> list[dict]:
    agg = defaultdict(lambda: {"names": defaultdict(int), "addr": "", "docs": 0, "types": set()})
    for d in docs:
        f = d.fields
        g = f.get("supplier_gstin")
        if not g or d.direction == OUTWARD:
            continue
        a = agg[g]
        if f.get("supplier_legal_name"):
            a["names"][f["supplier_legal_name"].upper()] += 1
        a["addr"] = a["addr"] or f.get("supplier_address", "")
        a["docs"] += 1
        a["types"].add(d.doc_type)
    rows = []
    for g, a in sorted(agg.items()):
        names = sorted(a["names"].items(), key=lambda x: -x[1])
        status = "Auto-extracted - verify" if len(names) <= 1 else "Conflicting names - review"
        rows.append({"GSTIN": g, "Supplier Legal Name": names[0][0] if names else "",
                     "Supplier Trade Name": "", "State Code": g[:2], "State": state_name(g),
                     "Address": a["addr"], "Documents": a["docs"], "Document Types": ", ".join(sorted(a["types"])),
                     "Name Variants": " | ".join(n for n, _ in names[1:]), "Verification Status": status})
    return rows


def build_workbook(res: dict, settings, out_path: Optional[Path] = None, upload_problems: Optional[list] = None,
                   as_bytes: bool = False):
    docs: list[Document] = res["documents"]
    files = res["files"]
    master = res["master"]
    dups = res["duplicates"]
    wb = Workbook()
    wb.remove(wb.active)
    cols = INVOICE_COLUMNS + EXTRA_COLUMNS
    money_totals = ["Taxable Value", "IGST", "CGST", "SGST", "UTGST", "Cess", "Invoice Total"]

    def is_final(d):
        return (d.doc_type in INVOICE_TYPES and d.status == ST_VERIFIED and
                (d.direction == INWARD or (settings.include_outward_in_final and d.direction == OUTWARD)))

    final = [d for d in docs if is_final(d)]
    pending = [d for d in docs if d.status == ST_REVIEW and d.direction != OUTWARD and
               (d.doc_type in INVOICE_TYPES | NOTE_TYPES or d.doc_type == UNKNOWN)]
    notes = [d for d in docs if d.doc_type in NOTE_TYPES]
    non_tax = [d for d in docs if d.doc_type in NON_TAX_TYPES | {UNRELATED, EXCEL_REGISTER}]
    outward = [d for d in docs if d.doc_type in INVOICE_TYPES and d.direction == OUTWARD]

    # 1 Final
    _write_sheet(wb, "Final_Tax_Invoices", cols, [invoice_row(d) for d in final], money_totals, "Document Status",
                 note=f"Verified inward tax invoices only ({len(final)}). Quotations, proformas, orders, notes, duplicates and unverified records are excluded from these totals.")
    # 2 Pending
    prow = []
    for d in pending:
        r = invoice_row(d)
        r["Issues"] = " | ".join(f"[{v['level']}] {v['message']}" for v in d.validations if v["level"] != "INFO")
        r["Classification Reason"] = d.type_reason
        prow.append(r)
    _write_sheet(wb, "Pending_Review", ["Issues", "Classification Reason"] + cols, prow, None, "Document Status",
                 note="Records with missing, conflicting, low-confidence or unverified information. Correct them in the app (Document Review) and re-export.")
    # 3 Document register
    reg = []
    by_path = defaultdict(list)
    for d in docs:
        by_path[d.rel_path].append(d)
    for f in files:
        ds = by_path.get(f.rel_path, [])
        if not ds:
            reg.append({"Source File Name": f.name, "Relative Path": f.rel_path, "Pages": f.pages or "",
                        "File Size (KB)": round(f.size / 1024, 1), "File Type": f.kind, "SHA-256": f.file_id,
                        "Processing Status": f.status, "Document Type": "", "Type Confidence": None,
                        "Classification Reason": "", "Direction": "", "Document Status": "Processing Failed" if f.status == "Failed" else f.status,
                        "Included in Final Register": "No", "Reason for Inclusion / Exclusion": f.error or f.status,
                        "Text Source": "", "OCR Confidence": None, "Duplicate Group": ""})
            continue
        for d in ds:
            incl = is_final(d)
            if incl:
                why = "Verified inward tax invoice"
            elif d.status == ST_DUPLICATE:
                why = f"Duplicate/version of another document ({d.duplicate_group})"
            elif d.status == ST_OUTWARD:
                why = "Caliber's own outward document - see Outward_Sales_Register"
            elif d.doc_type not in INVOICE_TYPES | NOTE_TYPES:
                why = f"Not a tax invoice: {d.doc_type}"
            elif d.doc_type in NOTE_TYPES:
                why = "Credit/debit note - see Credit_Debit_Notes"
            else:
                why = "Pending review: " + "; ".join(v["message"] for v in d.validations if v["level"] == "ERROR")[:300] \
                      if any(v["level"] == "ERROR" for v in d.validations) else "Pending review (warnings / confidence)"
            reg.append({"Source File Name": d.file_name, "Relative Path": d.rel_path, "Pages": d.pages,
                        "File Size (KB)": round(f.size / 1024, 1), "File Type": f.kind, "SHA-256": f.file_id,
                        "Processing Status": f.status, "Document Type": d.doc_type, "Type Confidence": d.type_confidence,
                        "Classification Reason": d.type_reason, "Direction": d.direction, "Document Status": d.status,
                        "Included in Final Register": "Yes" if incl else "No", "Reason for Inclusion / Exclusion": why,
                        "Text Source": d.text_source,
                        "OCR Confidence": d.ocr_confidence, "Duplicate Group": f"{d.duplicate_group} {d.duplicate_role}".strip()})
    _write_sheet(wb, "Document_Register", list(reg[0].keys()) if reg else ["Source File Name"], reg, None, "Document Status",
                 note=f"Every processed file ({len(files)} files, {len(docs)} documents). Nothing is silently excluded.")
    # 4 Non-tax
    nrows = [{"Source File Name": d.file_name, "Relative Path": d.rel_path, "Pages": d.pages, "Document Type": d.doc_type,
              "Classification Reason": d.type_reason, "Direction": d.direction,
              "Issuer / Supplier": d.fields.get("supplier_legal_name") or d.fields.get("issuer_name", ""),
              "Supplier GSTIN": d.fields.get("supplier_gstin", ""), "Customer GSTIN": d.fields.get("customer_gstin", ""),
              "Document Number": d.fields.get("invoice_number", ""), "Document Date": _date(d.fields.get("invoice_date")),
              "Taxable Value": d.fields.get("taxable_value"), "Tax": sum(x or 0 for x in (d.fields.get("igst"), d.fields.get("cgst"), d.fields.get("sgst"))) or None,
              "Invoice Total": d.fields.get("invoice_total"), "Duplicate Group": f"{d.duplicate_group} {d.duplicate_role}".strip(),
              "Remarks": " | ".join(dict.fromkeys(d.remarks))} for d in non_tax]
    _write_sheet(wb, "Non_Tax_Documents", list(nrows[0].keys()) if nrows else ["Source File Name"], nrows, None, None,
                 note="Quotations, proformas, purchase/sales orders, challans, payments, statements, registers and unrelated files - never counted as tax invoices. Amounts are informational.")
    # 5 Notes
    _write_sheet(wb, "Credit_Debit_Notes", cols, [invoice_row(d) for d in notes], money_totals, "Document Status",
                 note="Credit and debit notes (inward and Caliber-issued). Recorded separately - not netted into the invoice register.")
    # 6 Duplicates
    dcols = ["Group", "Level", "Primary Document", "Duplicate Candidate", "Document Type", "Supplier GSTIN",
             "Document Number", "Match Reason", "Differences", "Match Confidence", "Suggested Action", "User Decision",
             "Primary Doc ID", "Candidate Doc ID"]
    _write_sheet(wb, "Duplicate_Report", dcols, dups, None, None,
                 note="No file is deleted. Confirm each candidate in the app (Duplicate Review): Exact duplicate / Revised invoice / Separate transaction.")
    # 7 Master
    mdf = master.to_dataframe()
    _write_sheet(wb, "Caliber_GST_Master", list(mdf.columns), mdf.to_dict("records"), None, "Verification Status",
                 note="Evidence-based: only registrations found on the documents. Verify in the app before relying on the register.")
    # 8 Supplier master
    srows = supplier_master_rows(docs)
    _write_sheet(wb, "Supplier_Master", list(srows[0].keys()) if srows else ["GSTIN"], srows, None, "Verification Status")
    # 9 Tax summary
    _tax_summary(wb, final, docs)
    # 10 Errors
    errs = []
    for p in upload_problems or []:
        errs.append({"File": p.get("path"), "Stage": "Upload / ZIP extraction", "Severity": p.get("level", "ERROR"), "Issue": p.get("error")})
    for f in files:
        if f.status in ("Failed", "Skipped"):
            errs.append({"File": f.rel_path, "Stage": "Text extraction" if f.status == "Failed" else "Inventory",
                         "Severity": "ERROR" if f.status == "Failed" else "INFO", "Issue": f.error})
        elif f.error:
            errs.append({"File": f.rel_path, "Stage": "Text extraction", "Severity": "WARNING", "Issue": f.error})
    for d in docs:
        if d.ocr_confidence is not None and d.ocr_confidence < 0.6:
            errs.append({"File": d.rel_path, "Stage": "OCR", "Severity": "WARNING",
                         "Issue": f"Low OCR confidence {d.ocr_confidence:.2f} on pages {d.pages} - verify manually"})
    _write_sheet(wb, "Processing_Errors", ["File", "Stage", "Severity", "Issue"], errs, None, "Severity")
    # extras
    _write_sheet(wb, "Outward_Sales_Register", cols, [invoice_row(d) for d in outward], money_totals, "Document Status",
                 note="Caliber's own sales invoices found in the batch (Caliber is the supplier). Kept out of the purchase register.")
    li = []
    for d in docs:
        if d.doc_type in INVOICE_TYPES | NOTE_TYPES:
            for r in d.line_items:
                li.append({"Source File Name": d.file_name, "Page": d.pages, "Invoice Number": d.fields.get("invoice_number", ""),
                           "Supplier GSTIN": d.fields.get("supplier_gstin", ""), "Document Status": d.status,
                           "Tax Rate %": r.get("rate"), "Taxable Value": r.get("taxable_value"), "IGST": r.get("igst"),
                           "CGST": r.get("cgst"), "SGST": r.get("sgst"), "Breakdown Status": r.get("breakdown_status")})
    _write_sheet(wb, "Tax_Rate_Breakdown", list(li[0].keys()) if li else ["Source File Name"], li, None, None,
                 note="Rate-wise split derived from line/HSN rows where taxable x rate = tax. 'Partial' means the rows did not add up to the invoice taxable value.")
    rr = []
    for d in docs:
        for r in d.register_rows:
            rr.append({"Register File": d.rel_path, "Sheet": r.get("sheet"), "Row": r.get("row"),
                       "Caliber GSTIN": r.get("customer_gstin", ""), "Supplier GSTIN": r.get("supplier_gstin"),
                       "Supplier Name": r.get("supplier_name"), "Record Type": r.get("record_type"),
                       "Invoice Number": r.get("invoice_number"), "Invoice Date": _date(r.get("invoice_date")),
                       "Invoice Value": r.get("invoice_value"), "Taxable Value": r.get("taxable_value"),
                       "IGST": r.get("igst"), "CGST": r.get("cgst"), "SGST": r.get("sgst"), "Cess": r.get("cess"),
                       "Period": r.get("period"), "ITC Available": r.get("itc_available"),
                       "Reverse Charge": r.get("reverse_charge")})
    _write_sheet(wb, "GSTR2B_Register_Rows", list(rr[0].keys()) if rr else ["Register File"], rr,
                 ["Invoice Value", "Taxable Value", "IGST", "CGST", "SGST"], None,
                 note="Rows read from uploaded GSTR-2B / Excel registers; used to cross-check PDF invoices (column 'GSTR-2B Match').")

    _gstin_wise_sheets(wb, docs, final, master)
    _executive_summary(wb, res, files, docs, final, pending, notes, non_tax, outward)

    if as_bytes:
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()
    if out_path is None:
        exp = data_dir() / "exports"
        exp.mkdir(parents=True, exist_ok=True)
        out_path = exp / f"Caliber_GST_Register_{res['job_id']}.xlsx"
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


def _tax_summary(wb: Workbook, final: list[Document], docs: list[Document]) -> None:
    ws = wb.create_sheet("Tax_Summary")
    keys = ["taxable_value", "igst", "cgst", "sgst", "utgst", "cess", "invoice_total"]
    heads = ["Taxable Value", "IGST", "CGST", "SGST", "UTGST", "Cess", "Invoice Total"]
    row = 1

    def block(title, groups: dict, first_col: str, note=""):
        nonlocal row
        ws.cell(row=row, column=1, value=title).font = Font(bold=True, size=12)
        if note:
            ws.cell(row=row, column=3, value=note).font = Font(italic=True, color="555555")
        row += 1
        for c, h in enumerate([first_col, "Documents"] + heads, start=1):
            cell = ws.cell(row=row, column=c, value=h)
            cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        row += 1
        start = row
        for k, ds in sorted(groups.items()):
            ws.cell(row=row, column=1, value=k)
            ws.cell(row=row, column=2, value=len(ds))
            for c, key in enumerate(keys, start=3):
                vals = [d.fields.get(key) for d in ds if d.fields.get(key) is not None]
                cell = ws.cell(row=row, column=c, value=round(sum(vals), 2) if vals else None)
                cell.number_format = "#,##0.00"
            row += 1
        ws.cell(row=row, column=1, value="TOTAL").font = TOTAL_FONT
        for c in range(2, 3 + len(keys)):
            col = get_column_letter(c)
            cell = ws.cell(row=row, column=c, value=f"=SUM({col}{start}:{col}{row - 1})" if row > start else 0)
            cell.font, cell.number_format = TOTAL_FONT, "#,##0.00" if c > 2 else "0"
        row += 3

    by_state, by_sup, by_type = defaultdict(list), defaultdict(list), defaultdict(list)
    for d in final:
        by_state[f"{d.caliber_gstin} - {d.caliber_state}" if d.caliber_gstin else "(unassigned)"].append(d)
        by_sup[f"{d.fields.get('supplier_gstin', '')} - {d.fields.get('supplier_legal_name', '')}"].append(d)
    for d in docs:
        if d.doc_type in INVOICE_TYPES | NOTE_TYPES and d.status != ST_DUPLICATE:
            by_type[f"{d.doc_type} | {d.direction.split(' (')[0]} | {d.status}"].append(d)
    block("A. Caliber registration (state) - final verified invoices", by_state, "Caliber GSTIN / State")
    block("B. Supplier - final verified invoices", by_sup, "Supplier")
    block("C. Document type / direction / status - all extracted (informational)", by_type, "Type | Direction | Status",
          note="Includes unverified and outward documents - NOT final totals")
    ws.column_dimensions["A"].width = 70
    for c in range(2, 10):
        ws.column_dimensions[get_column_letter(c)].width = 16
    ws.freeze_panes = "B2"


# ------------------------------------------------------------------ executive summary (audit enhancement)
WARN_FONT = Font(bold=True, color="9C0006")


def _executive_summary(wb: Workbook, res: dict, files, docs: list[Document], final, pending, notes, non_tax,
                       outward) -> None:
    """First sheet: what was processed, where every document went, control totals and the
    reasons records are held back.  Every document is counted exactly once in section B."""
    ws = wb.create_sheet("Executive_Summary", 0)
    ws.column_dimensions["A"].width = 58
    for col in "BCDE":
        ws.column_dimensions[col].width = 18
    row = [1]

    def put(*vals, bold=False, fill=None, fmt=None, font=None):
        r = row[0]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            if bold:
                cell.font = TOTAL_FONT
            if font is not None:
                cell.font = font
            if fill is not None:
                cell.fill = fill
            if fmt and isinstance(v, (int, float)) and c > 1:
                cell.number_format = fmt
        row[0] += 1

    def head(title):
        row[0] += 1
        r = row[0]
        for c in range(1, 6):
            ws.cell(row=r, column=c).fill = HEADER_FILL
        ws.cell(row=r, column=1, value=title).font = HEADER_FONT
        row[0] += 1

    ws.cell(row=1, column=1, value="Caliber GST Register - Executive Summary").font = Font(bold=True, size=14)
    row[0] = 2
    put("Batch / job ID", res.get("job_id", ""))
    put("Generated", datetime.now().strftime("%d-%m-%Y %H:%M"))
    put("Application version", APP_VERSION)

    # --- warnings first: anything that makes the Final register incomplete
    held = [d for d in docs if d.status == ST_REVIEW and any(v["rule"] == "MASTER_UNVERIFIED" for v in d.validations)]
    if held:
        head("ATTENTION")
        only = [d for d in held if not any(v["level"] == "ERROR" or (v["level"] == "WARNING" and v["rule"] != "MASTER_UNVERIFIED")
                                           for v in d.validations)]
        put(f"{len(held)} invoice(s) in Pending_Review are flagged 'Caliber GSTIN not verified' - for {len(only)} of them "
            "this is the only issue. Verify the registrations in the app (Caliber GST Master) and export again; until then "
            "Final_Tax_Invoices is incomplete.", font=WARN_FONT, fill=ERROR_FILL)

    # --- A: input reconciliation
    head("A. Input files")
    by_status = defaultdict(int)
    for f in files:
        by_status[f.status or "Unknown"] += 1
    put("Files received", len(files), bold=True)
    for k in sorted(by_status):
        put(f"   {k}", by_status[k])
    put("Documents identified (a file can hold several invoices / pages)", len(docs), bold=True)

    # --- B: where every document went
    head("B. Document reconciliation (each document counted once)")
    put("Destination", "Documents", "Taxable Value", "Total Tax", "Invoice Total", bold=True)

    def sums(ds):
        tv = sum(d.fields.get("taxable_value") or 0 for d in ds)
        tx = sum(sum(d.fields.get(k) or 0 for k in ("igst", "cgst", "sgst", "utgst", "cess")) for d in ds)
        tot = sum(d.fields.get("invoice_total") or 0 for d in ds)
        return tv, tx, tot

    ids_final, ids_pend = {d.doc_id for d in final}, {d.doc_id for d in pending}
    ids_notes, ids_non, ids_out = {d.doc_id for d in notes}, {d.doc_id for d in non_tax}, {d.doc_id for d in outward}
    buckets = [("Final_Tax_Invoices (Verified inward tax invoices)", [d for d in docs if d.doc_id in ids_final]),
               ("Pending_Review (needs a person)", [d for d in docs if d.doc_id in ids_pend and d.doc_id not in ids_final]),
               ("Duplicate candidates (Duplicate_Report)", [d for d in docs if d.status == ST_DUPLICATE]),
               ("Credit_Debit_Notes", [d for d in docs if d.doc_id in ids_notes and d.status != ST_DUPLICATE
                                       and d.doc_id not in ids_pend]),
               ("Outward_Sales_Register (Caliber is supplier)", [d for d in docs if d.doc_id in ids_out
                                                                 and d.status != ST_DUPLICATE]),
               ("Non_Tax_Documents (quotations, proformas, e-way bills ...)",
                [d for d in docs if d.doc_id in ids_non and d.status != ST_DUPLICATE])]
    seen = set()
    for name, ds in buckets:
        ds = [d for d in ds if d.doc_id not in seen]
        seen.update(d.doc_id for d in ds)
        tv, tx, tot = sums(ds)
        put(name, len(ds), tv or None, tx or None, tot or None, fmt="#,##0.00")
    other = [d for d in docs if d.doc_id not in seen]
    tv, tx, tot = sums(other)
    put("Other / not classified (see Document_Register)", len(other), tv or None, tx or None, tot or None,
        fmt="#,##0.00", fill=REVIEW_FILL if other else None)
    put("TOTAL documents", len(docs), bold=True)
    ok = len(seen) + len(other) == len(docs)
    put("Control check: every document accounted for", "OK" if ok else "MISMATCH", bold=True,
        fill=OK_FILL if ok else ERROR_FILL)

    # --- C: why records are pending
    head("C. Pending_Review - reasons (one document can have several)")
    reasons = defaultdict(int)
    for d in pending:
        for v in d.validations:
            if v["level"] in ("ERROR", "WARNING"):
                reasons[(v["level"], v["rule"])] += 1
    put("Rule", "Level", "Documents", bold=True)
    for (lvl, rule), n in sorted(reasons.items(), key=lambda x: (-x[1], x[0])):
        put(rule, lvl, n, fill=ERROR_FILL if lvl == "ERROR" else REVIEW_FILL)
    if not reasons:
        put("None")

    # --- D: type breakdown
    head("D. Documents by type")
    types = defaultdict(int)
    for d in docs:
        types[d.doc_type] += 1
    for t, n in sorted(types.items(), key=lambda x: -x[1]):
        put(t, n)
    ws.freeze_panes = "A2"


# ------------------------------------------------------------------ GSTIN-wise portal check (v1.1.1)
PORTAL_COLS = ["GSTIN of supplier", "Trade/Legal name", "Invoice number", "Invoice type", "Invoice Date",
               "Invoice Value", "Place of supply", "Supply Attract Reverse Charge", "Taxable Value",
               "Integrated Tax", "Central Tax", "State/UT Tax", "Cess"]
EXTRA_PORTAL = ["Status in this register", "Reason / issues", "GSTR-2B Match (uploaded files)", "Source File", "Pages"]
PORTAL_MONEY = {"Invoice Value", "Taxable Value", "Integrated Tax", "Central Tax", "State/UT Tax", "Cess"}


def _portal_row(d: Document) -> dict:
    f = d.fields
    pos = f.get("place_of_supply") or ""
    issues = "; ".join(v["message"] for v in d.validations if v["level"] in ("ERROR", "WARNING"))
    return {"GSTIN of supplier": f.get("supplier_gstin", ""), "Trade/Legal name": f.get("supplier_legal_name", ""),
            "Invoice number": f.get("invoice_number", ""), "Invoice type": "Regular", "Invoice Date": _date(f.get("invoice_date")),
            "Invoice Value": f.get("invoice_total"), "Place of supply": pos.split(" (")[0] if pos else "",
            "Supply Attract Reverse Charge": "Y" if f.get("reverse_charge") == "Yes" else "N",
            "Taxable Value": f.get("taxable_value"), "Integrated Tax": f.get("igst"), "Central Tax": f.get("cgst"),
            "State/UT Tax": (f.get("sgst") or 0) + (f.get("utgst") or 0) or None, "Cess": f.get("cess"),
            "Status in this register": d.status, "Reason / issues": issues or "", "GSTR-2B Match (uploaded files)": d.gstr2b_match,
            "Source File": d.rel_path, "Pages": d.pages}


def _gstin_wise_sheets(wb: Workbook, docs: list[Document], final: list[Document], master) -> None:
    """One sheet per Caliber GSTIN in the GSTR-2B 'B2B' column order, so the final invoices can be ticked off
    against the GST portal.  Part 1 = Final (verified) invoices, supplier-wise with subtotals;
    part 2 = invoices billed to the same GSTIN that are NOT in Final yet (pending / duplicate) with the reason."""
    final_ids = {d.doc_id for d in final}
    gstins = sorted({d.caliber_gstin for d in docs if d.caliber_gstin} | set(getattr(master, "entries", {}) or {}))
    pos_idx = wb.sheetnames.index("Final_Tax_Invoices") + 1
    cols = PORTAL_COLS + EXTRA_PORTAL
    for g in gstins:
        ws = wb.create_sheet(f"GSTIN_{g}"[:31], pos_idx)
        pos_idx += 1
        mine = [d for d in docs if d.caliber_gstin == g and d.doc_type in INVOICE_TYPES]
        fin = sorted([d for d in mine if d.doc_id in final_ids],
                     key=lambda d: (d.fields.get("supplier_gstin", ""), d.fields.get("invoice_date", ""), d.fields.get("invoice_number", "")))
        rest = [d for d in mine if d.doc_id not in final_ids]
        entry = (master.lookup(g) if master else None) or {}
        ws.cell(row=1, column=1, value=f"Caliber GSTIN {g} - {state_name(g)} - {entry.get('verification_status', '')}").font = Font(bold=True, size=13)
        ws.cell(row=2, column=1, value="Columns follow the GSTR-2B B2B layout on the GST portal (Taxable, IGST, CGST, SGST, Cess). "
                "Part 1 = invoices in Final_Tax_Invoices; Part 2 = billed to this GSTIN but not final yet.").font = Font(italic=True, color="555555")
        r = 4

        def header(r):
            for c, h in enumerate(cols, 1):
                cell = ws.cell(row=r, column=c, value=h)
                cell.fill, cell.font = HEADER_FILL, HEADER_FONT
                cell.alignment = Alignment(wrap_text=True, vertical="center")
            ws.row_dimensions[r].height = 30

        def write(r, row, font=None, fill=None):
            for c, h in enumerate(cols, 1):
                v = row.get(h)
                cell = ws.cell(row=r, column=c, value=v)
                if h in PORTAL_MONEY:
                    cell.number_format = "#,##0.00"
                elif isinstance(v, datetime):
                    cell.number_format = "DD-MM-YYYY"
                elif h in ("GSTIN of supplier", "Invoice number"):
                    cell.number_format = "@"
                if font:
                    cell.font = font
                if fill:
                    cell.fill = fill

        def sums(ds):
            out = {}
            for h, k in (("Invoice Value", "invoice_total"), ("Taxable Value", "taxable_value"), ("Integrated Tax", "igst"),
                         ("Central Tax", "cgst"), ("Cess", "cess")):
                out[h] = round(sum(d.fields.get(k) or 0 for d in ds), 2)
            out["State/UT Tax"] = round(sum((d.fields.get("sgst") or 0) + (d.fields.get("utgst") or 0) for d in ds), 2)
            return out

        ws.cell(row=r, column=1, value=f"PART 1 - FINAL VERIFIED INVOICES ({len(fin)})").font = Font(bold=True, color="1F3A5F", size=12)
        r += 1
        header(r)
        first_data = r + 1
        r += 1
        by_sup = defaultdict(list)
        for d in fin:
            by_sup[d.fields.get("supplier_gstin", "")].append(d)
        for sup, ds in by_sup.items():
            for d in ds:
                write(r, _portal_row(d))
                r += 1
            if len(ds) > 1:
                write(r, dict(sums(ds), **{"GSTIN of supplier": sup, "Trade/Legal name": f"Subtotal ({len(ds)} invoices)"}),
                      font=TOTAL_FONT, fill=OK_FILL)
                r += 1
        tot = sums(fin)
        write(r, dict(tot, **{"GSTIN of supplier": "GRAND TOTAL", "Trade/Legal name": f"{len(fin)} invoices, {len(by_sup)} suppliers"}),
              font=TOTAL_FONT, fill=OK_FILL)
        ws.freeze_panes = ws.cell(row=first_data, column=4)
        ws.auto_filter.ref = f"A{first_data - 1}:{get_column_letter(len(cols))}{max(first_data, r - 1)}"
        r += 3
        ws.cell(row=r, column=1, value=f"PART 2 - BILLED TO {g} BUT NOT IN FINAL ({len(rest)}) - check the reason before claiming ITC").font = Font(bold=True, color="9C0006", size=12)
        r += 1
        header(r)
        r += 1
        for d in sorted(rest, key=lambda d: (d.status, d.fields.get("supplier_gstin", ""))):
            write(r, _portal_row(d), fill=DUP_FILL if d.status == ST_DUPLICATE else REVIEW_FILL)
            r += 1
        if not rest:
            ws.cell(row=r, column=1, value="None")
        widths = {"GSTIN of supplier": 18, "Trade/Legal name": 34, "Invoice number": 20, "Invoice type": 9, "Invoice Date": 12,
                  "Place of supply": 14, "Supply Attract Reverse Charge": 10, "Status in this register": 18,
                  "Reason / issues": 60, "GSTR-2B Match (uploaded files)": 22, "Source File": 45, "Pages": 7}
        for c, h in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(c)].width = widths.get(h, 14)

"""Excel / CSV extraction.

Three shapes are recognised:
  * GSTR-2B downloads (portal format: 'Read me' + B2B / B2BA / B2B-CDNR
    sheets with two header rows and merged cells) -> parsed into normalised
    register rows, which are used to cross-check PDF invoices;
  * generic invoice registers (a header row containing GSTIN + invoice
    number + value columns) -> register rows;
  * anything else (an invoice laid out in a sheet, a bank statement...) ->
    flattened to text so the normal classifier/extractor can handle it.
The source workbook is opened read-only and never written.
"""
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd

from models import PageText

_NUM_RE = re.compile(r"^-?[\d,]+(\.\d+)?$")


def _cell_text(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d-%m-%Y")
    if isinstance(v, float):
        return f"{v:.2f}" if v != int(v) else str(int(v))
    return str(v).strip()


def _num(v) -> float | None:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("₹", "").replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def _date(v) -> str:
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    s = str(v or "").strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d", "%d-%b-%Y", "%d-%b-%y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return ""


def _load_sheets(path: Path) -> dict[str, list[list[Any]]]:
    ext = path.suffix.lower()
    if ext == ".csv":
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                df = pd.read_csv(path, header=None, dtype=object, encoding=enc, keep_default_na=False)
                return {"CSV": df.values.tolist()}
            except UnicodeDecodeError:
                continue
        raise ValueError("CSV encoding not recognised")
    if ext == ".xls":
        xl = pd.read_excel(path, sheet_name=None, header=None, dtype=object, engine="xlrd")
        return {k: v.where(pd.notna(v), None).values.tolist() for k, v in xl.items()}
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheets = {}
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        sheets[ws.title] = rows
    wb.close()
    return sheets


# ---------------------------------------------------------------- GSTR-2B
_2B_COLS = {
    "gstin of supplier": "supplier_gstin", "trade/legal name": "supplier_name",
    "invoice number": "invoice_number", "note number": "invoice_number",
    "invoice type": "invoice_type", "note type": "note_type", "note supply type": "invoice_type",
    "invoice date": "invoice_date", "note date": "invoice_date",
    "invoice value(₹)": "invoice_value", "note value (₹)": "invoice_value",
    "place of supply": "place_of_supply", "supply attract reverse charge": "reverse_charge",
    "taxable value (₹)": "taxable_value", "integrated tax(₹)": "igst", "central tax(₹)": "cgst",
    "state/ut tax(₹)": "sgst", "cess(₹)": "cess", "gstr-1/iff/gstr-5 period": "period",
    "itc availability": "itc_available", "irn": "irn", "source": "source",
}


def _parse_2b_sheet(name: str, rows: list[list[Any]], owner_gstin: str) -> list[dict]:
    hdr_idx = next((i for i, r in enumerate(rows[:15]) if r and str(r[0] or "").strip().lower() == "gstin of supplier"), None)
    if hdr_idx is None:
        return []
    top, sub = rows[hdr_idx], rows[hdr_idx + 1] if hdr_idx + 1 < len(rows) else []
    colmap: dict[int, str] = {}
    width = max(len(top), len(sub))
    for c in range(width):
        for label in ((sub[c] if c < len(sub) else None), (top[c] if c < len(top) else None)):
            key = str(label or "").strip().lower()
            if key in _2B_COLS and c not in colmap:
                # the 2B CDNR sheet repeats tax headings for 'ITC reduction'; keep first
                if _2B_COLS[key] in colmap.values():
                    continue
                colmap[c] = _2B_COLS[key]
    out = []
    for r_i, row in enumerate(rows[hdr_idx + 2:], start=hdr_idx + 3):
        if not row or not row[0] or not re.match(r"^\d{2}[A-Z]{5}", str(row[0]).strip()):
            continue
        rec = {"sheet": name, "row": r_i, "customer_gstin": owner_gstin}
        for c, k in colmap.items():
            rec[k] = row[c] if c < len(row) else None
        for k in ("invoice_value", "taxable_value", "igst", "cgst", "sgst", "cess"):
            rec[k] = _num(rec.get(k))
        rec["invoice_date"] = _date(rec.get("invoice_date"))
        rec["invoice_number"] = str(rec.get("invoice_number") or "").strip()
        rec["supplier_gstin"] = str(rec.get("supplier_gstin") or "").strip().upper()
        rec["record_type"] = rec.get("note_type") or "Invoice"
        rec["amendment"] = name.upper().endswith("A") or "A(" in name.upper()
        out.append(rec)
    return out


def _is_gstr2b(sheets: dict) -> tuple[bool, str, str]:
    rm = sheets.get("Read me") or sheets.get("Read Me") or []
    flat = " ".join(_cell_text(c) for r in rm[:10] for c in (r or []))
    if "GSTR-2B" in flat.upper():
        g = re.search(r"\b\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z]Z[0-9A-Z]\b", flat)
        period = re.search(r"Tax Period\s+(\w+)", flat)
        fy = re.search(r"Financial Year\s+([\d-]+)", flat)
        return True, g.group(0) if g else "", f"{period.group(1) if period else ''} {fy.group(1) if fy else ''}".strip()
    return False, "", ""


# ------------------------------------------------------- generic registers
_GEN_HINTS = {
    "supplier_gstin": r"(supplier|vendor|party).*gstin|gstin.*(supplier|vendor|party)|^gstin$|gst no",
    "supplier_name": r"(supplier|vendor|party).*name|^name$|trade/legal",
    "invoice_number": r"invoice\s*(no|number|#)|bill\s*no|voucher\s*no",
    "invoice_date": r"invoice\s*date|bill\s*date|^date$",
    "taxable_value": r"taxable",
    "igst": r"igst|integrated",
    "cgst": r"cgst|central",
    "sgst": r"sgst|state/ut|utgst",
    "invoice_value": r"invoice\s*(value|total|amount)|total\s*(amount|value)|gross",
}


def _parse_generic_register(name: str, rows: list[list[Any]]) -> list[dict]:
    for h_i, row in enumerate(rows[:30]):
        labels = [str(c or "").strip().lower() for c in row]
        colmap = {}
        for key, pat in _GEN_HINTS.items():
            for c, lab in enumerate(labels):
                if lab and re.search(pat, lab) and c not in colmap:
                    colmap[c] = key
                    break
        keys = set(colmap.values())
        if {"invoice_number", "taxable_value"} <= keys and ({"supplier_gstin", "invoice_value"} & keys):
            out = []
            for r_i, r in enumerate(rows[h_i + 1:], start=h_i + 2):
                rec = {"sheet": name, "row": r_i}
                for c, k in colmap.items():
                    rec[k] = r[c] if c < len(r) else None
                if not rec.get("invoice_number"):
                    continue
                for k in ("invoice_value", "taxable_value", "igst", "cgst", "sgst"):
                    rec[k] = _num(rec.get(k))
                rec["invoice_date"] = _date(rec.get("invoice_date"))
                rec["invoice_number"] = _cell_text(rec.get("invoice_number"))
                rec["supplier_gstin"] = _cell_text(rec.get("supplier_gstin")).upper()
                rec["record_type"] = "Invoice"
                out.append(rec)
            return out
    return []


def _sheet_text(rows: list[list[Any]], limit: int = 400) -> str:
    lines = []
    for r in rows[:limit]:
        cells = [_cell_text(c) for c in (r or [])]
        cells = [c for c in cells if c]
        if cells:
            lines.append("  ".join(cells))
    return "\n".join(lines)


def extract_excel(path: Path) -> dict:
    """Return {'kind': 'gstr2b'|'register'|'sheet', 'owner_gstin', 'period',
    'register_rows': [...], 'pages': [PageText per sheet]}"""
    sheets = _load_sheets(path)
    is2b, owner, period = _is_gstr2b(sheets)
    if is2b:
        rows = []
        for name, data in sheets.items():
            if name.upper().startswith(("B2B", "ECO", "IMPG", "ISD")) and "REJECTED" not in name.upper():
                rows.extend(_parse_2b_sheet(name, data, owner))
        text = f"GSTR-2B register\nGSTIN {owner}\nTax period {period}\nRows {len(rows)}"
        return {"kind": "gstr2b", "owner_gstin": owner, "period": period, "register_rows": rows,
                "pages": [PageText(page_no=1, text=text, source="excel", views=[text])]}
    reg_rows = []
    for name, data in sheets.items():
        reg_rows.extend(_parse_generic_register(name, data))
    pages = []
    for i, (name, data) in enumerate(sheets.items(), start=1):
        t = _sheet_text(data)
        if t.strip():
            pages.append(PageText(page_no=i, text=f"[Sheet: {name}]\n{t}", source="excel", views=[f"[Sheet: {name}]\n{t}"]))
    kind = "register" if len(reg_rows) >= 2 else "sheet"
    return {"kind": kind, "owner_gstin": "", "period": "", "register_rows": reg_rows if kind == "register" else [],
            "pages": pages}

"""Build a self-contained HTML dashboard from a Caliber GST register workbook.

    python dashboard.py Caliber_GST_Register.xlsx -o site/index.html

The page needs no server: open it locally, host it on GitHub Pages, or attach it to an email.
Data (final invoices per Caliber GSTIN, pending reasons, duplicates) is embedded in the file.
"""
from __future__ import annotations

import argparse
import io
import json
import re
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / "dashboard_template.html"
STATES = {"29": "Karnataka", "33": "Tamil Nadu", "27": "Maharashtra", "07": "Delhi", "36": "Telangana", "09": "Uttar Pradesh",
          "06": "Haryana", "24": "Gujarat", "32": "Kerala", "37": "Andhra Pradesh"}


def _sheet(wb, name: str) -> list[dict]:
    """Rows of a register sheet as dicts (row 1 is a note, row 2 the header; the TOTAL row is skipped)."""
    if name not in wb.sheetnames:
        return []
    rows = list(wb[name].iter_rows(values_only=True))
    if len(rows) < 2:
        return []
    hdr = rows[1]
    out = []
    for r in rows[2:]:
        d = dict(zip(hdr, r))
        first = next(iter(d.values()), None)
        if first == "TOTAL" or not any(v not in (None, "") for v in d.values()):
            continue
        out.append(d)
    return out


def _date(v) -> str:
    return v.strftime("%d-%m-%Y") if isinstance(v, datetime) else (v or "")


def _num(v):
    return round(float(v), 2) if isinstance(v, (int, float)) else None


def _invoice(d: dict) -> dict:
    cal = d.get("Caliber Registration (State)") or ""
    pos = d.get("Place of Supply") or ""
    issues = d.get("Issues") or ""
    return {
        "caliber_gstin": cal.split(" ")[0] if cal else (d.get("Customer GSTIN") or ""),
        "supplier_gstin": d.get("Supplier GSTIN") or "", "supplier_name": d.get("Supplier Legal Name") or "",
        "invoice_number": str(d.get("Invoice Number") or ""), "invoice_date": _date(d.get("Invoice Date")),
        "invoice_value": _num(d.get("Invoice Total")), "pos": pos.split(" (")[0],
        "rcm": "Y" if d.get("Reverse Charge Indicator") == "Yes" else "N",
        "taxable": _num(d.get("Taxable Value")), "igst": _num(d.get("IGST")), "cgst": _num(d.get("CGST")),
        "sgst": (_num(d.get("SGST")) or 0) + (_num(d.get("UTGST")) or 0) or None, "cess": _num(d.get("Cess")),
        "status": d.get("Document Status") or "", "issues": issues.replace(" | ", "; "),
        "rules": [],
        "file": d.get("Source File Path / Relative Path") or "",
    }


def build_data(xlsx: bytes | Path, source_name: str = "") -> dict:
    wb = load_workbook(io.BytesIO(xlsx) if isinstance(xlsx, (bytes, bytearray)) else xlsx, read_only=True, data_only=True)
    final = sorted((_invoice(d) for d in _sheet(wb, "Final_Tax_Invoices")),
                   key=lambda r: (r["caliber_gstin"], r["supplier_gstin"], r["invoice_date"][6:] + r["invoice_date"][3:5] + r["invoice_date"][:2]))
    pending = [_invoice(d) for d in _sheet(wb, "Pending_Review")]
    labels = [(r"Taxable \+ taxes", "Total does not add up"), (r"No IGST/CGST", "Tax amount missing"),
              (r"Taxable value not found", "Taxable value missing"), (r"Invoice total not found", "Total missing"),
              (r"not been verified", "Caliber GSTIN not verified"), (r"buyer GSTIN not found", "Caliber GSTIN missing"),
              (r"Supplier GSTIN not found", "Supplier GSTIN missing"), (r"date not found", "Invoice date missing"),
              (r"number not found", "Invoice number missing"), (r"auto-corrected", "GSTIN read by OCR - verify"),
              (r"is printed with state", "GSTIN / state mismatch"), (r"CGST/SGST charged", "CGST/SGST on interstate"),
              (r"Bill of Supply", "Bill of Supply - no ITC"), (r"differs from state printed", "Supplier state mismatch"),
              (r"not a possible GST rate", "Implausible tax rate")]

    def label(msg: str) -> str:
        return next((lab for rx, lab in labels if re.search(rx, msg, re.I)), msg.strip()[:40])

    for p in pending:
        p["rules"] = [f"{lvl} {label(msg)}" for lvl, msg in re.findall(r"\[(ERROR|WARNING)\] ([^|;]+)", p["issues"])]
    for f in final:
        f.pop("rules", None)
    master = _sheet(wb, "Caliber_GST_Master")
    regs = [{"gstin": m["GSTIN"], "state": m.get("State Name") or STATES.get(str(m["GSTIN"])[:2], ""),
             "status": m.get("Verification Status") or ""} for m in master if m.get("GSTIN")]
    known = {r["gstin"] for r in regs}
    for g in sorted({r["caliber_gstin"] for r in final + pending if r["caliber_gstin"]} - known):
        regs.append({"gstin": g, "state": STATES.get(g[:2], ""), "status": "Not in master"})
    dups = [{"group": d.get("Group"), "level": (d.get("Level") or "").split(" - ")[-1], "number": d.get("Document Number"),
             "supplier": d.get("Supplier GSTIN"), "primary": (d.get("Primary Document") or "").split("/")[-1],
             "candidate": (d.get("Duplicate Candidate") or "").split("/")[-1], "action": d.get("Suggested Action")}
            for d in _sheet(wb, "Duplicate_Report")]
    # header facts from Executive_Summary (present from v1.1.0)
    facts = {}
    if "Executive_Summary" in wb.sheetnames:
        for r in wb["Executive_Summary"].iter_rows(values_only=True):
            if r and r[0]:
                facts[str(r[0]).strip()] = r[1]
    docreg = _sheet(wb, "Document_Register")
    return {
        "company": (master[0].get("Customer Legal Name") or "Caliber Inno Enterprise").title().replace("Private Limited", "Pvt Ltd")
        if master else "Caliber Inno Enterprise",
        "job_id": facts.get("Batch / job ID", ""), "version": facts.get("Application version", ""),
        "generated": facts.get("Generated", datetime.now().strftime("%d-%m-%Y %H:%M")),
        "files": facts.get("Files received", len({d.get("Relative Path") for d in docreg})),
        "documents": len(docreg), "registrations": regs, "final": final, "pending": pending, "duplicates": dups,
        "non_tax": len(_sheet(wb, "Non_Tax_Documents")), "source": source_name or "the register workbook",
    }


def render_html(data: dict, fragment: bool = False) -> str:
    body = TEMPLATE.read_text(encoding="utf-8")
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    body = body.replace("/*__DATA__*/null", payload, 1)
    if fragment:
        return body
    return ('<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n'
            '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
            '<meta name="robots" content="noindex, nofollow">\n' + body + "\n</html>\n")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("register", help="Caliber_GST_Register_*.xlsx exported by the app")
    ap.add_argument("-o", "--out", default="site/index.html", help="output HTML (default: site/index.html)")
    ap.add_argument("--fragment", action="store_true", help="omit <html>/<head> wrapper (for hosts that add their own)")
    a = ap.parse_args()
    src = Path(a.register)
    html = render_html(build_data(src, src.name), a.fragment)
    out = Path(a.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html, encoding="utf-8")
    print(f"dashboard written: {out} ({len(html) // 1024} KB)")


if __name__ == "__main__":
    main()

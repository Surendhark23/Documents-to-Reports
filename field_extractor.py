"""Normalised invoice field extraction.

Design notes (from the Jun-26 sample):
* Invoices come from ~25 different generators (Tally, Zoho, Vyapar, Swipe,
  Busy, Odoo, Excel, hand-made Word/PDF, phone photos).  Rather than one
  template per supplier, extraction is label-driven and tries several text
  *views* of the same page (native order, layout-sorted, OCR).
* Party roles are assigned with a small state machine over the lines: the
  most recent role label (Bill To / Buyer, Ship To / Consignee, Dispatch From,
  Supplier...) owns the GSTINs that follow it; two-column headers
  ("Billed to | Shipped to") hand out roles in order.  A GSTIN with no label
  in the top part of the page is the issuer's header GSTIN.
* Amounts are reconciled: the combination of taxable value, taxes,
  round-off and total that balances is preferred; if nothing balances, the
  labelled values are kept and validation reports the difference.
"""
from __future__ import annotations

import re
from collections import OrderedDict
from datetime import datetime
from typing import Optional

from gst_validator import (GSTIN_RE, GSTIN_LOOSE_RE, correct_gstin, gstin_valid,
                           state_code_from_name, state_name)
from models import (INWARD, OUTWARD, THIRD_PARTY, CONSIGNEE_ONLY, DIRECTION_UNKNOWN,
                    CREDIT_NOTE, DEBIT_NOTE, PROFORMA, QUOTATION, DELIVERY_CHALLAN, PURCHASE_ORDER, SALES_ORDER)

# ------------------------------------------------------------------ numbers
AMOUNT_RE = re.compile(r"(?<![\w.])(?:₹|ī|Rs\.?|INR|`)?\s*(\(-\)\s*|-\s*)?((?:\d{1,3}(?:,\d{2,3})+|\d+)(?:\.\d{1,3})?)(?![\d,]*[A-Za-z/])")
RATE_SET = (0.1, 0.25, 1.0, 1.5, 2.5, 3.0, 5.0, 6.0, 7.5, 9.0, 12.0, 14.0, 18.0, 28.0, 40.0)


def parse_amount(s: str) -> Optional[float]:
    if s is None:
        return None
    m = AMOUNT_RE.search(str(s))
    if not m:
        return None
    v = float(m.group(2).replace(",", ""))
    return -v if m.group(1) else v


# ------------------------------------------------------------------ dates
_MONTHS = "jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec"
DATE_PATTERNS = [
    (re.compile(rf"\b(\d{{1,2}})[-/ .]({_MONTHS})[a-z]*[-/ .,]\s*(\d{{2,4}})\b", re.I), "dmy_text"),
    (re.compile(rf"\b({_MONTHS})[a-z]*\.?\s+(\d{{1,2}}),?\s+(\d{{4}})\b", re.I), "mdy_text"),
    (re.compile(r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b"), "ymd"),
    (re.compile(r"\b(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})\b"), "dmy"),
]
_MON = {m: i for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}


def parse_date(s: str) -> tuple[str, str]:
    """Return (iso_date or '', raw_match).  Invalid dates return ('', raw)."""
    if not s:
        return "", ""
    first_invalid = ""
    for rx, kind in DATE_PATTERNS:
        # audit fix E2: try every match - '87/26-27  19.09.26' used to stop at the
        # invoice-number-shaped '87/26-27' and report an invalid date
        for m in rx.finditer(s):
            raw = m.group(0)
            try:
                if kind == "dmy_text":
                    d, mon, y = int(m.group(1)), _MON[m.group(2).lower()[:3]], int(m.group(3))
                elif kind == "mdy_text":
                    mon, d, y = _MON[m.group(1).lower()[:3]], int(m.group(2)), int(m.group(3))
                elif kind == "ymd":
                    y, mon, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
                else:
                    d, mon, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if y < 100:
                    y += 2000
                return datetime(y, mon, d).strftime("%Y-%m-%d"), raw
            except (ValueError, KeyError):
                first_invalid = first_invalid or raw
    return "", first_invalid


# ------------------------------------------------------------------ labels
ROLE_PATTERNS = OrderedDict([
    ("BILL", re.compile(r"buyer\s*\(?\s*bill\s*to|bill(?:ed)?\s*to\b|customer\s*details?\b|invoicing\s*address|"
                        r"details\s*of\s*receiver|^\s*receiver\b|recipient\s*details|^\s*sold\s*to|quotation\s*to|"
                        r"^\s*to[\s:.,]*$|^\s*to\b\s*[:.]|^customer:|^m/s\b|^\s*buyer\s*(?:$|:|\()|details\s*of\s*buyer", re.I)),
    ("SHIP", re.compile(r"consignee|ship(?:ped)?\s*to|shipping\s*address|deliver(?:y)?\s*(?:to|address)|place\s*of\s*delivery", re.I)),
    ("DISPATCH", re.compile(r"dispatch(?:ed)?\s*from|ship\s*from", re.I)),
    ("SUPPLIER", re.compile(r"^\s*(?:supplier|seller|vendor)\b(?!\s*code)|details\s*of\s*supplier|invoicing\s*party|goods\s*supplier|^\s*from\s*:", re.I)),
])
_ROLE_WINDOW = 12

COMPANY_HINT = re.compile(r"\b(LTD|LIMITED|LLP|PVT|PRIVATE|ENTERPRISES?|INTERNATIONAL|TRADERS|SOLUTIONS|PRINTS|"
                          r"INDUSTRIES|CORPORATION|COMPANY|CO\.|GIFTS|APPLIANCES|APPARELS|EXPRESS|LOGISTICS|EVENTS|"
                          r"WORKS|PACKAGING|CREATIONS|INC|AGENCY|STORE|MART|SYSTEMS|TECH|DIARIES|FASHIONS)\b", re.I)
NOT_NAME = re.compile(r"(?i)^(tax\s*invoice|invoice|irn|ack|e-?invoice|original|duplicate|triplicate|gstin|state|"
                      r"address|phone|mobile|email|e-mail|contact|pan|cin|date|dated|place|buyer|consignee|bill|ship|"
                      r"customer|m/s|to\b|name\s*:|udyam|msme|website|www|ph\b|tel\b|credit note|debit note|"
                      r"performa|proforma|quotation|purchase order|delivery challan|sr\.|s\.?no|sl\b|code|"
                      r"invoice no|\(original|copy|kind attn|subject|dear)")


def _lines(text: str) -> list[str]:
    return [l.rstrip() for l in (text or "").splitlines()]


def _first_chunk(s: str) -> str:
    """Value portion of 'label : value     other column' text."""
    s = s.strip().lstrip(":#-–.+*•|").strip()
    s = re.split(r"\s{3,}|\t|\|", s)[0].strip()
    # OCR joins columns with single spaces: cut at the next label word
    s = re.split(r"\s+(?=(?:GSTIN|GST\s*No|Dated?|PAN|State|Ack|IRN|Place\s+of|e-?Way|Vehicle|Station|Transport|"
                 r"Due\s*Date|Mode\s*of|Mode/|Buyer|Ref(?:erence)?\s*No|Delivery\s*Note|Challan\s*No)\b)", s, flags=re.I)[0]
    return s.strip()


# ------------------------------------------------------------------ GSTINs
def find_gstins(text: str, known: set[str]) -> list[dict]:
    """GSTIN occurrences with an assigned role, in reading order."""
    lines = _lines(text)
    occ = []
    ctx, ctx_line, queue = None, -999, []
    pending_bill_line = -999
    for i, line in enumerate(lines):
        up = line.upper()
        labels = [r for r, rx in ROLE_PATTERNS.items() if rx.search(line)]
        if "BILL" in labels and "SHIP" in labels:
            queue, ctx, ctx_line = ["BILL", "SHIP"], "BILL", i
        elif labels:
            lab = labels[0]
            if lab == "SHIP" and i - pending_bill_line <= 2 and not occ_since(occ, pending_bill_line):
                queue = ["BILL", "SHIP"]           # 'Billed to' / 'Shipped to' on adjacent lines
            else:
                queue = []
            if lab == "BILL":
                pending_bill_line = i
            ctx, ctx_line = lab, i
        found = []
        for m in GSTIN_RE.finditer(up):
            found.append((m.group(0), m.group(0), ""))
        if not found:
            for m in GSTIN_LOOSE_RE.finditer(up):
                g, how = correct_gstin(m.group(0), known)
                if how == "corrected":
                    found.append((g, m.group(0), "corrected"))
        for g, raw, how in found:
            if queue:
                role = queue.pop(0)
            elif ctx and i - ctx_line <= _ROLE_WINDOW:
                role = ctx
            else:
                role = "HEADER"
            occ.append({"gstin": g, "raw": raw, "how": how, "line": i, "role": role,
                        "valid": gstin_valid(g)})
    return occ


def occ_since(occ, line_no) -> bool:
    return any(o["line"] > line_no for o in occ)


# ------------------------------------------------------------------ labelled values
def _segments(line: str):
    """(start, end, text) of column cells - text runs separated by 2+ spaces."""
    return [(m.start(), m.end(), m.group(0)) for m in re.finditer(r"\S+(?: \S+)*", line)]


def _value_below_column(lines: list[str], i: int, col: int, validator, lookahead: int = 3) -> Optional[str]:
    """Audit fix E1: Tally/Busy layouts print the label in a right-hand column
    ('Consignor   Invoice No.   Dated') and the value directly below it
    ('VIRAJ INDUSTRIES   VI 87/26-27   19.09.26').  Take the cell under the label."""
    for j in range(i + 1, min(len(lines), i + 1 + lookahead)):
        near = []
        for start, end, txt in _segments(lines[j]):
            dist = 0 if start <= col <= end else min(abs(start - col), abs(end - col))
            if dist <= 15:
                near.append((dist, txt))
        for _, txt in sorted(near):
            cand = _first_chunk(txt)
            if cand and validator(cand):
                return cand
    return None


def _value_after(lines: list[str], i: int, m_end: int, validator, lookahead: int = 3, m_start: int = 0) -> Optional[str]:
    same = _first_chunk(lines[i][m_end:])
    if same and validator(same):
        return same
    if m_start >= 15 and lines[i][:m_start].strip() and re.search(r"\s{3,}$", lines[i][:m_start]):
        # label sits in a right-hand column with other text to its left
        col = _value_below_column(lines, i, m_start, validator, min(lookahead, 2))
        if col:
            return col
    for j in range(i + 1, min(len(lines), i + 1 + lookahead)):
        cand = _first_chunk(lines[j])
        if not cand or cand in (":", "-"):
            continue
        if validator(cand):
            return cand
        if re.search(r"[A-Za-z]{3,}", cand) and not validator(cand):
            # another label: keep looking only one more line (Tally layouts)
            continue
    return None


def find_labelled(text: str, label_rx: str, validator, lookahead: int = 3, flags=re.I,
                  same_line_only: bool = False) -> Optional[str]:
    lines = _lines(text)
    rx = re.compile(label_rx, flags)
    for i, line in enumerate(lines):
        for m in rx.finditer(line):
            if same_line_only:
                same = _first_chunk(line[m.end():])
                if same and validator(same):
                    return same
                continue
            v = _value_after(lines, i, m.end(), validator, lookahead, m.start())
            if v:
                return v
    return None


def _valid_docno(v: str) -> bool:
    v = v.strip()
    if not (1 <= len(v) <= 30) or not re.search(r"\d", v):
        return False
    iso, raw = parse_date(v)
    # audit fix E1: only a *real* calendar date disqualifies the value.  Numbers such as
    # '87/26-27', '26-27/2540' or '26-27/046' look date-like but are invoice numbers.
    if iso and len(raw) >= len(v) - 2:
        return False
    if re.search(r"\s+dt\.?\s", v, re.I):        # 'Reference No. & Date' cell: '2540 dt. 18-Sep-26'
        return False
    if re.search(r"(?i)\b(dated|date|e-?way|ack|irn|gstin|state|code|page)\b", v):
        return False
    if re.fullmatch(r"[\d,]+\.\d{2}", v):     # an amount
        return False
    if " " in v and re.search(r"[A-Za-z]{5,}", v):   # address / sentence, not a number
        return False
    if re.fullmatch(r"(?:\d\s?){12}", v):            # e-way bill number
        return False
    if re.match(r"(?i)(no\.?\s*\d+,|plot|door|#\d)", v):
        return False
    return True


DOCNO_LABELS = {
    "credit": r"Credit\s*Note\s*(?:No|Number|#)\.?\s*:?",
    "debit": r"Debit\s*Note\s*(?:No|Number|#)\.?\s*:?",
    "invoice": r"(?<!Against )(?<!Original )(?:Tax\s*)?Invoice\s*(?:No|Number|#)\.?\s*:?|Invoice\s*#\s*:?|\bInv\.?\s*No\.?\s*:?|"
               r"(?<![Ww]ay )(?<![Ww]ay)(?<!E-way )Bill\s*(?:No|Number)\.?\s*:?|Voucher\s*No\.?\s*:?",
    "proforma": r"Proforma\s*(?:Invoice\s*)?No\.?\s*:?|P\.?I\.?\s*No\.?\s*:?",
    "quotation": r"Quotation\s*(?:No\.?|#)\s*:?|Quote\s*(?:No\.?|#)\s*:?",
    "challan": r"(?<!BILLING )Challan\s*No\.?\s*:?",
    "order": r"Purchase\s*Order\s*(?:No\.?|Number:?)|PO\s*Number\s*:?|Order\s*Number\s*:?",
    "secondary": r"(?<!Buyer.s )(?<!Buyer's )Reference\s*:|(?<!Buyer.s )(?<!Buyer's )(?<!Customer )Order\s*No\s*:",
}


def find_doc_number_views(views: list[str], kind_hint: str = "") -> tuple[Optional[str], int]:
    """Prefer a value printed on the same line as its label in any view
    (scrambled PDFs put labels and values far apart in native order)."""
    for idx, v in enumerate(views):
        val = find_doc_number(v, kind_hint)
        if val:
            return val, idx
    for idx, v in enumerate(views):
        val = find_doc_number_secondary(v)
        if val:
            return val, idx
    return None, -1


_HINT_LABELS = {CREDIT_NOTE: "credit", DEBIT_NOTE: "debit", PROFORMA: "proforma", QUOTATION: "quotation",
                DELIVERY_CHALLAN: "challan", PURCHASE_ORDER: "order", SALES_ORDER: "order"}


def _docno_label_order(kind_hint: str) -> list[str]:
    first = _HINT_LABELS.get(kind_hint)
    order = [DOCNO_LABELS[first]] if first else []
    for k in ("invoice", "proforma", "quotation", "challan", "order"):
        if k != first:
            order.append(DOCNO_LABELS[k])
    return order


def find_doc_number_secondary(text: str) -> Optional[str]:
    v = find_labelled(text, DOCNO_LABELS["secondary"], _valid_docno, lookahead=2)
    return v.strip().rstrip(".,") if v else None


_DOCNO_TOKEN = re.compile(r"^(?=.*\d)(?=.*[A-Za-z/\-])[A-Za-z0-9][A-Za-z0-9/\-]{3,24}$")


def _docno_token_scan(text: str, label_rx: str) -> Optional[str]:
    """OCR fallback: label found but values are smeared across columns -
    look for an invoice-number-shaped token in the next lines."""
    lines = _lines(text)
    rx = re.compile(label_rx, re.I)
    for i, line in enumerate(lines):
        lm = rx.search(line)
        if lm and not line[lm.end():].strip():
            # audit fix: OCR line ending in 'Invoice No.' with the value on the next line between
            # address words and the date ('... Sapphire MHPL/26-27/1351 24-Sep-26')
            for j in range(i + 1, min(len(lines), i + 3)):
                for tm in re.finditer(r"(\S+)\s+(\d{1,2}[-/.][A-Za-z]{3}[-/.]\d{2,4}|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4})\b", lines[j]):
                    tok = tm.group(1).strip(".,;:()")
                    if _DOCNO_TOKEN.match(tok) and _valid_docno(tok) and parse_date(tm.group(2))[0] \
                            and not GSTIN_RE.fullmatch(tok.upper()) and not parse_date(tok)[0]:
                        return tok
        # only for OCR-merged header rows such as 'Invoice No.  e-Way Bill No.  Dated'
        if not lm or not re.search(r"(?i)e-?way|dated|\bdate\b", line):
            continue
        for j in range(i, min(len(lines), i + 4)):
            for raw_tok in re.split(r"[\s|]+", lines[j]):
                if raw_tok.endswith(","):
                    continue          # address fragment such as '139/102-103,'
                tok = raw_tok.strip(".,;:()")
                if _DOCNO_TOKEN.match(tok) and _valid_docno(tok) and not GSTIN_RE.fullmatch(tok.upper()) \
                        and not parse_date(tok)[1] and not re.search(r"(?i)^(no|gstin|code|www)", tok):
                    return tok
    return None


def find_doc_number(text: str, kind_hint: str = "") -> Optional[str]:
    m = re.search(r"(?m)^\s*#\s*\n\s*:\s*([A-Z]{1,6}-?\d[\w/\-]*)", text)
    if m and _valid_docno(m.group(1)):
        return m.group(1)
    m = re.search(r"(?i:Tax\s*Invoice)\s+((?:INV|BILL)[\w/\-]*\d[\w/\-]*)", text)
    if m and _valid_docno(m.group(1)):
        return m.group(1)
    for rx in _docno_label_order(kind_hint):
        v = find_labelled(text, rx, _valid_docno, lookahead=4)
        if v:
            return v.strip().rstrip(".,")
    m = re.search(r"(?im)^\s*No\s*:\s*(\S+)", text)  # 'Invoice Details:\nNo: 112'
    if m and _valid_docno(m.group(1)):
        return m.group(1)
    return _docno_token_scan(text, DOCNO_LABELS["invoice"])


DATE_LABEL = (r"(?<!Due )(?<!Ack )(?<!PO )(?<!P\.O\. )(?<!Order )(?<!Against Invoice\n)(?<!Bill )(?:Invoice|Credit\s*Note|"
              r"Debit\s*Note|Bill|Challan|Proforma|Quotation|Purchase\s*Order|Inv\.?)\s*Date\s*:?|\bDated\s*:?|^\s*Date\s*:?|\bDate\s*:|"
              r"\bDate\s*&\s*Time\s*:?")


def find_date(text: str) -> tuple[str, str]:
    lines = _lines(text)
    invalid = None
    rx = re.compile(DATE_LABEL, re.I | re.M)
    bad_ctx = re.compile(r"(?i)(due|ack|p\.?o\.?|order|against|delivery note|e-?way|eway|buyer|irn)\s*date")
    for i, line in enumerate(lines):
        for m in rx.finditer(line):
            pre = line[max(0, m.start() - 15):m.end()]
            if bad_ctx.search(pre):
                continue
            prev = lines[i - 1] if i else ""
            if re.search(r"(?i)against invoice\s*$|^\s*against", prev + " " + line[:m.start()]):
                continue
            for j in range(i, min(len(lines), i + 4)):
                seg = line[m.end():] if j == i else lines[j]
                if j > i and re.search(r"(?i)(due|ack|po|order)\s*date", seg):
                    break
                if j == i and re.search(r"\b\d{1,2}[-/.]\s*$", seg):
                    # audit fix E2: date wrapped onto a later line ('23-' ... 'Jul-2026 19:20')
                    for k in range(i + 1, min(len(lines), i + 4)):
                        tail = re.match(rf"\s*((?:{_MONTHS})[a-z]*[-/ .]\s*\d{{2,4}}|\d{{1,2}}[-/.]\d{{2,4}})\b",
                                        lines[k][max(0, m.start() - 12):], re.I)
                        if tail:
                            iso2, raw2 = parse_date(seg.strip() + tail.group(1))
                            if iso2:
                                return iso2, raw2
                iso, raw = parse_date(seg)
                if iso:
                    return iso, raw
                if raw:
                    invalid = invalid or (iso, raw)   # keep looking for a real date
    # fallback: first plausible date in the top half that is not in a bad context
    for i, line in enumerate(lines[: max(10, len(lines) // 2)]):
        if bad_ctx.search(line) or (i and bad_ctx.search(lines[i - 1])):
            continue
        iso, raw = parse_date(line)
        if iso:
            return iso, raw
    return invalid or ("", "")


# ------------------------------------------------------------------ amounts
_CUR = r"(?:₹|ī|Rs\.?|INR|`)"
_AMT_LABELS = {
    "taxable_value": r"Taxable\s*Amount|Total\s*Taxable(?:\s*Value)?|Untaxed\s*Amount|Sub\s*-?\s*Total|Basic\s*Amount|"
                     r"(?:Amount|Amt\.?)\s*Before\s*Tax|Assessable\s*Value|Taxable\s*Value|Total\s*Value\s*of\s*Goods",
    "igst": r"(?:Add\s*:?\s*)?(?:OUTPUT\s*)?IGST\s*(?:\(INR\)|Amt|AMT)?\s*:?(?:\s*\d{0,2}\s*\(?\s*@?\s*[\d.]+\s*%\s*\)?)?|Integrated\s*Tax",
    "cgst": r"(?:Add\s*:?\s*)?(?:OUTPUT\s*)?CGST(?:\s*@?\s*\(?[\d.]+\s*%\)?)?\s*:?|Central\s*Tax",
    "sgst": r"(?:Add\s*:?\s*)?(?:OUTPUT\s*)?SGST(?!/UTGST\s*$)(?:\s*@?\s*\(?[\d.]+\s*%\)?)?\s*:?|State\s*Tax",
    "utgst": r"(?:Add\s*:?\s*)?UTGST(?:\s*@?\s*[\d.]+\s*%)?",
    "cess": r"(?:Add\s*:?\s*)?\bCess\b(?:\s*@?\s*[\d.]+\s*%)?",
    "total_tax": r"Total\s*Tax(?:\s*Amount)?(?!\s*in\s*words)(?:\s*\(₹\))?\s*:?|Total\s*GST\b",
    "round_off": r"Round(?:ed|ing)?\s*-?\s*off(?:\s*to)?|\bR\.?\s*Off\b|ROUND\s*[+-]|Rounding\s*O\b",
    "discount": r"(?:Less\s*:?\s*)?Discount(?:\s*Amt)?|Disc\s*Amt\s*\(-\)",
    "other_charges": r"(?:Add\s*:?\s*)?(?:Packing|Freight|Forwarding|Labour|Pickup|Shipping|Fuel)\s*(?:&\s*\w+\s*)?(?:charges?)\s*(?:@[\d.]+%)?",
    "invoice_total": r"(?:Total\s*)?(?:Amount|Amt\.?)\s*After\s*Tax|Grand\s*Total|Invoice\s*(?:Total|Value|Amount)(?!\s*in)|Total\s*Amount\s*\(INR\)|"
                     r"Total\s*\(INR\)|Net\s*Amount|Net\s*Value|Total\s*Invoice\s*Value|Total\s*Inv\.?\s*Amt|Total\s*Amount(?!\s*in)|Amount\s*Payable|"
                     r"Total\s*Value(?!\s*of)|Bill\s*Amount|Balance\s*Due|Current\s*Balance|^\s*Total\s*:?|^\s*TOTAL\s*:?",
}
_STRONG_TOTAL = re.compile(r"(?i)after\s*tax|grand|invoice\s*(total|value|amount)|total\s*amount|net\s*amount|payable|total\s*invoice|total\s*\(inr\)")


def amounts_in(line: str) -> list[tuple[float, bool]]:
    """[(value, currency_marked)] - rates (followed by %) and quantities with units excluded."""
    out = []
    for m in re.finditer(rf"({_CUR}\s*)?(\(-\)\s*|-\s*)?((?:\d{{1,3}}(?:,\d{{2,3}})+|\d+)(?:\.\d{{1,3}})?)(?![\d,])", line):
        tail = line[m.end():m.end() + 6]
        if re.match(r"\s*%", tail):
            continue
        if re.match(r"\s*(?:pcs|nos|no\b|mtr|kg|box|units?|qty|pc\b|sets?)", tail, re.I):
            continue
        head = line[max(0, m.start() - 1):m.start()]
        if head and (head.isalnum() or head in "/") and not m.group(1):
            continue
        try:
            v = float(m.group(3).replace(",", ""))
        except ValueError:
            continue
        out.append((-v if m.group(2) else v, bool(m.group(1))))
    return out


def _amount_candidates(text: str) -> dict[str, list[tuple[float, float]]]:
    """{field: [(value, priority)]} lower priority = more trusted."""
    lines = _lines(text)
    out: dict[str, list] = {k: [] for k in _AMT_LABELS}
    out["tax_lines"] = []
    for key, pat in _AMT_LABELS.items():
        rx = re.compile(pat, re.I | re.M)
        for i, line in enumerate(lines):
            for m in rx.finditer(line):
                if key == "taxable_value" and re.fullmatch(r"(?i)\s*taxable\s*value\s*", line[m.start():]) \
                        and i + 1 < len(lines) and re.search(r"[A-Za-z]{3,}", lines[i + 1]):
                    continue  # column header
                if key in ("igst", "cgst", "sgst") and re.search(r"(?i)total|amount\s*$|rate\s*$", line[:m.start()]):
                    continue
                rest = line[m.end():]
                vals = amounts_in(rest)
                j = i + 1
                while not vals and j < min(len(lines), i + 4):
                    nxt = lines[j].strip()
                    if nxt in ("", ":", "%", "₹", "ī", "(-)", "-"):
                        j += 1
                        continue
                    if re.search(r"[A-Za-z]{3,}", nxt) and not re.match(rf"(?i)^\s*{_CUR}", nxt):
                        break
                    vals = amounts_in(nxt)
                    j += 1
                if not vals:
                    continue
                plain_total = key == "invoice_total" and not _STRONG_TOTAL.search(m.group(0))
                if key == "invoice_total":
                    marked = [v for v, c in vals if c]
                    base = 0 if not plain_total else (1 if marked else 3)
                    if re.search(r"(?i)balance|payment\s*made", m.group(0)):
                        # audit fix E5: 'Balance Due 0.00' is what is left to pay, not the invoice value
                        vals = [(v, c) for v, c in vals if v > 0]
                        marked = [v for v in marked if v > 0]
                        base = 4
                        if not vals:
                            continue
                    ordered = (marked or []) + [v for v, c in reversed(vals) if v not in marked]
                    for n, v in enumerate(ordered[:4]):
                        out[key].append((v, base + n * 0.2))
                    # numbers after a 'Total' row of a tax summary are taxable/tax candidates too
                    if plain_total and len(vals) >= 2:
                        for v, _ in vals[:4]:
                            out["taxable_value"].append((v, 2.5))
                            out["total_tax"].append((v, 2.5))
                    continue
                v = vals[0][0]
                if key == "round_off" and "(-)" in (rest + (lines[i + 1] if i + 1 < len(lines) else ""))[:24]:
                    v = -abs(v)
                if key in ("igst", "cgst", "sgst", "utgst"):
                    rm_ = re.search(r"([\d.]+)\s*%", m.group(0))
                    out["tax_lines"].append((key, rm_.group(1) if rm_ else "", i, v))
                prio = 0 if key != "taxable_value" or "sub" not in m.group(0).lower() else 1.5
                out[key].append((v, prio + len(out[key]) * 0.1))
    for k, vals in tax_table_candidates(text).items():
        key = "taxable_value" if k == "tv" else k
        for v in vals:
            out[key].insert(0, (v, -0.1))     # column-aligned summary row: most trusted
    # every amount printed on the page: a computed total (taxable + tax + round-off) that is
    # printed somewhere is a valid total candidate even without a recognised label (audit fix E5)
    out["printed"] = sorted({round(v, 2) for line in lines for v, _ in amounts_in(line) if v > 0})
    # stand-alone amount lines right before the first tax line (template sub-totals)
    for i, line in enumerate(lines):
        if re.match(r"(?i)\s*(IGST|CGST)\s*\(", line) and i > 0:
            vals = amounts_in(lines[i - 1])
            if len(vals) == 1 and re.fullmatch(r"\s*[\d,]+\.\d{2}\s*", lines[i - 1]):
                out["taxable_value"].append((vals[0][0], 1.0))
            break
    return out


_RATE_VALUES = set(RATE_SET) | {r / 2 for r in RATE_SET}


def _header_kinds(header: str) -> list[tuple[int, int, str]]:
    """(start, end, kind) for each word of a tax-summary header row.
    kind: tv | igst | cgst | sgst | utgst | rate | other."""
    out, last_tax, prev = [], "", ""
    for m in re.finditer(r"\S+", header):
        w = m.group(0).lower().strip(":()")
        kind = "other"
        tax = next((t for t in ("igst", "cgst", "sgst", "utgst") if t in w), "")
        if w.startswith("taxable"):
            kind = "tv"
        elif w == "value" and prev == "tv":
            kind = "tv"
        elif tax and "%" in w:
            kind, last_tax = "rate", tax
        elif tax:
            kind, last_tax = tax, tax
        elif w in ("%", "rate", "rate%"):
            kind = "rate"
        elif w in ("amt", "amount", "amt.") and last_tax:
            kind = last_tax
        out.append((m.start(), m.end(), kind))
        prev = kind
    return out


def tax_table_candidates(text: str) -> dict[str, list[float]]:
    """Audit fix E7: read a column-aligned tax summary ('Taxable  CGST  SGST' header with the
    values in the row(s) below) by column position.  Used for layouts where the tax amounts
    sit under header labels instead of next to them (Caliber-2 / J00089, C.pdf / C686)."""
    lines = _lines(text)
    found: dict[str, list[float]] = {}
    for i, header in enumerate(lines):
        if not (re.search(r"(?i)taxable", header) and re.search(r"(?i)\b(?:IGST|CGST|SGST|UTGST)", header)
                and re.search(r"\S\s{2,}\S", header)):
            continue
        cols = _header_kinds(header)
        if not any(k == "tv" for _, _, k in cols) or not any(k in ("igst", "cgst", "sgst") for _, _, k in cols):
            continue
        rows = []
        for j in range(i + 1, min(len(lines), i + 8)):
            row = lines[j]
            nums = [(m.start(), m.end(), m.group(1)) for m in
                    re.finditer(r"(?<![\w/.-])(\d{1,3}(?:,\d{2,3})+(?:\.\d{1,3})?|\d+\.\d{1,3}|\d+)(?![\w/-])", row)]
            if not nums:
                if rows and re.search(r"[A-Za-z]{4,}", row):
                    break
                continue
            rows.append((row, nums))
        parsed = []
        # each amount column owns the space up to the next amount column (numbers are usually
        # right-aligned under left-aligned headings); rate columns hold no amounts
        amt_cols = [(cs, kind) for cs, ce, kind in cols if kind != "rate"]
        for row, nums in rows:
            got: dict[str, float] = {}
            for a, b, raw in nums:
                v = float(raw.replace(",", ""))
                if (v in _RATE_VALUES and v <= 40) or re.match(r"\s*%", row[b:b + 2]):
                    continue        # a rate (9, 18%, 2.5)
                anchor = a + 3
                if not amt_cols or anchor < amt_cols[0][0] - 6:
                    continue
                kind = [k for cs, k in amt_cols if cs <= max(anchor, amt_cols[0][0])][-1]
                if kind != "other" and kind not in got:
                    got[kind] = v
            if "tv" in got and any(k in got for k in ("igst", "cgst", "sgst")):
                parsed.append((bool(re.search(r"(?i)\btotal\b", row)), got))
        # a 'Total' row of the summary wins; otherwise only an unambiguous single row is used
        use = [g for is_total, g in parsed if is_total][:1] or ([parsed[0][1]] if len(parsed) == 1 else [])
        for got in use:
            for k, v in got.items():
                found.setdefault(k, []).append(v)
    return found


def tax_triplets(text: str) -> list[tuple[float, float, float]]:
    """(taxable, rate, tax) triples where tax == taxable * rate (rounded)."""
    toks = []
    for line in _lines(text):
        for m in re.finditer(r"(?<![\w/.-])(\d{1,3}(?:,\d{2,3})+(?:\.\d{1,3})?|\d+(?:\.\d{1,3})?)(\s*%)?", line):
            try:
                toks.append((float(m.group(1).replace(",", "")), bool(m.group(2))))
            except ValueError:
                pass
    found = []
    n = len(toks)
    for i in range(n):
        a = toks[i][0]
        if a < 1:
            continue
        for j in range(i + 1, min(n, i + 4)):
            r = toks[j][0]
            if r not in RATE_SET:
                continue
            for k in range(j + 1, min(n, j + 3)):
                t = toks[k][0]
                if t <= 0 or t >= a:
                    continue
                if a >= 100 and t >= 5 and abs(a * r / 100 - t) <= max(0.51, t * 0.002):
                    found.append((round(a, 2), r, round(t, 2)))
        if a in RATE_SET:
            # rate before or after both amounts: '5% 20,000.00 1,000.00' / '1,000.00 5% 20,000.00'
            for j in range(max(0, i - 3), min(n, i + 5)):
                for k in range(max(0, i - 3), min(n, i + 6)):
                    if len({i, j, k}) < 3:
                        continue
                    b, t = toks[j][0], toks[k][0]
                    if 5 <= t < b and b >= 100 and abs(b * a / 100 - t) <= max(0.51, t * 0.002) and toks[j][0] not in RATE_SET:
                        found.append((round(b, 2), a, round(t, 2)))
    return list(OrderedDict.fromkeys(found))


def _uniq(items):
    seen, out = set(), []
    for v, p in sorted(items, key=lambda x: x[1]):
        k = round(v, 2)
        if k not in seen:
            seen.add(k)
            out.append((v, p))
    return out


def reconcile(c: dict, triplets, tol: float, intra_hint: bool) -> tuple[dict, bool]:
    """Pick one value per amount field such that the invoice balances."""
    taxables = list(c.get("taxable_value", []))
    if triplets:
        taxables.append((round(sum(a for a, r, t in triplets), 2), 0.5))
    # a zero taxable value / zero invoice total is a blank column, never the invoice value
    taxables = [t for t in _uniq(taxables) if abs(t[0]) >= 0.005][:10]
    totals = [t for t in _uniq(c.get("invoice_total", [])) if abs(t[0]) >= 0.005][:10]
    ro_opts = [(0.0, 0.3)] + [(v, 0) for v, _ in _uniq(c.get("round_off", []))[:2]]
    ocs = _uniq(c.get("other_charges", []))
    oc_opts = [(0.0, 0)] + [(v, 0.5) for v, _ in ocs[:1]] + ([(round(sum(v for v, _ in ocs), 2), 0.6)] if len(ocs) > 1 else [])
    disc = c["discount"][0][0] if c.get("discount") else 0.0
    printed = set(c.get("printed", []))
    cess = c["cess"][0][0] if c.get("cess") else None

    variants = []   # (dict, prio)
    by_rate = {}
    for key, rate, line, v in c.get("tax_lines", []):
        by_rate.setdefault(key, OrderedDict())
        by_rate[key].setdefault(rate or f"#{line}", v)
    for key in ("igst",):
        for v, p in _uniq(c.get(key, []))[:3]:
            variants.append(({"igst": v}, p))
        if len(by_rate.get(key, {})) > 1:
            variants.append(({"igst": round(sum(by_rate[key].values()), 2)}, 0.2))
    cg, sg = _uniq(c.get("cgst", []))[:3], _uniq(c.get("sgst", []))[:3]
    for v, p in cg:
        mates = [s for s in sg if abs(s[0] - v) <= 1] or [(v, p + 1)]
        variants.append(({"cgst": v, "sgst": mates[0][0]}, p))
    if len(by_rate.get("cgst", {})) > 1:
        tc = round(sum(by_rate["cgst"].values()), 2)
        variants.append(({"cgst": tc, "sgst": tc}, 0.2))
    for v, p in _uniq(c.get("utgst", []))[:2]:
        variants.append(({"cgst": cg[0][0] if cg else None, "utgst": v}, p))
    for v, p in _uniq(c.get("total_tax", []))[:4]:
        split = {"cgst": round(v / 2, 2), "sgst": round(v / 2, 2), "_from_total_tax": True}
        whole = {"igst": v, "_from_total_tax": True}
        variants.append((split, p + 1) if intra_hint else (whole, p + 1))
        variants.append((whole, p + 1.4) if intra_hint else (split, p + 1.4))
    if triplets:
        tt = round(sum(t for a, r, t in triplets), 2)
        # audit fix E7: offer both IGST and CGST+SGST readings of the tax summary and let the
        # balance decide - the place-of-supply hint alone mis-labelled CGST/SGST as IGST
        # (Caliber-2 / J00089).  The hinted reading keeps the better priority.
        p_i, p_c = (1.9, 1.5) if intra_hint else (1.5, 1.9)
        variants.append(({"igst": tt}, p_i))
        variants.append(({"cgst": tt, "sgst": tt}, p_c))
        variants.append(({"cgst": round(tt / 2, 2), "sgst": round(tt / 2, 2)}, p_c + 0.1))
        # single summary rows (the largest taxable rows) - junk triplets from item
        # lines no longer pollute the sum (C.pdf / Vel Agency)
        for a, r, t in sorted(set(triplets), key=lambda x: -x[0])[:3]:
            variants.append(({"igst": t}, p_i + 0.3))
            variants.append(({"cgst": t, "sgst": t}, p_c + 0.3))
            taxables.append((a, 2.0))
        taxables = _uniq(taxables)[:12]
    explicit_zero = any(v == 0 for k in ("igst", "cgst", "sgst") for v, _ in c.get(k, []))
    if not variants or explicit_zero:
        variants.append(({}, 3))
    else:
        # last resort so taxable/total are still reported when every printed tax amount
        # is implausible (validation then reports the missing/unbalanced tax)
        variants.append(({"_fallback": True}, 6))
    best = None
    for tv, ptv in taxables or [(None, 9)]:
        for tax, ptax in variants:
            taxsum = sum(v or 0 for k, v in tax.items() if not k.startswith("_")) + (cess or 0)
            if tv is not None and taxsum > 0 and (tv <= 0 or taxsum > tv * 0.405 + 1):
                # audit fix E3: 'taxable 0 + tax = total' (or a tax rate above the 40% GST
                # ceiling) is arithmetic coincidence, not a balanced invoice.  It auto-verified
                # C.pdf (500 instead of 5,987) and Sagar 1501 (234 instead of 1,534).
                continue
            trivial = taxsum == 0 and any(v > 0 for k in ("igst", "cgst", "sgst", "total_tax") for v, _ in c.get(k, []))
            if trivial:
                ptax += 5    # avoid the trivial 'no tax' balance when real tax amounts are printed
            for ro, pro in ro_opts:
                for oc, poc in oc_opts:
                    tot_opts = list(totals) or [(None, 9)]
                    if tv is not None and taxsum > 0:
                        computed = round(tv + taxsum + oc - disc + ro, 2)
                        if computed in printed and all(abs(computed - t) > 0.005 for t, _ in totals):
                            tot_opts.append((computed, 2.5))
                    for tot, ptot in tot_opts:
                        if tv is None or tot is None:
                            diff = None
                        else:
                            diff = abs(tv + taxsum + oc - disc + ro - tot)
                        # audit fix E3: 'taxable = total, tax 0' while tax amounts are printed is not a balance
                        ok = diff is not None and diff <= tol and not tax.get("_fallback") and not (trivial and tv == tot)
                        miss = round(diff / max(abs(tot or 1), 1), 4) if (diff is not None and not ok) else 0
                        if tax.get("_fallback") or (trivial and tv == tot):
                            miss = 1.0     # never prefer the 'no tax' fallback over a real near-miss
                        key = (0 if ok else 1, miss, ptv + ptax + ptot + pro + poc)
                        if best is None or key < best[0]:
                            best = (key, tv, tax, tot, ro, oc)
    if best is None:   # every combination was implausible: report nothing rather than guess
        best = ((1, 0, 99), None, {}, totals[0][0] if totals else None, 0.0, 0.0)
    key, tv, tax, tot, ro, oc = best
    ok = key[0] == 0
    out = {"taxable_value": tv, "igst": tax.get("igst"), "cgst": tax.get("cgst"), "sgst": tax.get("sgst"),
           "utgst": tax.get("utgst"), "cess": cess, "round_off": ro if ro else None,
           "other_charges": oc or None, "discount": disc or None, "invoice_total": tot}
    if not ok:
        # derive missing pieces only when unambiguous; mark them
        taxsum = sum(v or 0 for k, v in tax.items() if not k.startswith("_"))
        if tv is None and tot is not None and tax:
            out["taxable_value"] = round(tot - taxsum, 2)
            out["_derived"] = "taxable_value"
        out["round_off"] = c["round_off"][0][0] if c.get("round_off") else None
    if tax.get("_from_total_tax"):
        out["_derived"] = (out.get("_derived", "") + " tax split from 'Total Tax'").strip()
    out["_prio"] = key[2]
    return out, ok


# ------------------------------------------------------------------ names / addresses
def _clean_name(s: str) -> str:
    s = re.sub(r"(?i)^(m/s\.?|to\s*[:.]?|name\s*:|customer\s*:)\s*", "", s.strip())
    s = re.sub(r"\s{2,}.*$", "", s)        # drop columns to the right
    s = s.strip(" :,-|\\")
    return s


def name_after_label(text: str, role_rx: re.Pattern) -> tuple[str, str]:
    lines = _lines(text)
    for i, line in enumerate(lines):
        m = role_rx.search(line)
        if not m:
            continue
        rest = _clean_name(line[m.end():])
        cand_lines = ([rest] if len(rest) > 3 else []) + [_clean_name(l) for l in lines[i + 1:i + 6]]
        name, addr = "", []
        for c in cand_lines:
            if not c or c in (":", "-") or NOT_NAME.match(c) and not COMPANY_HINT.search(c):
                if name:
                    break
                continue
            if not name:
                if re.search(r"[A-Za-z]{3,}", c) and not GSTIN_RE.search(c.upper()):
                    name = c
                continue
            if re.search(r"(?i)gstin|state\s*name|state\s*:|phone|mobile|email|pan\b|contact", c) or GSTIN_RE.search(c.upper()):
                break
            addr.append(c)
            if len(addr) >= 4:
                break
        if name:
            return name, ", ".join(addr)
    return "", ""


def issuer_from_signature(text: str) -> str:
    """'for Mahendra Appliances' / 'For Caliber Inno Enterprise Pvt Ltd'."""
    best = ""
    for m in re.finditer(r"(?im)^\s*(?:for|For|FOR)\s+([A-Z][A-Za-z0-9&.,()' \-]{3,70}?)\s*(?:\(\d{4}-\d{4}\))?\s*$", text):
        name = m.group(1).strip()
        if re.search(r"(?i)^(the|any|all|delivery|further|more|payment|our|your|goods|interest|which|this|that)\b", name):
            continue
        best = name
    if not best:
        m = re.search(r"(?i)A/?c\s*Holder'?s\s*Name\s*:?\s*([A-Z][A-Za-z0-9&. ]{3,60})", text)
        if m:
            best = m.group(1).strip()
    return best


def header_company(text: str) -> str:
    for line in _lines(text)[:18]:
        c = _clean_name(line)
        if len(c) < 4 or NOT_NAME.match(c) or GSTIN_RE.search(c.upper()):
            continue
        if re.search(r"(?i)caliber", c):
            continue
        if COMPANY_HINT.search(c) or (c.isupper() and len(re.findall(r"[A-Z]{2,}", c)) >= 2):
            if not re.search(r"\d{3,}|road|street|floor|nagar|sector|plot|block|cross|main", c, re.I):
                return c
    return ""


def company_near(text: str, line_no: int, exclude_rx) -> str:
    """Company-looking line at or just above a GSTIN (footer layouts: Zoho, Odoo)."""
    lines = _lines(text)
    for j in range(line_no, max(-1, line_no - 6), -1):
        c = lines[j].strip()
        m = COMPANY_HINT.search(c)
        if not m or exclude_rx.search(c) or GSTIN_RE.search(c[:m.start()].upper()):
            continue
        name = re.split(r"\s+\d", c[:m.end()] + c[m.end():].split(",")[0], 1)[0].strip(" ,:-")
        name = re.sub(r"(?i)^(gstin|for)\s*:?\s*", "", name)
        if len(name) >= 4:
            return name
    return ""


_ADDR_HINT = re.compile(r"(?i)(road|street|\bst\.|floor|nagar|sector|plot|block|cross|main|lane|market|layout|area|"
                        r"\b\d{6}\b|\bpin|tower|building|complex|marg|bazar|bazaar|colony|village|industrial|estate|"
                        r"chennai|bangalore|bengaluru|delhi|mumbai|noida|hyderabad|surat|thane|pune)")
_ADDR_NOISE = re.compile(r"(?i)gstin|e-?mail|phone|mobile|irn|\back\b|jurisdiction|subject\s*to|terms|"
                         r"invoice|dated?\b|contact|www\.|@|pan\b|cin\b|udyam|msme|state\s*name|^\s*[:•]")


def address_near(text: str, line_no: int) -> str:
    """Address lines around the supplier GSTIN (above first, else below)."""
    lines = _lines(text)

    def ok(c):
        return c and _ADDR_HINT.search(c) and not _ADDR_NOISE.search(c) and not any(rx.search(c) for rx in ROLE_PATTERNS.values())

    above = []
    for j in range(max(0, line_no - 6), line_no):
        c = re.sub(r"\s{3,}.*$", "", lines[j].strip())
        if any(rx.search(c) for rx in ROLE_PATTERNS.values()):
            above = []
            continue
        if ok(c):
            above.append(c)
    if above:
        return ", ".join(above[-4:])
    below = []
    for j in range(line_no + 1, min(len(lines), line_no + 8)):
        c = re.sub(r"\s{3,}.*$", "", lines[j].strip())
        if any(rx.search(c) for rx in ROLE_PATTERNS.values()):
            break
        if ok(c):
            below.append(c)
    return ", ".join(below[:3])


# ------------------------------------------------------------------ misc fields
def find_place_of_supply(text: str) -> tuple[str, str]:
    m = re.search(r"(?is)Place\s*of\s*Supply\s*[:\-]?\s*\n?\s*:?\s*([A-Za-z][A-Za-z .()&\-]*?\(?\s*\d{0,2}\s*\)?)\s*(?:\n|$|\s{3,}|Reverse|E-?Way)", text)
    if not m:
        m = re.search(r"(?is)Place\s*of\s*\n\s*Supply\s*\n\s*([^\n]+)", text)
    if m:
        raw = m.group(1).strip(" :")
        return raw, state_code_from_name(raw)
    m = re.search(r"(?i)Place\s*of\s*Supply\s*:?\s*(\d{2})\s*-\s*([A-Za-z ]+)", text)
    if m:
        return f"{m.group(2).strip()} ({m.group(1)})", m.group(1)
    return "", ""


def find_irn(text: str) -> str:
    m = re.search(r"(?i)IRN(?:\s*Number)?\s*:?-?\s*([0-9a-f][0-9a-f\-\s]{40,90})", text)
    if not m:
        return ""
    hexs = re.sub(r"[^0-9a-f]", "", m.group(1).lower())
    return hexs[:64] if len(hexs) >= 64 else ""


def find_eway(text: str) -> str:
    m = re.search(r"(?i)(?:e-?\s*way\s*bill|waybill)\s*(?:no\.?|number)?\s*[:\-]*\s*\n?\s*((?:\d\s?){12})(?!\d)", text)
    if m:
        return re.sub(r"\s", "", m.group(1))
    m = re.search(r"(?is)e-?way\s*bill\s*no\.?.{0,120}?\b(\d{12})\b", text)
    return m.group(1) if m else ""


def find_po(text: str) -> str:
    v = find_labelled(text, r"P\.?\s*O\.?\s*(?:No|Number|#)\.?\s*:?|Buyer'?s\s*(?:Ref\.?/)?Order\s*No\.?\s*:?|Customer\s*PO\s*NO|"
                            r"Purchase\s*Order\s*No\.?\s*:?|P\.O\.#\s*:?|Buyer\s*PO\s*No\s*:?",
                      lambda s: _valid_docno(s) and not re.search(r"(?i)dated|dispatch|date", s), lookahead=2)
    return v or ""


def find_original_ref(text: str) -> str:
    v = find_labelled(text, r"Against\s*Invoice\s*No\.?\s*:?|Original\s*Invoice\s*(?:No|Number)\.?\s*:?|"
                            r"Ref(?:erence)?\s*Invoice\s*(?:No)?\.?\s*:?|Invoice\s*Reference\s*:?",
                      _valid_docno, lookahead=2)
    return v or ""


def find_reverse_charge(text: str) -> str:
    m = re.search(r"(?i)Reverse\s*Charge\s*[:\-]?\s*\n?\s*:?\s*(Yes|No|Y|N)\b", text)
    if not m:
        return ""
    return "Yes" if m.group(1).upper().startswith("Y") else "No"


def find_hsn(text: str) -> list[str]:
    lines = _lines(text)
    start = next((i for i, l in enumerate(lines) if re.search(r"(?i)\bHSN", l)), None)
    if start is None:
        return []
    stop = len(lines)
    for i in range(start + 1, len(lines)):
        if re.search(r"(?i)amount chargeable|total in words|bank details|terms\s*(&|and)\s*conditions|declaration", lines[i]):
            stop = i
            break
    pins = set(re.findall(r"(?:-|\b(?:pin(?:code)?|chennai|bangalore|delhi)\b)\s*(\d{6})\b", text, re.I))
    out = []
    for l in lines[start:stop]:
        if re.search(r"(?i)gstin|irn|phone|mobile|a/?c|account|ifsc|ack", l):
            continue
        for m in re.finditer(r"(?<![\d,.\-/])(\d{4}(?:\d{2}){0,2})(?![\d,.%/\-])", l):
            v = m.group(1)
            if v in pins or v.startswith("0") or v in out:
                continue
            if len(v) == 4 and 1990 < int(v) < 2100:
                continue
            if 1 <= int(v[:2]) <= 99:
                out.append(v)
    return out[:12]


# ------------------------------------------------------------------ main entry
def extract_fields(views: list[str], doc_type: str, master, settings) -> tuple[dict, dict, list]:
    """Return (fields, confidence per field, tax-rate breakdown rows)."""
    views = [v for v in views if v and v.strip()]
    if not views:
        return {}, {}, []
    known = master.all_gstins() if master else set()
    f: dict = {}
    conf: dict = {}

    def first_view(fn, *a):
        for idx, v in enumerate(views):
            r = fn(v, *a)
            if r:
                return r, (1.0 if idx == 0 else 0.85)
        return None, 0.0

    # --- parties ----------------------------------------------------------------
    occ_best, occ_view = [], 0
    for idx, v in enumerate(views):
        occ = find_gstins(v, known)
        if occ and (not occ_best or (len({o['role'] for o in occ}) > len({o['role'] for o in occ_best}))):
            occ_best, occ_view = occ, idx
    occ = occ_best
    all_g = list(OrderedDict.fromkeys(o["gstin"] for o in occ if o["valid"]))
    caliber_g = [g for g in all_g if master and master.is_caliber_gstin(g)]
    other_g = [g for g in all_g if g not in caliber_g]

    def role_of(g):
        return [o["role"] for o in occ if o["gstin"] == g]

    bill = next((o for o in occ if o["role"] == "BILL" and o["valid"]), None)
    ship = next((o for o in occ if o["role"] == "SHIP" and o["valid"]), None)
    full_text = "\n".join(views)
    issuer = issuer_from_signature(views[0]) or issuer_from_signature(full_text)
    cal_rx = re.compile(settings.caliber_name_regex, re.I)
    issuer_is_caliber = bool(issuer and cal_rx.search(issuer.replace(" ", " ")))
    header_cal = next((o for o in occ if o["gstin"] in caliber_g and o["role"] in ("HEADER", "DISPATCH", "SUPPLIER")), None)

    direction = DIRECTION_UNKNOWN
    supplier_g = customer_g = ""
    if issuer_is_caliber or (header_cal and bill and bill["gstin"] not in caliber_g and header_cal["line"] < bill["line"]):
        direction = OUTWARD
        supplier_g = header_cal["gstin"] if header_cal else (caliber_g[0] if caliber_g else "")
        customer_g = bill["gstin"] if bill and bill["gstin"] not in caliber_g else next(iter(other_g), "")
    elif bill and bill["gstin"] in caliber_g:
        direction = INWARD
        customer_g = bill["gstin"]
    elif bill and bill["gstin"] not in caliber_g:
        direction = CONSIGNEE_ONLY if (ship and ship["gstin"] in caliber_g) else THIRD_PARTY
        customer_g = bill["gstin"]
    elif caliber_g:
        direction = INWARD
        customer_g = (ship["gstin"] if ship and ship["gstin"] in caliber_g else caliber_g[0])
        conf["customer_gstin_note"] = 0.6
    elif cal_rx.search(full_text):
        direction = INWARD    # Caliber named as buyer but no GSTIN printed
    elif len(all_g) >= 2:
        direction = THIRD_PARTY
    # audit fix E14: a single (supplier) GSTIN with no buyer label and no Caliber name is not proof
    # that someone else is the buyer - typically a handwritten/poor scan; leave it 'Not determined'.
    if direction in (INWARD, CONSIGNEE_ONLY, THIRD_PARTY, DIRECTION_UNKNOWN) and not supplier_g:
        cands = [o for o in occ if o["valid"] and o["gstin"] not in caliber_g and o["gstin"] != customer_g
                 and o["role"] in ("HEADER", "SUPPLIER", "DISPATCH")]
        cands = cands or [o for o in occ if o["valid"] and o["gstin"] not in caliber_g and o["gstin"] != customer_g
                          and o["role"] not in ("BILL", "SHIP")]
        if cands:
            supplier_g = cands[0]["gstin"]
    f["direction"] = direction
    f["supplier_gstin"] = supplier_g
    f["customer_gstin"] = customer_g
    f["bill_to_gstin"] = bill["gstin"] if bill else ""
    f["ship_to_gstin"] = ship["gstin"] if ship else ""
    f["all_gstins"] = all_g
    f["issuer_name"] = issuer
    for key in ("supplier_gstin", "customer_gstin", "bill_to_gstin", "ship_to_gstin"):
        o = next((o for o in occ if o["gstin"] == f[key] and o["how"] == "corrected"), None)
        # only warn when no text view prints the GSTIN verbatim (embedded text confirms the OCR repair)
        if o and f[key] not in full_text.upper():
            f[f"{key}_corrected"] = o["raw"]
    conf["supplier_gstin"] = 0.95 if supplier_g else 0.0
    conf["customer_gstin"] = conf.pop("customer_gstin_note", 0.95) if customer_g else 0.0

    # --- names / states / addresses ----------------------------------------------
    pv = views[min(occ_view, len(views) - 1)]
    bill_name, bill_addr = name_after_label(pv, ROLE_PATTERNS["BILL"])
    ship_name, ship_addr = name_after_label(pv, ROLE_PATTERNS["SHIP"])
    f["bill_to_name"], f["ship_to_name"] = bill_name, ship_name
    if direction == OUTWARD:
        cal = master.lookup(supplier_g) if (master and supplier_g) else None
        f["supplier_legal_name"] = (cal or {}).get("legal_name") or issuer or "Caliber Inno Enterprise Private Limited"
        f["supplier_trade_name"] = (cal or {}).get("trade_name", "")
        f["supplier_address"] = (cal or {}).get("address", "")
        f["customer_legal_name"] = bill_name
        f["customer_address"] = bill_addr
    else:
        sup_name = issuer if (issuer and not cal_rx.search(issuer)) else header_company(pv) or header_company(full_text)
        sup_line = next((o["line"] for o in occ if o["gstin"] == supplier_g), None)
        if not sup_name and sup_line is not None:
            sup_name = company_near(pv, sup_line, cal_rx)
        # v1.1.1: drop OCR-merged neighbours ('NOORI STORE  (27AHHPK0276E1ZI)', '... LIMITED Invoice No.')
        sup_name = re.sub(r"\(?\b\d{2}[A-Z]{5}\d{4}[A-Z][0-9A-Z]Z[0-9A-Z]\b\)?", "", sup_name or "")
        sup_name = re.split(r"\s+(?=(?:Invoice|Tax\s*Invoice|Dated?|GSTIN)\b)", sup_name, flags=re.I)[0].strip(" ,:-")
        if re.search(r"(?i)authori[sz]ed|use\s*only|signatory", sup_name):
            sup_name = ""
        f["supplier_legal_name"] = sup_name
        f["supplier_trade_name"] = ""
        f["supplier_address"] = address_near(pv, sup_line) if sup_line is not None else ""
        cal = master.lookup(customer_g) if (master and customer_g) else None
        f["customer_legal_name"] = (cal or {}).get("legal_name") or bill_name
        f["customer_trade_name"] = (cal or {}).get("trade_name", "")
        f["customer_address"] = bill_addr or (cal or {}).get("address", "")
    conf["supplier_legal_name"] = 0.9 if f.get("supplier_legal_name") else 0.0
    for key, code_key in (("supplier_gstin", "supplier"), ("customer_gstin", "customer")):
        g = f.get(key) or ""
        f[f"{code_key}_state_code"] = g[:2] if g else ""
        f[f"{code_key}_state"] = state_name(g) if g else ""
    m = re.search(r"(?i)State\s*Name\s*:?\s*([A-Za-z ]+?)\s*,?\s*Code\s*:?\s*(\d{2})", pv)
    if m and f.get("supplier_gstin"):
        f["supplier_state_stated"] = f"{m.group(1).strip()} ({m.group(2)})"
    # printed state next to each GSTIN (catches a ship-to GSTIN from another state)
    conflicts = []
    plines = _lines(pv)
    for o in occ:
        window = " ".join(plines[o["line"]:o["line"] + 3])
        sm = re.search(r"(?i)State\s*(?:Name)?\s*:?\s*([A-Za-z ]{3,30}?)\s*,?\s*Code\s*:?\s*(\d{2})", window)
        if sm and sm.group(2) != o["gstin"][:2]:
            conflicts.append(f"{o['role'].title()} GSTIN {o['gstin']} (state code {o['gstin'][:2]}) is printed with "
                             f"state '{sm.group(1).strip()}, Code {sm.group(2)}'")
    f["gstin_state_conflicts"] = list(dict.fromkeys(conflicts))

    # --- numbers, dates -----------------------------------------------------------
    no, vidx = find_doc_number_views(views, doc_type)
    f["invoice_number"], conf["invoice_number"] = (no or ""), (0.0 if not no else 1.0 if vidx == 0 else 0.9)
    iso, raw, cdate = "", "", 0.0
    for idx, v in enumerate(views):
        iso, raw = find_date(v)
        if raw:
            cdate = 1.0 if idx == 0 else 0.85
            break
    f["invoice_date"], f["invoice_date_raw"], conf["invoice_date"] = iso, raw, (cdate if iso else 0.0)
    f["original_invoice_number"] = first_view(find_original_ref)[0] or ""
    pos_raw, pos_code = first_view(find_place_of_supply)[0] or ("", "")
    f["place_of_supply"] = f"{state_name(pos_code)} ({pos_code})" if pos_code else pos_raw
    f["place_of_supply_code"] = pos_code
    f["reverse_charge"] = first_view(find_reverse_charge)[0] or ""
    f["irn"] = first_view(find_irn)[0] or ""
    f["eway_bill"] = first_view(find_eway)[0] or ""
    f["po_number"] = first_view(find_po)[0] or ""
    f["hsn"] = (first_view(find_hsn)[0] or [])
    f["currency"] = "INR" if re.search(r"₹|\bINR\b|\bRs\.?|Rupees", full_text) else ""

    # --- amounts --------------------------------------------------------------------
    has_c = bool(re.search(r"\bCGST|\bSGST|Central\s*Tax", full_text, re.I))
    has_i = bool(re.search(r"\bIGST|Integrated\s*Tax", full_text, re.I))
    if has_c != has_i:
        intra = has_c
    else:
        sc = (f.get("supplier_gstin") or "")[:2]
        intra = bool(sc) and sc == (f.get("place_of_supply_code") or (f.get("customer_gstin") or "")[:2])
    best = None
    for idx, v in enumerate(views):
        cands = _amount_candidates(v)
        trip = tax_triplets(v)
        amounts, ok = reconcile(cands, trip, settings.amount_tolerance, intra)
        filled = sum(1 for k in ("taxable_value", "invoice_total") if amounts.get(k) is not None)
        # audit fix E7: compare balanced readings of all views by how well-labelled they are,
        # instead of taking the first view that balances (a 'GST Total' read as IGST balanced
        # just as well as the column-aligned CGST/SGST row of another view)
        score = (ok, filled, -round(amounts.get("_prio", 99), 3), -idx)
        if best is None or score > best[0]:
            best = (score, amounts, trip, idx)
    (ok, _, _, _), amounts, trip, aidx = best
    for k in ("taxable_value", "igst", "cgst", "sgst", "utgst", "cess", "other_charges", "discount", "round_off", "invoice_total"):
        val = amounts.get(k)
        f[k] = round(val, 2) if isinstance(val, (int, float)) else None
    f["amounts_balanced"] = bool(ok)
    f["amounts_derived"] = amounts.get("_derived", "")
    base = 1.0 if aidx == 0 else 0.85
    for k in ("taxable_value", "invoice_total"):
        conf[k] = (base if ok else 0.6) if f.get(k) is not None else 0.0
    conf["tax"] = (base if ok else 0.6) if any(f.get(k) is not None for k in ("igst", "cgst", "sgst", "utgst")) else 0.0
    if amounts.get("_derived"):
        conf["tax"] = min(conf["tax"], 0.7)
        conf["taxable_value"] = min(conf["taxable_value"], 0.7)

    # --- tax-rate breakdown -----------------------------------------------------------
    breakdown = []
    if trip:
        intra = bool(f.get("cgst") or f.get("sgst"))
        agg: dict = OrderedDict()
        for a, r, t in trip:
            agg.setdefault(r, {"taxable": 0.0, "tax": 0.0})
            agg[r]["taxable"] += a
            agg[r]["tax"] += t
        tot_tax = sum(v or 0 for v in (f.get("igst"), f.get("cgst"), f.get("sgst"), f.get("utgst")))
        sum_tax = sum(v["tax"] for v in agg.values())
        halves = intra and tot_tax and abs(sum_tax * 2 - tot_tax) <= 1.0
        doubled = intra and tot_tax and abs(sum_tax - tot_tax) <= 1.0 and len(trip) % 2 == 0
        for r, v in agg.items():
            taxable = v["taxable"] / 2 if doubled else v["taxable"]
            tax = v["tax"] / 2 if doubled else v["tax"]
            eff = r * 2 if (halves or doubled) else r
            row = {"rate": eff, "taxable_value": round(taxable, 2)}
            if intra:
                row.update({"igst": None, "cgst": round(tax, 2), "sgst": round(tax, 2)})
            else:
                row.update({"igst": round(tax, 2), "cgst": None, "sgst": None})
            breakdown.append(row)
        bt = sum(r["taxable_value"] for r in breakdown)
        status = "Complete" if f.get("taxable_value") and abs(bt - f["taxable_value"]) <= 1.0 else "Partial - verify"
        for r in breakdown:
            r["breakdown_status"] = status
        if status == "Complete":
            f["tax_rates"] = sorted({r["rate"] for r in breakdown})
    if not f.get("tax_rates"):
        rates = set()
        for m in re.finditer(r"(?<![\d.])(\d{1,2}(?:\.\d{1,2})?)\s*%", full_text):
            r = float(m.group(1))
            if r in (0.1, 0.25, 1, 1.5, 3, 5, 12, 18, 28, 40):
                rates.add(r)
        tv, taxsum = f.get("taxable_value"), sum(v or 0 for v in (f.get("igst"), f.get("cgst"), f.get("sgst"), f.get("utgst")))
        if tv and taxsum:
            eff = round(taxsum / tv * 100, 2)
            match = [r for r in rates if abs(r - eff) < 0.2]
            f["tax_rates"] = match or ([eff] if not rates else sorted(rates))
        else:
            f["tax_rates"] = sorted(rates)
    return f, conf, breakdown


def extraction_confidence(conf: dict, ocr_conf: Optional[float], balanced: bool) -> float:
    keys = ["invoice_number", "invoice_date", "supplier_gstin", "customer_gstin", "taxable_value", "tax", "invoice_total"]
    vals = [conf.get(k, 0.0) for k in keys]
    score = sum(vals) / len(vals)
    if ocr_conf is not None:
        score *= max(0.5, min(1.0, ocr_conf + 0.1))
    if not balanced:
        score *= 0.85
    return round(score, 3)

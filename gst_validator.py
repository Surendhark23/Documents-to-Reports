"""GSTIN utilities and document-level GST validation rules.

Nothing here changes an extracted amount: every rule only records a
validation message (ERROR / WARNING / INFO) on the document.  GST
conclusions that depend on facts not visible on the document are raised as
WARNINGs for human review instead of being decided automatically.
"""
from __future__ import annotations

import itertools
import re
from datetime import date, datetime
from typing import Iterable, Optional

from models import (Document, INVOICE_TYPES, NOTE_TYPES, INWARD, OUTWARD,
                    THIRD_PARTY, CONSIGNEE_ONLY, DIRECTION_UNKNOWN,
                    ST_VERIFIED, ST_REVIEW, ST_NON_TAX, ST_OUTWARD,
                    UNKNOWN, EXCEL_REGISTER, REVISED_TAX_INVOICE, TAX_INVOICE as TAX_INVOICE_TYPE)

# ------------------------------------------------------------------ states
STATE_CODES = {
    "01": "Jammu and Kashmir", "02": "Himachal Pradesh", "03": "Punjab",
    "04": "Chandigarh", "05": "Uttarakhand", "06": "Haryana", "07": "Delhi",
    "08": "Rajasthan", "09": "Uttar Pradesh", "10": "Bihar", "11": "Sikkim",
    "12": "Arunachal Pradesh", "13": "Nagaland", "14": "Manipur", "15": "Mizoram",
    "16": "Tripura", "17": "Meghalaya", "18": "Assam", "19": "West Bengal",
    "20": "Jharkhand", "21": "Odisha", "22": "Chhattisgarh", "23": "Madhya Pradesh",
    "24": "Gujarat", "25": "Daman and Diu",
    "26": "Dadra and Nagar Haveli and Daman and Diu", "27": "Maharashtra",
    "28": "Andhra Pradesh (Old)", "29": "Karnataka", "30": "Goa", "31": "Lakshadweep",
    "32": "Kerala", "33": "Tamil Nadu", "34": "Puducherry",
    "35": "Andaman and Nicobar Islands", "36": "Telangana", "37": "Andhra Pradesh",
    "38": "Ladakh", "97": "Other Territory", "99": "Centre Jurisdiction",
}
# Union territories without legislature -> UTGST instead of SGST
UTGST_CODES = {"04", "25", "26", "31", "35", "38"}

_STATE_ALIASES = {
    "tamilnadu": "33", "tamil nadu": "33", "tn": "33", "chennai": "33",
    "karnataka": "29", "kamataka": "29", "bangalore": "29", "bengaluru": "29", "banglore": "29",
    "maharashtra": "27", "mumbai": "27", "pune": "27", "thane": "27",
    "delhi": "07", "new delhi": "07", "uttar pradesh": "09", "up": "09", "noida": "09",
    "telangana": "36", "hyderabad": "36", "andhra pradesh": "37", "gujarat": "24",
    "gujrat": "24", "haryana": "06", "kerala": "32", "west bengal": "19",
    "rajasthan": "08", "punjab": "03", "goa": "30", "odisha": "21", "orissa": "21",
    "bihar": "10", "jharkhand": "20", "assam": "18", "madhya pradesh": "23",
    "chhattisgarh": "22", "puducherry": "34", "pondicherry": "34", "chandigarh": "04",
    "uttarakhand": "05", "himachal pradesh": "02", "jammu and kashmir": "01",
}
for _c, _n in STATE_CODES.items():
    _STATE_ALIASES.setdefault(_n.lower(), _c)


def state_name(code: str) -> str:
    return STATE_CODES.get((code or "")[:2], "")


def state_code_from_name(text: str) -> str:
    """Best-effort state code from free text such as 'Tamil Nadu ( 33 )'."""
    if not text:
        return ""
    m = re.search(r"\b(\d{2})\b", text)
    if m and m.group(1) in STATE_CODES:
        return m.group(1)
    t = re.sub(r"[^a-z ]", " ", text.lower())
    t = re.sub(r"\s+", " ", t).strip()
    # longest alias first so 'andhra pradesh' beats 'pradesh'
    for alias in sorted(_STATE_ALIASES, key=len, reverse=True):
        if len(alias) <= 2:
            if re.search(rf"\b{alias}\b", t) and t == alias:
                return _STATE_ALIASES[alias]
            continue
        if re.search(rf"\b{re.escape(alias)}\b", t):
            return _STATE_ALIASES[alias]
    return ""


# ------------------------------------------------------------------ GSTIN
_CHARS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
GSTIN_RE = re.compile(r"\b\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z][A-Z][0-9A-Z]\b")
# Looser pattern used to find OCR-damaged GSTIN candidates.
GSTIN_LOOSE_RE = re.compile(r"(?<![A-Z0-9])[0-9OQDZ][0-9OQDISZ][A-Z0-9]{5}[0-9A-Z]{4}[A-Z0-9][0-9A-Z][Z2][0-9A-Z](?![A-Z0-9])")


def gstin_checksum(g14: str) -> str:
    total = 0
    for i, ch in enumerate(g14[:14]):
        v = _CHARS.index(ch) * (2 if i % 2 else 1)
        total += v // 36 + v % 36
    return _CHARS[(36 - total % 36) % 36]


def gstin_format_ok(g: str) -> bool:
    return bool(g) and len(g) == 15 and bool(GSTIN_RE.fullmatch(g)) and g[:2] in STATE_CODES


def gstin_valid(g: str) -> bool:
    try:
        return gstin_format_ok(g) and gstin_checksum(g) == g[14]
    except ValueError:
        return False


def gstin_pan(g: str) -> str:
    return g[2:12] if g and len(g) >= 12 else ""


_TO_DIGIT = {"O": "0", "D": "0", "Q": "09", "I": "1", "L": "1", "Z": "2", "S": "5",
             "B": "8", "G": "6", "T": "7", "A": "4"}
_TO_ALPHA = {"0": "O", "1": "I", "2": "Z", "5": "S", "8": "B", "6": "G", "4": "A", "7": "T"}
_DIGIT_POS = {0, 1, 7, 8, 9, 10}
_ALPHA_POS = {2, 3, 4, 5, 6, 11}


def correct_gstin(candidate: str, known: Iterable[str] = ()) -> tuple[str, str]:
    """Try to repair an OCR-damaged GSTIN.

    Returns (gstin, how) where how is '' (already valid), 'corrected' or
    'invalid'.  A correction is only accepted when exactly one repaired value
    passes format + checksum, or when it matches a known GSTIN.
    """
    c = re.sub(r"[^A-Z0-9]", "", (candidate or "").upper())
    if len(c) != 15:
        return candidate, "invalid"
    if gstin_valid(c):
        return c, ""
    options = []
    for i, ch in enumerate(c):
        opts = {ch}
        if i in _DIGIT_POS and not ch.isdigit():
            opts |= set(_TO_DIGIT.get(ch, ""))
        if i in _ALPHA_POS and ch.isdigit():
            opts |= set(_TO_ALPHA.get(ch, ""))
        if i == 13 and ch != "Z":
            opts.add("Z")
        options.append(sorted(opts))
    combos = 1
    for o in options:
        combos *= len(o)
    valid = set()
    if combos <= 4096:
        for parts in itertools.product(*options):
            g = "".join(parts)
            if gstin_valid(g):
                valid.add(g)
    known = {k for k in known if k}
    for k in known:
        if sum(a != b for a, b in zip(k, c)) <= 2:
            return k, "corrected"
    if len(valid) == 1:
        return valid.pop(), "corrected"
    return c, "invalid"


# ------------------------------------------------------------------ helpers
def _f(v) -> Optional[float]:
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _close(a: float, b: float, settings) -> bool:
    tol = max(settings.amount_tolerance, abs(b) * settings.pct_tolerance)
    return abs(a - b) <= tol


# ------------------------------------------------------------------ rules
def validate_document(doc: Document, master, settings, today: Optional[date] = None) -> None:
    """Populate ``doc.validations`` for invoice-like documents."""
    f = doc.fields
    today = today or date.today()
    is_money_doc = doc.doc_type in INVOICE_TYPES | NOTE_TYPES or (
        doc.doc_type == UNKNOWN and (f.get("invoice_total") is not None or f.get("taxable_value") is not None)
        and bool(f.get("supplier_gstin") or f.get("customer_gstin")))

    # --- GSTIN format / checksum / state code ------------------------------
    for key, label in (("supplier_gstin", "Supplier"), ("customer_gstin", "Customer"),
                       ("bill_to_gstin", "Bill To"), ("ship_to_gstin", "Ship To")):
        g = f.get(key) or ""
        if not g:
            continue
        if not gstin_format_ok(g):
            doc.add_validation("GSTIN_FORMAT", "ERROR", f"{label} GSTIN {g} has an invalid format")
        elif not gstin_valid(g):
            doc.add_validation("GSTIN_CHECKSUM", "ERROR", f"{label} GSTIN {g} fails the checksum")
        if f.get(f"{key}_corrected"):
            doc.add_validation("GSTIN_OCR_CORRECTED", "WARNING",
                               f"{label} GSTIN auto-corrected from OCR text '{f[key + '_corrected']}' -> {g}; verify against the document")

    for key, skey, label in (("supplier_gstin", "supplier_state_stated", "Supplier"),
                             ("customer_gstin", "customer_state_stated", "Customer")):
        g, stated = f.get(key) or "", f.get(skey) or ""
        if g and stated:
            sc = state_code_from_name(stated)
            if sc and sc != g[:2]:
                doc.add_validation("STATE_MISMATCH", "WARNING",
                                   f"{label} GSTIN state code {g[:2]} ({state_name(g)}) differs from state printed on document '{stated}'")

    for msg in f.get("gstin_state_conflicts") or []:
        doc.add_validation("GSTIN_STATE_CONFLICT", "WARNING", msg + " - confirm the correct registration with the supplier")

    if not is_money_doc:
        return

    # --- Caliber registration ---------------------------------------------
    if doc.direction == INWARD:
        cg = f.get("customer_gstin") or ""
        if not cg:
            doc.add_validation("CUSTOMER_GSTIN_MISSING", "ERROR", "Caliber buyer GSTIN not found on the document")
        else:
            entry = master.lookup(cg) if master else None
            if entry is None:
                doc.add_validation("CUSTOMER_NOT_IN_MASTER", "ERROR", f"Customer GSTIN {cg} is not in the Caliber GST master")
            else:
                if not entry.get("active", True):
                    doc.add_validation("CUSTOMER_INACTIVE", "ERROR", f"Caliber registration {cg} is marked inactive")
                if entry.get("verification_status") != "Verified":
                    doc.add_validation("MASTER_UNVERIFIED", "WARNING",
                                       f"Caliber registration {cg} has not been verified in the master yet")
        st, sh = f.get("bill_to_gstin"), f.get("ship_to_gstin")
        if st and sh and st != sh and master and master.lookup(st) and master.lookup(sh):
            doc.add_validation("BILL_SHIP_DIFFERENT_CALIBER", "WARNING",
                               f"Bill-to {st} and ship-to {sh} are different Caliber registrations; register uses bill-to {st}")
        sup = f.get("supplier_gstin") or ""
        if not sup:
            doc.add_validation("SUPPLIER_GSTIN_MISSING", "ERROR", "Supplier GSTIN not found")
        elif master and master.lookup(sup):
            doc.add_validation("SUPPLIER_IS_CALIBER", "ERROR", "Supplier GSTIN is a Caliber registration")
    elif doc.direction in (THIRD_PARTY, CONSIGNEE_ONLY):
        doc.add_validation("NOT_BILLED_TO_CALIBER", "ERROR",
                           "Buyer (bill-to) is not a Caliber registration" +
                           (" - Caliber appears only as consignee/ship-to" if doc.direction == CONSIGNEE_ONLY else ""))
    elif doc.direction == DIRECTION_UNKNOWN:
        doc.add_validation("DIRECTION_UNKNOWN", "ERROR", "Could not determine supplier/buyer roles")

    # --- number and date -----------------------------------------------------
    if not f.get("invoice_number"):
        doc.add_validation("INVOICE_NUMBER_MISSING", "ERROR", "Invoice/note number not found")
    elif not re.search(r"\d", str(f["invoice_number"])) or len(str(f["invoice_number"])) > 30:
        doc.add_validation("INVOICE_NUMBER_SUSPECT", "ERROR", f"Invoice number '{f['invoice_number']}' does not look like a document number")
    d = f.get("invoice_date")
    if not d:
        if f.get("invoice_date_raw"):
            doc.add_validation("INVOICE_DATE_INVALID", "ERROR",
                               f"Invoice date '{f['invoice_date_raw']}' is not a valid calendar date")
        else:
            doc.add_validation("INVOICE_DATE_MISSING", "ERROR", "Invoice date not found")
    else:
        dd = datetime.strptime(d, "%Y-%m-%d").date()
        if dd > today:
            doc.add_validation("DATE_IN_FUTURE", "ERROR", f"Invoice date {dd:%d-%m-%Y} is in the future")
        if dd.year < today.year - 2:
            doc.add_validation("DATE_OLD", "WARNING", f"Invoice date {dd:%d-%m-%Y} is unusually old")

    # --- amounts -----------------------------------------------------------
    tv = _f(f.get("taxable_value"))
    igst, cgst, sgst, utgst, cess = (_f(f.get(k)) for k in ("igst", "cgst", "sgst", "utgst", "cess"))
    total = _f(f.get("invoice_total"))
    if tv is None:
        doc.add_validation("TAXABLE_MISSING", "ERROR", "Taxable value not found")
    if total is None:
        doc.add_validation("TOTAL_MISSING", "ERROR", "Invoice total not found")
    if all(x is None for x in (igst, cgst, sgst, utgst)):
        doc.add_validation("TAX_MISSING", "ERROR", "No IGST/CGST/SGST/UTGST amount found")
    for k in ("taxable_value", "igst", "cgst", "sgst", "invoice_total"):
        v = _f(f.get(k))
        if v is not None and v < 0:
            doc.add_validation("NEGATIVE_VALUE", "WARNING", f"{k} is negative ({v}); confirm it is a credit note")

    if tv is not None and total is not None:
        calc = tv + sum(x or 0 for x in (igst, cgst, sgst, utgst, cess)) \
            + (_f(f.get("other_charges")) or 0) - (_f(f.get("discount")) or 0) \
            + (_f(f.get("round_off")) or 0)
        if not _close(calc, total, settings):
            doc.add_validation("TOTAL_MISMATCH", "ERROR",
                               f"Taxable + taxes + charges + round-off = {calc:,.2f} but invoice total = {total:,.2f} (difference {total - calc:,.2f})")
    if (cgst or sgst) and igst:
        doc.add_validation("IGST_AND_CGST", "ERROR", "Both IGST and CGST/SGST present on one invoice")
    if cgst is not None and sgst is not None and not _close(cgst, sgst, settings):
        doc.add_validation("CGST_SGST_UNEQUAL", "ERROR", f"CGST {cgst:,.2f} and SGST {sgst:,.2f} differ")
    if (cgst is None) != (sgst is None) and not utgst:
        doc.add_validation("CGST_SGST_PAIR", "WARNING", "Only one of CGST/SGST found")

    # --- tax rate plausibility ------------------------------------------------
    taxes = sum(x or 0 for x in (igst, cgst, sgst, utgst))
    # audit fix E3: a balanced invoice with taxable 0 (or tax above the 40% GST ceiling) was
    # auto-verified; this is never a valid tax invoice, so stop it from reaching Final
    if taxes > 0 and tv is not None and (tv <= 0 or taxes > tv * 0.405 + 1):
        doc.add_validation("TAX_RATE_IMPLAUSIBLE", "ERROR",
                           f"Tax {taxes:,.2f} on taxable value {tv:,.2f} is not a possible GST rate - amounts mis-read")
    if re.search(r"(?i)\bbill\s*of\s*supply\b", doc.text or "") and doc.doc_type == TAX_INVOICE_TYPE:
        doc.add_validation("BILL_OF_SUPPLY", "WARNING",
                           "Document is headed 'Bill of Supply' (exempt / composition supplier) - no GST input "
                           "credit can be claimed; confirm before booking")
    if tv and taxes:
        eff = taxes / tv * 100
        rates = f.get("tax_rates") or []
        if not rates and not any(abs(eff - r) < 0.15 for r in (0.1, 0.25, 1, 1.5, 3, 5, 6, 12, 18, 28, 40)):
            doc.add_validation("MIXED_OR_ODD_RATE", "INFO",
                               f"Effective tax rate {eff:.2f}% - mixed rates or partly exempt items; see Tax_Rate_Breakdown")
        if len(rates) > 1:
            doc.add_validation("MULTIPLE_RATES", "INFO", f"Multiple GST rates on invoice: {', '.join(str(r) + '%' for r in rates)}")

    # --- place of supply vs tax type --------------------------------------------
    sup = f.get("supplier_gstin") or ""
    pos_code = f.get("place_of_supply_code") or ""
    if not pos_code and doc.direction == INWARD and f.get("customer_gstin"):
        pos_code = f["customer_gstin"][:2]
        doc.add_validation("POS_ASSUMED", "INFO", "Place of supply not printed; buyer state used for the tax-type check")
    if sup and pos_code:
        intra = sup[:2] == pos_code
        if intra and igst:
            doc.add_validation("POS_IGST_INTRASTATE", "WARNING",
                               f"Supplier state {sup[:2]} equals place of supply {pos_code} but IGST charged - review GST treatment (e.g. SEZ / special supply)")
        if not intra and (cgst or sgst):
            doc.add_validation("POS_CGST_INTERSTATE", "WARNING",
                               f"Supplier state {sup[:2]} differs from place of supply {pos_code} but CGST/SGST charged - review (place-of-supply rules may still allow this)")
        if intra and sgst and pos_code in UTGST_CODES:
            doc.add_validation("UTGST_EXPECTED", "INFO", "Place of supply is a Union Territory; tax may be UTGST rather than SGST")

    if doc.doc_type == REVISED_TAX_INVOICE and not f.get("original_invoice_number"):
        doc.add_validation("REVISED_NO_ORIGINAL", "WARNING", "Revised invoice - original invoice not identified on the document")
    if doc.doc_type in NOTE_TYPES and not f.get("original_invoice_number"):
        doc.add_validation("NOTE_NO_ORIGINAL", "WARNING", "Credit/debit note does not reference an original invoice")
    if f.get("reverse_charge") == "Yes":
        doc.add_validation("REVERSE_CHARGE", "WARNING", "Reverse charge applies - tax payable by recipient")


def decide_status(doc: Document, settings) -> str:
    """Translate classification + validations into a document status."""
    if doc.doc_type == EXCEL_REGISTER:
        return ST_NON_TAX
    if doc.doc_type not in INVOICE_TYPES | NOTE_TYPES:
        if doc.doc_type in (UNKNOWN,):
            return ST_REVIEW
        return ST_NON_TAX
    if doc.direction == OUTWARD and not settings.include_outward_in_final:
        return ST_OUTWARD
    errors = [v for v in doc.validations if v["level"] == "ERROR"]
    warnings = [v for v in doc.validations if v["level"] == "WARNING"]
    if errors or warnings:
        return ST_REVIEW
    if doc.extraction_confidence < settings.min_confidence_for_final:
        return ST_REVIEW
    if doc.type_confidence < 0.75:
        return ST_REVIEW
    return ST_VERIFIED

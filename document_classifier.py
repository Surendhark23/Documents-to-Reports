"""Document type identification and PDF segmentation.

Classification is content-first: the printed *title* of the document
(first ~25 lines, matched with spacing removed so 'TA X I N V O I C E'
still counts) carries most weight, then structural evidence (GSTINs, IRN,
invoice-number label, tax amounts), then body keywords.  The file name is
only used as supporting evidence and is reported in the reason.

Safety rules from the specification:
* any proforma / quotation wording in the header beats a 'Tax Invoice'
  heading -> never promoted to a tax invoice;
* conflicting strong signals (e.g. 'proforma' and an IRN) -> manual review;
* purchase / sales orders, challans, receipts and statements are never
  tax invoices, even when they show GST computations.
"""
from __future__ import annotations

import re

from models import (TAX_INVOICE, REVISED_TAX_INVOICE, CREDIT_NOTE, DEBIT_NOTE, PROFORMA,
                    QUOTATION, PURCHASE_ORDER, SALES_ORDER, DELIVERY_CHALLAN, EWAY_BILL,
                    PAYMENT, STATEMENT, EXCEL_REGISTER, OTHER_SUPPORTING, UNRELATED, UNKNOWN,
                    PageText)
from field_extractor import find_doc_number, find_irn
from gst_validator import GSTIN_RE

# compact title keyword -> type (checked longest first)
TITLE_KEYS = [
    ("REVISEDTAXINVOICE", REVISED_TAX_INVOICE), ("REVISEDINVOICE", REVISED_TAX_INVOICE),
    ("PROFORMAINVOICE", PROFORMA), ("PERFORMAINVOICE", PROFORMA), ("PROFOMAINVOICE", PROFORMA),
    ("PROFORMA", PROFORMA), ("PERFORMA", PROFORMA),
    ("CREDITNOTE", CREDIT_NOTE), ("DEBITNOTE", DEBIT_NOTE),
    ("PURCHASEORDER", PURCHASE_ORDER), ("WORKORDER", PURCHASE_ORDER), ("SALESORDER", SALES_ORDER),
    ("ORDERCONFIRMATION", SALES_ORDER), ("DELIVERYCHALLAN", DELIVERY_CHALLAN),
    ("SALESQUOTATION", QUOTATION), ("QUOTATION", QUOTATION), ("ESTIMATE", QUOTATION), ("QUOTE", QUOTATION),
    ("EWAYBILL", EWAY_BILL), ("PAYMENTRECEIPT", PAYMENT), ("PAYMENTADVICE", PAYMENT),
    ("RECEIPTVOUCHER", PAYMENT), ("CASHRECEIPT", PAYMENT), ("SHIPMENTCASHRECEIPT", PAYMENT),
    ("STATEMENTOFACCOUNT", STATEMENT), ("ACCOUNTSTATEMENT", STATEMENT),
    ("TAXINVOICE", TAX_INVOICE), ("BILLOFSUPPLY", TAX_INVOICE), ("TAXBILL", TAX_INVOICE),
    ("INVOICE", TAX_INVOICE),
]

RX = {
    "proforma": re.compile(r"\bpro\s*-?\s*forma\b|\bperforma\b|\bprofoma\b", re.I),
    "quotation": re.compile(r"\bquotation\b|\bquote\b(?!d)|\bestimate\b(?!d)|quotation\s*/\s*p\.?i", re.I),
    "po": re.compile(r"\bpurchase\s*order\b", re.I),
    "so": re.compile(r"\bsales\s*order\b|\border\s*confirmation\b", re.I),
    "challan": re.compile(r"\bdelivery\s*challan\b", re.I),
    "credit": re.compile(r"\bcredit\s+note\b", re.I),
    "debit": re.compile(r"\bdebit\s+note\b", re.I),
    "revised": re.compile(r"\brevised\s+(tax\s+)?invoice\b", re.I),
    "payment": re.compile(r"payment\s*(advice|receipt|voucher)|receipt\s*voucher|\bpaid\s+to\b|transaction\s*(id|successful)|"
                          r"scan\s*(here\s*)?to\s*pay|\bUPI\b|money\s*receipt|cash\s*receipt|payment\s*made", re.I),
    "statement": re.compile(r"statement\s*of\s*account|account\s*statement|opening\s*balance|closing\s*balance|"
                            r"transaction\s*date.{0,80}balance", re.I | re.S),
    "transport": re.compile(r"\bL\.?R\.?\s*No|lorry\s*receipt|consignment\s*note|\bconsignor|way\s*bill\s*number|"
                            r"air\s*way\s*bill|\bAWB\b|\bdocket\b|freight\s*type|booking\s*date|actual\s*weight|"
                            r"charged\s*weight|shipment\s*details|transporter\s*id|declared\s*value", re.I),
    "vendor_form": re.compile(r"vendor\s*(registration|creation|code\s*form)|cancelled\s*cheque|msme\s*declaration|"
                              r"address\s*of\s*supplier|statewise\s*gst", re.I),
    "unrelated": re.compile(r"non-?\s*disclosure|confidentiality\s*agreement|this\s*agreement|hereinafter|"
                            r"whereas|in\s*witness\s*whereof|mutual\s*nda", re.I),
    "eway": re.compile(r"e-?way\s*bill\s*system|ewb\s*no|generated\s*by|valid\s*upto", re.I),
    # audit fix E8: small generators (Vyapar etc.) label the tax column just 'GST' with '(18.0%)'
    "tax_amount": re.compile(r"\b(IGST|CGST|SGST|UTGST|Integrated\s*Tax|Central\s*Tax)\b", re.I),
    "tax_gst_col": re.compile(r"\bGST\b[^\n]{0,40}\n?[^\n]{0,60}\(\s*(?:0\.1|0\.25|1|1\.5|3|5|12|18|28|40)(?:\.0+)?\s*%\s*\)", re.I),
    "total": re.compile(r"grand\s*total|total\s*amount|invoice\s*total|total\s*\(inr\)|amount\s*chargeable|"
                        r"total\s*amount\s*after\s*tax|\btotal\b", re.I),
    "inv_label": re.compile(r"invoice\s*(no|number|#)|bill\s*(no|number)|inv\.?\s*no|"
                           r"invoice\s*details\s*:?[^\n]*\n[^\n]*\bNo\s*:", re.I),
}


def _compact(s: str) -> str:
    return re.sub(r"[^A-Z]", "", s.upper())


def title_types(text: str, max_lines: int = 25) -> list[tuple[int, str, str]]:
    """[(line_index, type, line)] for title-like lines near the top."""
    out = []
    lines = [l for l in text.splitlines() if l.strip()][:max_lines]
    for i, line in enumerate(lines):
        comp = _compact(line)
        if not comp or len(comp) > 60:
            # long lines: check for title words surrounded by big gaps (sorted layout)
            parts = re.split(r"\s{4,}|\|", line)
            comps = [_compact(p) for p in parts]
        else:
            comps = [comp]
        for c in comps:
            for key, t in TITLE_KEYS:
                if key in c and len(c) <= len(key) + 22:
                    # 'INVOICE' alone must be a heading, not a field label like 'Invoice No'
                    if key == "INVOICE" and re.search(r"INVOICE(NO|NUMBER|DATE|DETAILS|TO|PERIOD|VALUE|COPY|CONTACT|ADDRESS|ING)", c):
                        continue
                    if key == "PURCHASEORDER" and re.search(r"PURCHASEORDER(NO|NUMBER|DATE)", c) and t:
                        continue
                    if re.search(key + r"(NO|NUMBER|DATE|DETAILS)", c) and key not in ("INVOICE",):
                        continue
                    out.append((i, t, line.strip()))
                    break
            else:
                continue
            break
    return out


def classify(views: list[str], filename: str = "", is_excel_register: bool = False) -> tuple[str, float, str]:
    if is_excel_register:
        return EXCEL_REGISTER, 0.98, "Spreadsheet with invoice-register columns (e.g. GSTR-2B download)"
    views = [v for v in views if v and v.strip()]
    if not views:
        return UNKNOWN, 0.0, "No readable text (OCR failed or empty document)"
    text = "\n".join(views)
    header = "\n".join("\n".join(v.splitlines()[:30]) for v in views)
    fname = filename.lower()

    titles = []
    for v in views:
        titles.extend(title_types(v))
    title_set = [t for _, t, _ in titles]
    first_title = titles[0][1] if titles else None
    gstins = set(GSTIN_RE.findall(text.upper()))
    irn = find_irn(text)
    signals = {
        "gstin": bool(gstins), "irn": bool(irn), "inv_label": bool(RX["inv_label"].search(text)),
        "tax": bool(RX["tax_amount"].search(text) or (gstins and RX["tax_gst_col"].search(text))),
        "total": bool(RX["total"].search(text)),
    }
    nsig = sum(signals.values())
    reasons = []
    fname_hint = ""
    for kw, t in (("proforma", PROFORMA), ("profoma", PROFORMA), ("_pi", PROFORMA), ("pi-", PROFORMA),
                  ("quot", QUOTATION), ("purchase_order", PURCHASE_ORDER), ("po ", PURCHASE_ORDER),
                  ("challan", DELIVERY_CHALLAN), ("credit", CREDIT_NOTE), ("debit", DEBIT_NOTE),
                  ("revised", REVISED_TAX_INVOICE), ("invoice", TAX_INVOICE), ("inv", TAX_INVOICE)):
        if kw in fname:
            fname_hint = t
            break

    def done(t, c, why):
        r = why
        if fname_hint:
            agree = "agrees" if fname_hint == t or (fname_hint == REVISED_TAX_INVOICE and t == TAX_INVOICE) else "differs"
            r += f"; file name suggests '{fname_hint}' ({agree})"
        return t, round(c, 2), r

    # ---- notes ------------------------------------------------------------
    if first_title in (CREDIT_NOTE, DEBIT_NOTE) or (CREDIT_NOTE in title_set or DEBIT_NOTE in title_set):
        t = CREDIT_NOTE if (CREDIT_NOTE in title_set and (first_title != DEBIT_NOTE)) else DEBIT_NOTE
        return done(t, 0.95, f"Heading '{next(l for _, x, l in titles if x == t)}'")

    # ---- e-way bill print-out (often stapled behind the invoice) ----------------
    if first_title == EWAY_BILL and RX["eway"].search(text):
        return done(EWAY_BILL, 0.9, "E-way bill print-out (generated by / valid upto)")

    # ---- proforma / quotation beat 'tax invoice' --------------------------------
    pro_head = RX["proforma"].search(header)
    if PROFORMA in title_set or pro_head:
        if irn:
            return done(UNKNOWN, 0.4, "Conflict: proforma wording but an IRN (e-invoice) is printed - manual review")
        why = "Proforma wording in header"
        if TAX_INVOICE in title_set:
            why += " (overrides a 'Tax Invoice' heading/label printed on the same document)"
        return done(PROFORMA, 0.92 if PROFORMA in title_set else 0.85, why)
    # an explicit order heading wins over quotation words in the terms ("vendor's quotation")
    for t in (PURCHASE_ORDER, SALES_ORDER):
        if t in title_set and QUOTATION not in title_set and TAX_INVOICE not in title_set[:1]:
            return done(t, 0.93, f"Heading '{next(l for _, x, l in titles if x == t)}'")
    quo_head = RX["quotation"].search(header)
    if QUOTATION in title_set or (quo_head and first_title != TAX_INVOICE):
        if irn:
            return done(UNKNOWN, 0.4, "Conflict: quotation wording but an IRN is printed - manual review")
        return done(QUOTATION, 0.9 if QUOTATION in title_set else 0.75, "Quotation/estimate wording in header")
    if quo_head and first_title == TAX_INVOICE and not irn:
        reasons.append("header also mentions quotation/estimate")

    # ---- orders / challans -----------------------------------------------------------
    for t in (PURCHASE_ORDER, SALES_ORDER, DELIVERY_CHALLAN):
        if t in title_set and (first_title == t or TAX_INVOICE not in title_set):
            return done(t, 0.93, f"Heading '{next(l for _, x, l in titles if x == t)}'")
    if RX["po"].search(header) and TAX_INVOICE not in title_set and not irn:
        return done(PURCHASE_ORDER, 0.8, "'Purchase Order' in header")

    # ---- invoices ---------------------------------------------------------------------
    if REVISED_TAX_INVOICE in title_set or RX["revised"].search(header):
        return done(REVISED_TAX_INVOICE, 0.9, "Heading says revised invoice")
    if TAX_INVOICE in title_set:
        transportish = len(RX["transport"].findall(text)) >= 3
        if transportish and not irn and "TAXINVOICE" not in _compact(header)[:400]:
            return done(OTHER_SUPPORTING, 0.7, "Transport / courier document (LR, consignor, freight) with an invoice-copy heading")
        line = next(l for _, x, l in titles if x == TAX_INVOICE)
        if nsig >= 3:
            c = 0.95 if (signals["gstin"] and signals["tax"]) else 0.85
            if "revised" in fname:
                return done(REVISED_TAX_INVOICE, 0.7, f"Heading '{line}' and file name says 'revised' - confirm revision and link original")
            if reasons:
                c -= 0.15
            return done(TAX_INVOICE, c, f"Heading '{line}' with {', '.join(k for k, v in signals.items() if v)}" +
                        (f"; {'; '.join(reasons)}" if reasons else ""))
        return done(UNKNOWN, 0.5, f"Heading '{line}' but weak invoice structure ({nsig}/5 signals) - manual review")

    # ---- other categories (body evidence) ------------------------------------------------
    if RX["statement"].search(text):
        return done(STATEMENT, 0.85, "Bank/account statement columns (transaction date, balance)")
    if RX["payment"].search(text) and not signals["inv_label"]:
        return done(PAYMENT, 0.8, "Payment wording (paid to / UPI / receipt)")
    if len(RX["transport"].findall(text)) >= 2:
        return done(OTHER_SUPPORTING, 0.75, "Transport / courier document (LR, consignment, waybill)")
    if RX["vendor_form"].search(text):
        return done(OTHER_SUPPORTING, 0.7, "Vendor registration / KYC form")
    if EWAY_BILL in title_set and RX["eway"].search(text):
        return done(EWAY_BILL, 0.85, "E-way bill layout")
    if RX["unrelated"].search(text) and not signals["tax"]:
        return done(UNRELATED, 0.85, "Agreement / NDA text, no invoice content")
    if RX["proforma"].search(text):
        return done(PROFORMA, 0.7, "Proforma wording in body")
    if RX["quotation"].search(text) and not irn:
        return done(QUOTATION, 0.65, "Quotation wording in body (no heading found)")
    if nsig >= 4 and signals["inv_label"]:
        return done(UNKNOWN, 0.55, "Looks like an invoice but no heading identifies the document type - manual review")
    if len(text.strip()) < 60:
        return done(UNKNOWN, 0.2, "Very little readable text (poor scan/photo) - manual review")
    return done(UNKNOWN, 0.3, "No document-type evidence found - manual review")


# --------------------------------------------------------------- segmentation
_STANDALONE_PAGE_TYPES = {EWAY_BILL, DELIVERY_CHALLAN, CREDIT_NOTE, DEBIT_NOTE, PAYMENT}


def _page_type_hint(text: str) -> str:
    t = title_types(text, 12)
    return t[0][1] if t else ""


def segment_pages(pages: list[PageText]) -> list[dict]:
    """Split a multi-page file into logical documents.

    Returns [{'pages': [n...], 'text_pages': [PageText...], 'copies': k, 'doc_no': str}].
    A page carrying a *different* document number starts a new document; a
    page with the same number and a DUPLICATE/TRIPLICATE copy label is a
    printed copy of the same document (kept in the page range but its text is
    not added again so amounts are not double counted)."""
    segs: list[dict] = []
    for p in pages:
        hint = _page_type_hint(p.text)
        no = None
        for v in (p.views or [p.text])[:2]:
            no = find_doc_number(v, hint)
            if no:
                break
        if not segs:
            segs.append({"pages": [p.page_no], "text_pages": [p], "copies": 1, "doc_no": no or "", "hint": hint})
            continue
        cur = segs[-1]
        if no and cur["doc_no"] and no == cur["doc_no"]:
            cur["pages"].append(p.page_no)
            if p.copy_label and ("DUPLICATE" in p.copy_label or "TRIPLICATE" in p.copy_label):
                cur["copies"] += 1
            else:
                cur["text_pages"].append(p)
        elif no and (no != cur["doc_no"]) and (cur["doc_no"] or hint):
            segs.append({"pages": [p.page_no], "text_pages": [p], "copies": 1, "doc_no": no, "hint": hint})
        elif hint in _STANDALONE_PAGE_TYPES and hint != cur.get("hint"):
            # e.g. an e-way bill print-out stapled behind the invoice
            segs.append({"pages": [p.page_no], "text_pages": [p], "copies": 1, "doc_no": no or "", "hint": hint})
        else:
            cur["pages"].append(p.page_no)
            cur["text_pages"].append(p)
            if no and not cur["doc_no"]:
                cur["doc_no"] = no
    return segs

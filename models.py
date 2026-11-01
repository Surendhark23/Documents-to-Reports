"""Normalised data structures shared by every pipeline stage.

PDFs, images, Word files and Excel sheets are all reduced to
``SourceFile`` -> one or more ``Document`` objects, so classification,
extraction, validation, duplicate detection and reporting never need to know
where a document came from.
"""
from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Optional

# ---------------------------------------------------------------- doc types
TAX_INVOICE = "Tax Invoice"
REVISED_TAX_INVOICE = "Revised Tax Invoice"
CREDIT_NOTE = "Credit Note"
DEBIT_NOTE = "Debit Note"
PROFORMA = "Proforma Invoice"
QUOTATION = "Quotation / Estimate"
PURCHASE_ORDER = "Purchase Order"
SALES_ORDER = "Sales Order"
DELIVERY_CHALLAN = "Delivery Challan"
EWAY_BILL = "E-Way Bill"
PAYMENT = "Payment Advice / Receipt"
STATEMENT = "Statement of Account"
EXCEL_REGISTER = "Excel Invoice Register"
OTHER_SUPPORTING = "Other Supporting Document"
UNRELATED = "Unrelated Document"
UNKNOWN = "Unknown / Manual Review"

DOC_TYPES = [TAX_INVOICE, REVISED_TAX_INVOICE, CREDIT_NOTE, DEBIT_NOTE, PROFORMA,
             QUOTATION, PURCHASE_ORDER, SALES_ORDER, DELIVERY_CHALLAN, EWAY_BILL,
             PAYMENT, STATEMENT, EXCEL_REGISTER, OTHER_SUPPORTING, UNRELATED, UNKNOWN]
INVOICE_TYPES = {TAX_INVOICE, REVISED_TAX_INVOICE}
NOTE_TYPES = {CREDIT_NOTE, DEBIT_NOTE}
NON_TAX_TYPES = {PROFORMA, QUOTATION, PURCHASE_ORDER, SALES_ORDER, DELIVERY_CHALLAN,
                 EWAY_BILL, PAYMENT, STATEMENT, OTHER_SUPPORTING}

# ---------------------------------------------------------------- direction
INWARD = "Inward (supplier -> Caliber)"
OUTWARD = "Outward (Caliber -> customer)"
THIRD_PARTY = "Third party (Caliber not buyer)"
CONSIGNEE_ONLY = "Caliber is consignee only"
DIRECTION_UNKNOWN = "Not determined"

# ---------------------------------------------------------------- statuses
ST_VERIFIED = "Verified"
ST_REVIEW = "Review Required"
ST_DUPLICATE = "Duplicate Candidate"
ST_NON_TAX = "Non-Tax Document"
ST_FAILED = "Processing Failed"
ST_OUTWARD = "Outward (Caliber Sales)"
ST_EXCLUDED = "Excluded"

# The 45 invoice columns, in the exact order required by the specification.
INVOICE_COLUMNS = [
    "Source File Name", "Source File Path / Relative Path", "Page Number or Page Range",
    "Document Type", "Document Status", "Invoice Date", "Invoice Number",
    "Original Invoice Number", "Supplier Legal Name", "Supplier Trade Name",
    "Supplier GSTIN", "Supplier State", "Supplier State Code", "Supplier Address",
    "Customer Legal Name", "Customer Trade Name", "Customer GSTIN", "Customer State",
    "Customer State Code", "Customer Address", "Bill To Name", "Bill To GSTIN",
    "Ship To Name", "Ship To GSTIN", "Place of Supply", "Reverse Charge Indicator",
    "Currency", "Taxable Value", "IGST", "CGST", "SGST", "UTGST", "Cess",
    "Other Charges", "Discount", "Round-Off", "Invoice Total", "HSN/SAC",
    "Tax Rate or Tax Rates", "PO Number", "IRN", "E-Way Bill Number",
    "OCR / Extraction Confidence", "Validation Status", "Review Remarks",
]
# Extra traceability columns appended after the mandatory 45.
EXTRA_COLUMNS = ["Transaction Direction", "Caliber Registration (State)",
                 "GSTR-2B Match", "Duplicate Group", "Document ID", "File SHA-256"]

AMOUNT_FIELDS = ["taxable_value", "igst", "cgst", "sgst", "utgst", "cess",
                 "other_charges", "discount", "round_off", "invoice_total"]


@dataclass
class SourceFile:
    file_id: str               # sha256 of content
    name: str
    rel_path: str
    ext: str
    size: int
    stored_path: str
    kind: str = ""             # pdf | image | excel | word | unsupported
    status: str = "Pending"
    error: str = ""
    pages: int = 0
    modified: str = ""


@dataclass
class PageText:
    page_no: int               # 1-based
    text: str
    source: str                # embedded | embedded-sorted | ocr | excel | word
    ocr_confidence: Optional[float] = None
    rotation: int = 0
    quality: float = 1.0
    views: list = field(default_factory=list)   # alternative readings (sorted layout, OCR)
    copy_label: str = ""


@dataclass
class Document:
    doc_id: str                         # file_id:pX-Y or file_id:sheet
    file_id: str
    file_name: str
    rel_path: str
    pages: str                          # "1" / "1-2" / sheet name
    text: str
    text_source: str
    ocr_confidence: Optional[float] = None
    views: list = field(default_factory=list)   # alternative text readings, same pages
    doc_type: str = UNKNOWN
    type_confidence: float = 0.0
    type_reason: str = ""
    direction: str = DIRECTION_UNKNOWN
    caliber_gstin: str = ""
    caliber_state: str = ""
    fields: dict = field(default_factory=dict)
    field_confidence: dict = field(default_factory=dict)
    line_items: list = field(default_factory=list)      # tax-rate breakdown rows
    validations: list = field(default_factory=list)     # [{rule, level, message}]
    status: str = ST_REVIEW
    remarks: list = field(default_factory=list)
    extraction_confidence: float = 0.0
    copies_in_file: int = 1
    duplicate_group: str = ""
    duplicate_role: str = ""
    gstr2b_match: str = ""
    register_rows: list = field(default_factory=list)   # for Excel registers

    def to_dict(self) -> dict:
        return asdict(self)

    @staticmethod
    def from_dict(d: dict) -> "Document":
        return Document(**d)

    def add_validation(self, rule: str, level: str, message: str) -> None:
        """level: ERROR | WARNING | INFO"""
        self.validations.append({"rule": rule, "level": level, "message": message})

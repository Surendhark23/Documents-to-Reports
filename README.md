# Caliber GST Document Processor

Turns a mixed folder or ZIP of supplier documents for **Caliber Inno Enterprise Pvt Ltd** into a validated, state-wise GST inward-invoice register. The input can hold PDFs, scans, phone photos, Excel/CSV and Word files, and may include duplicates, quotations, proformas, POs, challans, notes and unrelated files. The output is one Excel workbook.

Built and tested against the real sample `Jun-26.zip` (144 files).

> **Online access (v1.2.0):** deploy the app from GitHub to Streamlit Cloud and publish an HTML dashboard on GitHub Pages - step by step in `docs/DEPLOY_ONLINE.md`. `python dashboard.py <register.xlsx>` builds the dashboard page.
>
> **Version 1.1.1 (audit, 1 Nov 2026)** – re-validated on the Nov-2026 batch (79 files, 82 documents) against a hand-checked
> ground truth: documents fully correct 65.9% → 87.8%, field accuracy 90.8% → 95.4%, auto-Verified records with wrong data 4 → 0,
> false duplicates 5 → 0, no invoice counted twice. New GSTIN_<Caliber GSTIN> sheets in GSTR-2B layout for checking against the GST portal. See `docs/ERROR_RESOLUTION_REPORT.md`, `docs/ENHANCEMENT_REPORT.md`, `docs/TEST_RESULTS.md`,
> `docs/SAMPLE_ANALYSIS.md` (input validation) and `docs/EXECUTION_INSTRUCTIONS.md`. The workbook now opens with an
> **Executive_Summary** sheet. The Jun-26 figures below are from v1.0.

| Result on the sample (manually verified ground truth) | |
|---|---|
| Documents classified correctly | 198 / 206 (96.1%) – all 8 misses go to *manual review*, none become tax invoices |
| Non-invoices wrongly treated as tax invoices | **0** |
| Buyer/supplier direction (Caliber buying vs selling) | 136 / 136 |
| Field accuracy, 24 unique inward invoices × 9 key fields | 214 / 216 (99.1%) |
| Auto-verified invoices with any field error | **0 of 20** |
| Processing time, cold / cached | 148 s / 3 s (2 CPU cores, OCR on) |

These figures come from one month of documents. They are not a guarantee for other suppliers' layouts. Every uncertain record still goes through human review.

---

## 1. Install

### Windows (local use)
1. Install **Python 3.11–3.13** from python.org. Tick *"Add python.exe to PATH"* during setup.
2. Install **Tesseract OCR**. Use the UB-Mannheim build: <https://github.com/UB-Mannheim/tesseract/wiki>. With the default path `C:\Program Files\Tesseract-OCR\` it is found automatically; otherwise set `TESSERACT_CMD` (see §4).
3. *(Optional, for legacy `.doc` files only)* install LibreOffice.
4. In a command prompt inside this folder:
   ```bat
   python -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   streamlit run app.py
   ```
   The app opens at <http://localhost:8501>.

### Linux / macOS
```bash
sudo apt-get install -y tesseract-ocr tesseract-ocr-eng      # macOS: brew install tesseract
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Without Tesseract the app still runs. Text-based PDFs, Excel and Word files work normally, while scans and photos are listed under *Processing_Errors* as "needs OCR".

### Command line (no UI)
```bash
python run_cli.py Jun-26.zip --out Jun-26_register.xlsx      # or a folder path
python tests/evaluate.py Jun-26.zip --md results.md           # accuracy vs ground truth
python -m pytest -q tests/test_core.py                        # 21 unit tests
```

---

## 2. Monthly workflow

1. **Upload & Processing.** Drop the month's ZIP, or individual files. Sub-folders are processed recursively. On a local install you can type a folder path instead.
2. **Caliber GST Master.** The app proposes Caliber's registrations, using *only* GSTINs it found on the documents, and lists the files it found each one in. Check each one and set *Verified*, or correct/reject it. You can also import the master from Excel. Until a registration is verified, invoices billed to it stay in review.
3. **Document Review.** Work through *Review Required*. Each document shows the page image, the extracted fields with confidence, and the validation messages. Correct fields, change the document type if needed, then **Save & mark verified**. Every change is written to the audit trail.
4. **Duplicate Review.** Compare candidates side by side and decide for each: *Exact duplicate*, *Revised invoice* or *Separate transaction*.
5. **Reports & Export.** Download the consolidated workbook, single sheets as CSV/Excel, the audit log or the exceptions report.

Corrections and decisions are stored in SQLite and re-applied automatically. Re-processing a batch never loses review work, and unchanged files are not re-read (the cache is keyed by the file's SHA-256).

---

## 3. Output workbook

| # | Sheet | Content |
|---|---|---|
| 1 | Final_Tax_Invoices | Verified **inward** tax invoices only, with the 45 required columns (in the specified order) + 6 trace columns; totals via `SUBTOTAL` |
| 2 | Pending_Review | Anything missing, conflicting, low-confidence, unverified, or not billed to Caliber – with the issues listed |
| 3 | Document_Register | Every file and every document inside it, with type, status and the reason it is included or excluded |
| 4 | Non_Tax_Documents | Quotations, proformas, POs, challans, e-way bills, payments, statements, registers, unrelated files |
| 5 | Credit_Debit_Notes | Notes, kept separate and linked to their original invoice when it is in the batch |
| 6 | Duplicate_Report | L1–L4 candidates, primary vs candidate, reason, confidence, suggested action, user decision |
| 7 | Caliber_GST_Master | State-wise registrations with verification status and the source documents for each |
| 8 | Supplier_Master | Supplier GSTIN, name (with variants), state and address |
| 9 | Tax_Summary | By Caliber state and by supplier (final invoices); by type/direction/status (informational) |
| 10 | Processing_Errors | Rejected uploads, unreadable or unsupported files, OCR failures, low-confidence OCR |
| + | Outward_Sales_Register | Caliber's **own** sales invoices found in the batch, kept out of the purchase register |
| + | Tax_Rate_Breakdown | Rate-wise taxable/tax per invoice (multi-rate invoices) |
| + | GSTR2B_Register_Rows | Rows read from GSTR-2B downloads; each PDF invoice shows whether it appears in 2B and whether the values agree |

Dates are real Excel dates formatted DD-MM-YYYY. GSTINs and invoice numbers are text, so leading zeros are kept. Missing amounts stay **blank**, never 0.

---

## 4. Configuration

All limits are environment variables prefixed with `CALIBER_`. They can also be changed in the app sidebar.

| Variable | Default | Meaning |
|---|---|---|
| `CALIBER_DATA_DIR` | `./data` | Database, stored uploads, exports. Point this at persistent storage on a server |
| `CALIBER_MAX_UPLOAD_MB` / `CALIBER_MAX_FILE_MB` | 500 / 50 | Batch and per-file size limits (also set `server.maxUploadSize` in `.streamlit/config.toml`) |
| `CALIBER_MAX_ZIP_MEMBERS` / `CALIBER_MAX_ZIP_RATIO` | 5000 / 200 | ZIP-bomb guards |
| `CALIBER_MAX_PDF_PAGES` | 60 | Pages processed per PDF |
| `CALIBER_OCR_ENABLED` / `CALIBER_OCR_DPI` / `CALIBER_OCR_TIMEOUT_S` | true / 300 / 90 | OCR controls |
| `CALIBER_MAX_WORKERS` | min(4, CPUs) | Parallel extraction threads |
| `CALIBER_AMOUNT_TOLERANCE` | 1.0 | Rupee tolerance for total = taxable + taxes + round-off |
| `CALIBER_MIN_CONFIDENCE_FINAL` | 0.80 | Below this, a record is never auto-verified |
| `CALIBER_INCLUDE_OUTWARD_IN_FINAL` | false | Put Caliber's own sales invoices into Final_Tax_Invoices |
| `CALIBER_CALIBER_PANS` | `AAFCC7149G` | PAN(s) used to recognise Caliber GSTINs (comma separated) |
| `CALIBER_CLOUD` | – | Set to `1` on hosted deployments to hide the local-folder input |
| `TESSERACT_CMD` | – | Full path to `tesseract.exe` if it is not on PATH (Windows) |

On Windows the default Tesseract install folders are detected automatically. Set `TESSERACT_CMD` only if you installed it somewhere else.

### External AI
External AI is **off by default and not used anywhere in the pipeline.** No document content leaves the machine. `config.py` has `ai_enabled`, `ai_provider` and `ai_max_docs_per_job` placeholders. They are there so that an approved vision model can later be added for the *Unknown / Manual Review* queue only. Doing that would send confidential supplier documents to a third party, so it needs a written approval, a data-processing agreement, a per-batch document cap for cost control, and logging of every document sent.

---

## 5. Deployment

### Streamlit Community Cloud
1. Push this folder to a **private** GitHub repository. Never commit `data/`; it is excluded by `.gitignore`.
2. Create the app at <https://share.streamlit.io>, with `app.py` as the entry point. `packages.txt` installs Tesseract and `requirements.txt` installs the Python packages.
3. Under *Secrets / environment*, set `CALIBER_CLOUD = "1"`.
4. Restrict viewers to named Google accounts. The documents are confidential.

**Storage limitation:** the Community Cloud filesystem is *ephemeral*. The SQLite database, master, review decisions and stored uploads are lost whenever the app restarts or sleeps. Use it for demos only, or export the master and workbook after every session and re-import the master next time. For real use, deploy on a server with persistent storage.

### Own server / VM (recommended for production)
```bash
export CALIBER_DATA_DIR=/srv/caliber-data          # persistent disk, backed up
export CALIBER_CLOUD=1
streamlit run app.py --server.port 8501 --server.address 0.0.0.0
```
Put it behind a reverse proxy with HTTPS and authentication (nginx + basic auth / SSO). Run it as a service (systemd or a Docker container with `CALIBER_DATA_DIR` on a mounted volume), and back up `caliber.db` daily. For several concurrent users, move from SQLite to PostgreSQL. Only `database.py` needs to change.

### Folder upload in the browser
Browsers only hand individual files, or a whole folder in Chromium-based browsers, to a web page. The reliable path everywhere is: **zip the folder → upload the ZIP**. Sub-folders are kept. On a local install, you can type a folder path instead.

---

## 6. Architecture (Phase 2)

```
app.py ─ Streamlit UI (6 pages)            run_cli.py ─ batch runner
   │
pipeline.py ── run_job(): inventory → parallel extract (cached) → segment → classify
   │                     → master detection → field extraction → finalize()
   │           finalize(): overrides → validation → GSTR-2B match → duplicates → status
   │           revalidate(): finalize() on stored documents (after every correction)
   ├─ file_manager.py      safe ZIP/upload/folder ingest, SHA-256 inventory
   ├─ pdf_extractor.py     PDF text (native + layout-sorted views), OCR fallback, Word
   ├─ ocr_engine.py        Tesseract + OpenCV (rotation, deskew, denoise), confidence
   ├─ excel_extractor.py   GSTR-2B parser, generic registers, sheet→text
   ├─ document_classifier.py  16 document types, title-first rules, page segmentation
   ├─ field_extractor.py   party roles (bill/ship/dispatch/supplier), labels, amount reconciliation
   ├─ caliber_master.py    evidence-based state-wise GSTIN master, Excel import/export
   ├─ gst_validator.py     GSTIN checksum/OCR repair, state codes, GST rules, status
   ├─ duplicate_detector.py L1 hash · L2 GSTIN+number · L3 +values · L4 fuzzy; revisions, note links
   ├─ review_manager.py    corrections, approvals, duplicate decisions (audit trail)
   ├─ report_generator.py  13-sheet workbook
   ├─ database.py          SQLite: jobs, files, documents, cache, master, corrections, audit
   ├─ models.py            normalised Document / SourceFile structures, 45-column schema
   └─ config.py            limits and settings (env-overridable)
```

**Key design decisions** (all driven by the sample):
- **Direction first.** About 60% of the sample is Caliber's *own* outward documents. Every document gets a direction: inward, outward, third-party, or Caliber-only-as-consignee. Only inward invoices billed to a verified Caliber GSTIN can reach the final register.
- **Title beats keywords.** A "Tax Invoice" heading never wins over proforma or quotation wording. A proforma with an IRN goes to review.
- **Several text views per page.** These are the PDF's own reading order, a layout-sorted order, and OCR text. OCR is used only for missing or garbled text, such as broken font encodings or scans.
- **Amounts must balance.** Every labelled amount, tax-summary row and *taxable × rate = tax* triplet is a candidate. The combination that balances wins; if nothing balances, the record is flagged and no number is changed.
- **Never silently drop.** Duplicates, third-party invoices and unreadable files all appear in the workbook with a reason.

---

## 7. Known limitations
- Phone photos of transport receipts and crumpled challans OCR poorly. They are classified *Unknown / Manual Review* rather than guessed.
- A revision is recognised from the document heading, from same-number-different-values pairs, or (with a review flag) from a file name containing "revised". The link to the original invoice is made only when the document prints it.
- The tax-rate breakdown is derived from rows where taxable × rate = tax. When those rows don't add up to the invoice taxable value, it is marked *Partial*.
- Place-of-supply findings are warnings only. The app does not decide GST treatment that depends on facts not printed on the document.

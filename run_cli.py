"""Command-line batch runner (no UI).

    python run_cli.py path/to/Jun-26.zip  [--out report.xlsx] [--workers 4] [--no-ocr]
    python run_cli.py path/to/folder      [--out report.xlsx]

Prints a per-document summary and writes the consolidated workbook."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from config import default_settings
from file_manager import new_job_id, ingest_uploads, ingest_folder
from pipeline import run_job


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("source", help="ZIP file, single document or folder")
    ap.add_argument("--out", default="", help="output .xlsx path (default: data/exports/<job>.xlsx)")
    ap.add_argument("--workers", type=int, default=0)
    ap.add_argument("--no-ocr", action="store_true")
    ap.add_argument("--no-cache", action="store_true")
    ap.add_argument("--quiet", action="store_true")
    a = ap.parse_args(argv)

    s = default_settings()
    if a.workers:
        s.max_workers = a.workers
    if a.no_ocr:
        s.ocr_enabled = False
    if a.no_cache:
        s.use_cache = False
    src = Path(a.source)
    job = new_job_id()
    if src.is_dir():
        problems = ingest_folder(job, src, s)
    else:
        problems = ingest_uploads(job, [(src.name, src.read_bytes())], s)
    for p in problems:
        print(f"[upload] {p['path']}: {p['error']}", file=sys.stderr)

    def prog(done, total, msg):
        if not a.quiet:
            print(f"  [{done}/{total}] {msg}", file=sys.stderr)

    res = run_job(job, s, prog)
    from report_generator import build_workbook
    out = Path(a.out) if a.out else None
    path = build_workbook(res, s, out_path=out, upload_problems=problems)
    if not a.quiet:
        for d in res["documents"]:
            f = d.fields
            print(f"{d.status[:14]:14} | {d.doc_type[:18]:18} | {d.direction[:8]:8} | {d.rel_path[:55]:55} p{d.pages:5} | "
                  f"{f.get('invoice_number', '')!s:18} | {f.get('invoice_date', '')!s:10} | {f.get('supplier_gstin', '')!s:15} | "
                  f"{f.get('customer_gstin', '')!s:15} | {f.get('taxable_value')!s:>10} | {f.get('igst')!s:>9} | "
                  f"{f.get('cgst')!s:>9} | {f.get('invoice_total')!s:>10} | {d.extraction_confidence}")
    print(f"Job {job}: {len(res['files'])} files, {len(res['documents'])} documents -> {path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

"""State-wise GST registration master for Caliber Inno Enterprise Pvt Ltd.

The master is *evidence based*: entries are proposed only from GSTINs that
actually appear on the processed documents (and that share Caliber's PAN or
are printed next to Caliber's name).  Every proposed entry starts as
'Detected - Unverified' with the documents it was found in; a user verifies,
edits or deactivates it on the 'Caliber GST Master' page or by importing an
Excel sheet.  No registration is ever invented or assumed for a state.
"""
from __future__ import annotations

import re
from collections import Counter, defaultdict
from typing import Iterable, Optional

import pandas as pd

import database as db
from gst_validator import GSTIN_RE, gstin_valid, gstin_pan, state_name

MASTER_COLUMNS = ["Customer Legal Name", "Customer Trade Name", "GSTIN", "State Name", "State Code",
                  "Registered Address", "Active/Inactive", "Effective From", "Effective To",
                  "Verification Status", "Source Document", "Remarks"]
_DB_TO_COL = {"legal_name": "Customer Legal Name", "trade_name": "Customer Trade Name", "gstin": "GSTIN",
              "state_name": "State Name", "state_code": "State Code", "address": "Registered Address",
              "active": "Active/Inactive", "effective_from": "Effective From", "effective_to": "Effective To",
              "verification_status": "Verification Status", "source_document": "Source Document",
              "remarks": "Remarks"}
VERIFIED = "Verified"
DETECTED = "Detected - Unverified"
REJECTED = "Rejected"


class CaliberMaster:
    def __init__(self, settings):
        self.settings = settings
        self.name_rx = re.compile(settings.caliber_name_regex, re.I)
        self.pans = {p.strip().upper() for p in settings.caliber_pans if p.strip()}
        self.entries: dict[str, dict] = {}
        self.reload()

    # ----------------------------------------------------------- persistence
    def reload(self) -> None:
        with db.connect() as con:
            rows = [dict(r) for r in con.execute("SELECT * FROM caliber_master")]
        self.entries = {}
        for r in rows:
            r["active"] = bool(r["active"])
            self.entries[r["gstin"]] = r

    def upsert(self, entry: dict, user: str = "system") -> None:
        g = entry["gstin"].strip().upper()
        cur = self.entries.get(g, {})
        changes = {k: v for k, v in entry.items() if cur.get(k) != v}
        if cur and not changes:
            return      # audit fix E12: nothing changed - no rewrite and no empty audit-log row
        merged = {**cur, **entry, "gstin": g}
        merged.setdefault("state_code", g[:2])
        merged.setdefault("state_name", state_name(g))
        with db.connect() as con:
            con.execute("""INSERT OR REPLACE INTO caliber_master VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (g, merged.get("legal_name", ""), merged.get("trade_name", ""),
                         merged.get("state_name", ""), merged.get("state_code", ""), merged.get("address", ""),
                         1 if merged.get("active", True) else 0, merged.get("effective_from", ""),
                         merged.get("effective_to", ""), merged.get("verification_status", DETECTED),
                         merged.get("source_document", ""), merged.get("remarks", ""), db.now()))
        db.audit("", "master_upsert", {"gstin": g, "by": user, "changes": changes})
        self.reload()

    def delete(self, gstin: str) -> None:
        with db.connect() as con:
            con.execute("DELETE FROM caliber_master WHERE gstin=?", (gstin,))
        db.audit("", "master_delete", {"gstin": gstin})
        self.reload()

    # ----------------------------------------------------------- queries
    def lookup(self, gstin: str) -> Optional[dict]:
        e = self.entries.get((gstin or "").upper())
        if e and e.get("verification_status") == REJECTED:
            return None
        return e

    def is_caliber_gstin(self, gstin: str) -> bool:
        g = (gstin or "").upper()
        if g in self.entries and self.entries[g].get("verification_status") != REJECTED:
            return True
        return gstin_pan(g) in self.pans

    def all_gstins(self) -> set[str]:
        return {g for g, e in self.entries.items() if e.get("verification_status") != REJECTED}

    def by_state(self, code: str) -> list[dict]:
        return [e for e in self.entries.values() if e.get("state_code") == code]

    # ----------------------------------------------------------- detection
    def detect_from_documents(self, docs: Iterable, gstr2b_owners: dict[str, str] | None = None) -> list[dict]:
        """Propose master entries from documents. Returns proposals (not saved)."""
        evidence: dict[str, dict] = defaultdict(lambda: {"docs": [], "names": Counter(), "addresses": Counter()})
        for d in docs:
            text = d.text if hasattr(d, "text") else d.get("text", "")
            path = d.rel_path if hasattr(d, "rel_path") else d.get("rel_path", "")
            fields = d.fields if hasattr(d, "fields") else d.get("fields", {})
            for g in set(GSTIN_RE.findall((text or "").upper())):
                if not gstin_valid(g):
                    continue
                # Evidence: Caliber's PAN, or the GSTIN printed as the bill-to/ship-to
                # GSTIN directly under a Caliber name (a registration with a PAN that is
                # not configured yet is still proposed - the user decides).
                billed_as_caliber = any(
                    fields.get(k) == g and self.name_rx.search(fields.get(n) or "")
                    for k, n in (("bill_to_gstin", "bill_to_name"), ("ship_to_gstin", "ship_to_name")))
                if gstin_pan(g) in self.pans or billed_as_caliber:
                    ev = evidence[g]
                    if path not in ev["docs"]:
                        ev["docs"].append(path)
                    addr = fields.get("customer_address") or ""
                    # audit fix E13: a company-name line ('CALIBER INNO ENTERPRISES PVT. LTD.') was stored
                    # as the registered address; only keep text that looks like an address
                    if fields.get("customer_gstin") == g and re.search(r"\d|road|street|floor|nagar|block|main|layout", addr, re.I):
                        ev["addresses"][addr[:200]] += 1
                    for m in self.name_rx.finditer(text):
                        line = text[m.start():text.find("\n", m.start()) if text.find("\n", m.start()) > 0 else None]
                        ev["names"][re.sub(r"\s+", " ", line.strip())[:80].upper()] += 1
        for g, label in (gstr2b_owners or {}).items():
            evidence[g]["docs"].append(label)
            evidence[g]["gstr2b"] = True
        proposals = []
        for g, ev in sorted(evidence.items()):
            legal = "CALIBER INNO ENTERPRISE PRIVATE LIMITED"
            remarks = f"Seen in {len(ev['docs'])} document(s)"
            if ev.get("gstr2b"):
                remarks += "; owner GSTIN of an official GSTR-2B download"
            name_variants = [n for n, _ in ev["names"].most_common(4)]
            if name_variants:
                remarks += "; name variants: " + " | ".join(name_variants)
            proposals.append({
                "gstin": g, "legal_name": legal, "trade_name": "", "state_code": g[:2],
                "state_name": state_name(g),
                "address": (ev["addresses"].most_common(1)[0][0] if ev["addresses"] else ""),
                "active": True, "effective_from": "", "effective_to": "",
                "verification_status": DETECTED,
                "source_document": "; ".join(ev["docs"][:5]) + (" ..." if len(ev["docs"]) > 5 else ""),
                "remarks": remarks,
            })
        return proposals

    def apply_proposals(self, proposals: list[dict]) -> int:
        """Insert proposals that are not already in the master (never overwrite user edits)."""
        n = 0
        for p in proposals:
            if p["gstin"] in self.entries:
                cur = self.entries[p["gstin"]]
                if cur.get("verification_status") == DETECTED:
                    upd = {"source_document": p["source_document"], "remarks": p["remarks"]}
                    if not cur.get("address") and p.get("address"):
                        upd["address"] = p["address"]
                    self.upsert({"gstin": p["gstin"], **upd})
                continue
            self.upsert(p)
            n += 1
        return n

    # ----------------------------------------------------------- excel I/O
    def to_dataframe(self) -> pd.DataFrame:
        rows = []
        for e in sorted(self.entries.values(), key=lambda x: x["state_code"]):
            row = {col: e.get(k, "") for k, col in _DB_TO_COL.items()}
            row["Active/Inactive"] = "Active" if e.get("active", True) else "Inactive"
            rows.append(row)
        return pd.DataFrame(rows, columns=MASTER_COLUMNS)

    def import_excel(self, path_or_buffer) -> tuple[int, list[str]]:
        df = pd.read_excel(path_or_buffer, dtype=str).fillna("")
        missing = [c for c in ("GSTIN",) if c not in df.columns]
        if missing:
            return 0, [f"Missing column(s): {', '.join(missing)}"]
        errors, n = [], 0
        col_to_db = {v: k for k, v in _DB_TO_COL.items()}
        for i, row in df.iterrows():
            g = str(row["GSTIN"]).strip().upper()
            if not gstin_valid(g):
                errors.append(f"Row {i + 2}: GSTIN '{g}' is invalid (format/checksum) - skipped")
                continue
            entry = {col_to_db[c]: str(row[c]).strip() for c in df.columns if c in col_to_db}
            entry["gstin"] = g
            entry["active"] = str(row.get("Active/Inactive", "Active")).strip().lower() != "inactive"
            entry.setdefault("verification_status", VERIFIED)
            if not entry.get("verification_status"):
                entry["verification_status"] = VERIFIED
            if entry.get("state_code") and entry["state_code"] != g[:2]:
                errors.append(f"Row {i + 2}: State Code {entry['state_code']} does not match GSTIN prefix {g[:2]} - corrected to {g[:2]}")
            entry["state_code"] = g[:2]
            entry["state_name"] = entry.get("state_name") or state_name(g)
            self.upsert(entry, user="excel-import")
            n += 1
        return n, errors

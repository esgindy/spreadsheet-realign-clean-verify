# AGENTS.md — Spreadsheet Realign, Clean & Verify (`spreadsheet-realign-clean-verify`)

Read [`SKILL.md`](./SKILL.md) before modifying, auditing, or cleaning any multi-source Excel (`.xlsx`/`.xls`) or CSV dataset.

## Core Operating Rules for Any AI Coding Agent

1. **Run the Pre-Flight Interview First (Stage 0):** Before generating cleaned files, confirm output format (`.xlsx`, `.csv` `utf-8-sig`, or Both), destination save folder (same folder vs. custom path — never overwrite raw originals), granularity/split mode, hidden rows/sheets handling (skip by default vs. include), and whether to emit a `_needs_review.csv` quarantine log.
2. **Never assume uniform column schemas within an Excel sheet.** Inspect rows dynamically for mid-sheet column shifts, multi-form stacked blocks, and data accidentally placed in header rows.
3. **Use a Two-Stage Semantic Pass (Reader -> Realigner):**
   - First extract flagged anomalies + unique ambiguous cell values alongside original raw row snapshots (`scripts/profile_and_extract.py`).
   - Classify ambiguous bilingual (Arabic/English) cell values semantically before moving them into target columns.
4. **Never drop reachable contacts unnecessarily, and never keep empty noise:** Every retained lead row must contain a valid >=7-digit phone number OR a valid email address after semantic extraction.
5. **Enforce Entity Granularity:** Keep distinct owned properties/orders/units as individual filterable rows (one row per unique unit key) instead of concatenating them into `|`-delimited single-cell blobs.
6. **Apply Strict `TRIM` to Every Cell:** After all text cleanup and summarization, strip leading/trailing whitespace, collapse repeated internal spaces, and strip invisible Unicode control/bidi marks (`\xa0`, `\u200e`, `\u200f`, `\u202a–\u202e`).
7. **Honor Format & Encoding Standards:** Export in the requested format(s) using `utf-8-sig` (UTF-8 BOM) for `.csv` so Arabic/Unicode displays cleanly in Excel, and native `yyyy-mm-dd` dates + AutoFilter in `.xlsx`.
8. **Always Run Source Cross-Verification & End with a Numeric Summary Report:** Cross-check every cleaned phone number back against original source files for Name, Email, Country/City, Phone, and WhatsApp, then present a concise factual numeric audit table (funnel + repairs + verification + zero-defect check).
9. **Protect PII:** Never store real customer phone numbers, emails, or production datasets inside this skill repository.

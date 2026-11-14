---
name: spreadsheet-realign-clean-verify
description: >-
  Industrial-grade, cross-agent spreadsheet cleaning, semantic realignment, universal TRIM, and
  source verification for messy or large-scale (50k-1M+ rows) Excel (.xlsx/.xlsm/.xls/.xlsb) and
  CSV files. Starts with an interactive pre-flight interview (output format, save folder, split
  granularity, quarantine log, hidden rows/sheets policy), streams big sheets memory-safely
  (read_only/write_only + reset_dimensions + auto CSV delimiter & multiline sniffer), recovers
  merged-cell ffill gaps, stacked 2-tier headers, banner titles, hyperlink-embedded contacts
  (wa.me/tel/mailto), cell comments, Mac 1904 serial dates, phantom columns, header-as-data rows,
  and mid-sheet column shifts, translates Arabic/English free text and locale currencies while
  protecting country codes, quarantines fake placeholders & junk into _needs_review.csv, scrubs
  Excel errors, neutralizes CSV formula injection, exports styled XLSX + UTF-8 BOM CSV, cross-
  verifies 100% of phones back to raw sources, and outputs a concise numeric audit report.
author: Eslam Elgindy
version: 1.2.0
license: MIT
---

# Spreadsheet Realign, Clean & Verify (`spreadsheet-realign-clean-verify`)

**Author:** Eslam Elgindy  
**Compatibility:** Cross-Agent Open Standard (`SKILL.md` + `AGENTS.md` + standalone Python 3.10+). Works across **Antigravity, Claude Code, Cursor, OpenAI Codex CLI, Gemini CLI, Windsurf, and GitHub Copilot**.

## Problem Statement (21 Real-World & Big-Data Failures Solved)

Naive column-mapping scripts (`col_A -> Name`, `col_B -> Phone`) fail silently on messy exports or crash with Out-Of-Memory errors on massive workbooks. This skill solves **21 verified production failure modes**:

1. **Mid-sheet schema shifts & stacked sub-schemas:** Rows `1..13` follow one layout while `14+` shift without new headers, or different ad forms stack under static legacy column names (`LANG`, `REGION`, `when`).
2. **Header-row-as-data & merged cell `ffill` gaps:** Row 1 holds real customer records (creating `.1` suffixes), or vertically merged categories leave blank `NaN` rows underneath.
3. **Big-data crashes (`50k–1M+` rows) & phantom 200+ empty columns:** Default Excel readers bloat RAM `20x–50x` or stall on stale XML dimensions / empty formatting tails.
4. **Semantic scrambling & embedded dividers:** Questions, budgets, emails, `Total`/`Subtotal`, or single-letter junk (`Y`, `R`) inside `Phone`/`WhatsApp`/`Country`/`Time`.
5. **Entity collapse (`|` concatenation):** Distinct assets/units owned by one contact flattened into `|`-delimited blobs instead of clean unit-level rows.
6. **Excel corruption, serial dates (1900/1904) & security leaks:** `#REF!`/`#ERROR!`, scientific notation (`9.715E+11`), stripped leading zeros, Windows `1900` vs. Mac `1904` serial date shifts (`1,462` days), and CSV formula injection (`=`, `+`, `-`, `@`).
7. **Whitespace, bidi, Arabic variants & encoding Mojibake:** `\xa0`, invisible `\u200e`/`\u202d`, `أ/إ/آ` vs `ا` & `ة` vs `ه` breaking deduplication, and Arabic CSV text scrambling without UTF-8 BOM (`utf-8-sig`).
8. **Hidden rows/sheets, buried hyperlink contacts & comments:** Manually hidden/filtered rows resurrecting in `iter_rows`, `"Chat"`/`"Email"` cells hiding `wa.me/`/`tel:`/`mailto:` links, and agent remarks hidden in `cell.comment`.
9. **Multi-row stacked headers, title banners & mixed CSV dialects:** Sparse banner titles, 2-tier merged group headers (`header=[0, 1]`), semicolon/tab CSVs, multi-line `Alt+Enter` cells, and legacy `.xls`/`.xlsb` binaries.
10. **Mixed locale currencies & fake placeholder spam:** European vs. US number formats (`1.200,50 €` vs `$1,200.50`, `2.5M`), plus synthetic placeholder junk (`0000000`, `1234567`, `test@test.com`) slipping past naive length/regex checks.

---

## Solution Architecture: The Two-Agent Protocol

Do **not** blindly clean rows in a single pass. Always execute the pipeline below:

### Stage 0 — Mandatory Pre-Flight Interview (Before Modifying or Writing Output)
Unless the user already specified all preferences in their request, ask concise pre-flight questions first (using your native interactive question tool if available, or short numbered text questions):
1. **Output Format:** Do you want **Excel (`.xlsx`)**, **CSV (`utf-8-sig` for Arabic/Unicode Excel compatibility)**, or **Both** (Recommended)?
2. **Output Save Location:** Save inside the **same source folder**, a clean subfolder (e.g., `cleaned_output/`), or a **custom path**? *(Never overwrite original raw files.)*
3. **Granularity & Split Mode:** Merge by person/contact only, preserve **1 row per unique asset/order unit** for recurring entities, or **split into separate files** (e.g., Leads vs. Owners)?
4. **Quarantine Unreachable Rows:** Drop rows lacking both valid phone & valid email silently, or also export a `_needs_review.csv` quarantine log explaining why each row was excluded?
5. **Hidden Rows / Sheets Policy (if `.xlsx`/`.xls`):** Skip hidden/filtered rows and hidden tabs automatically (Default), or include them?

### Stage 1 — Deterministic Profiling, Big-Data Streaming & Raw Indexing
- **Adaptive Big-Data Mode:** Automatically switch to `openpyxl(..., read_only=True, data_only=True)` (plus `ws.reset_dimensions()` / `Workbook(write_only=True)` and chunked CSV reads) whenever files exceed ~25,000 rows or `--stream-mb 10` MB to avoid Out-Of-Memory (OOM) crashes. Supports `.xls`/`.xlsb` when `xlrd`/`pyxlsb` is installed.
- **Hidden Filtering, Link Rescue, Stacked Headers & Epoch Check:** Skip `sheet_state != "visible"` and `row_dimensions[r].hidden` by default (`--include-hidden-rows` / `--include-hidden-sheets` overrides); extract `cell.hyperlink.target` (`wa.me/`, `tel:`, `mailto:`, `=HYPERLINK(...)`) and `cell.comment.text`; detect `wb.epoch` (`1900` vs. Mac `1904`); skip sparse banner titles and flatten 2-row merged stacked headers cleanly.
- **Resilient CSV & Encoding Ingest:** Sniff `,`, `;`, `\t`, and `|` over `16 KB` samples while keeping `Alt+Enter` multi-line quoted cells intact; auto-fallback across `utf-8-sig` -> `utf-8` -> `cp1256` -> `latin1`.
- **Structural Hygiene & Header-as-Data Recovery:** Force string types (`dtype=str`) so IDs/phones never collapse (`9.715E+11`) or lose leading zeros; unmerge & forward-fill (`ffill`) grouped cells (`<stream-mb`); strip phantom empty overflow columns; detect `header_as_data` (`@`, phones, dates in row 1) and flag `FAKE_PLACEHOLDER` or mid-sheet column shifts by value pattern.
- **Raw Indexing:** Build compact phone-tail and email indexes linking every record (including rescued hyperlinks/comments) to `(file, sheet, excel_row)`, dumping flagged rows to `audit_batch.json`.

### Stage 2 — Agent 1: `Cell_Semantic_Reader` (Semantic Classification)
Agent 1 inspects the flagged batch and ambiguous unique values across both original and raw-source views, emitting deterministic realignment decisions (`agent1_semantic_decisions.json`):
- **Names & Divider Filtering:** Distinguish real human names vs. product/unit types (`Studio`, `1-Bedroom`, `غرفة`), section dividers (`Tower ( A )`), `Total`/`Subtotal` footers, placeholders (`duplicae`), customer questions, and keyboard noise.
- **Phones & WhatsApp:** Force text preservation (prevent scientific notation `9.715E+11` & dropped leading zeros), scrub `#ERROR!` / `#REF!` / `#N/A`, convert localized digits (`٠١٢٣٤٥٦٧٨٩` -> ASCII), split composite delimiters (`nullI971|...`, `/`, `;`, `-`) without breaking valid international codes (`+44`, `+1`), and relocate non-phone text/questions to semantic notes.
- **Emails & Domain Typo Audit:** Extract valid addresses; quarantine non-email strings into their true fields; flag probable domain typos (`gamil.com`, `icloude.com`, `.con`) in an audit note instead of corrupting contact records silently.
- **Campaign / Project / Country / Time / Budget:** Disentangle ad audience tags (`ARABIC`, `GCC`, `EUROPE`), split timestamp prefixes from campaign titles, remove affirmative prefixes (`نعم - Riyadh`), resolve swapped `when`/`budget`/`Feedback` form fields, and classify true contact times vs. budgets/units.
- **Keep vs. Quarantine Gate:** Retain an entity only if it passes the reachability rule (valid phone `>=7` digits **OR** valid email). Route rejected/unrecoverable rows to `_needs_review.csv` (`utf-8-sig`) with an `exclusion_reason` column.

### Stage 3 — Agent 2: `Row_Realigner_Verifier` (Reconstruction, Split, TRIM & Export)
Agent 2 reads `agent1_semantic_decisions.json` + extracted base records and executes:
1. **Apply Realignment Decisions:** Move every value into its true semantic column and strip `|`-repeated duplicate tokens.
2. **Entity Separation:** Keep transactional/prospect leads separate from recurring asset/owner portfolios (1 row per unique asset/unit key rather than `|`-collapsed multi-unit rows).
3. **Post-Clean Re-Deduplication (Arabic Key Normalization):** Map `أ/إ/آ -> ا`, `ى -> ي`, `ة -> ه` on deduplication/grouping keys **only** (retaining proper spelling in cells), and re-merge duplicates revealed after unshifting phones/emails.
4. **Universal `strict_trim` & Safety Pass (Mandatory on Every Cell):**
   - Strip `\xa0`, `\t`, `\r`, `\n`, and invisible Unicode directional characters (`\u200b`–`\u200f`, `\u202a`–`\u202e`, `\ufeff`).
   - Collapse internal consecutive whitespace to a single space, scrub Excel errors (`#REF!`, `#ERROR!`, `#N/A`), and neutralize CSV formula-injection risk (`=`, `+`, `-`, `@` on non-phone free text).
5. **Standardize Formats, Style & Export:**
   - Convert Excel date serials (Windows `1900` & Mac `1904` epochs, overridable via `--epoch`) and date strings strictly to native `YYYY-MM-DD` (`yyyy-mm-dd`).
   - Normalize mixed US/EU currency formats (`1.200,50 €`, `$1,200.50`, `2.5M`), fold rescued hyperlink/comment notes, and quarantine fake placeholder contacts (`0000000`, `1234567`, `test@test.com`) into `_needs_review.csv` (`fake_placeholder_contact`).
   - Use adaptive streaming (`Workbook(write_only=True)` when rows `> 50,000`), Text (`@`) format on phone/ID columns, `utf-8-sig` BOM for `.csv`, and conditional destination saving per Stage 0 choices.
6. **Requested Export (`.xlsx`, `.csv`, or Both):**
   - Save cleanly to the user-selected destination path.
   - Styled `.xlsx` (auto-fitted columns, wrapped headers, frozen top row, filter enabled, conditional highlights if specified).
   - `.csv` encoded strictly as `utf-8-sig` (UTF-8 with BOM) for seamless Arabic/UTF-8 display in Microsoft Excel.
   - Optional `_needs_review.csv` quarantine log for dropped/unverified rows (if requested in Stage 0).

### Stage 4 — 100% Phone-by-Phone Raw Source Cross-Verification
Run automated verification comparing **every phone number** in every output file back to its raw source record(s):
- Verify Phone, WhatsApp, Name, Email, and Country/City fidelity.
- Assert **0 remaining defects**: no non-phone text in phone/WhatsApp columns, no invalid formats in email, no timestamps in date-only fields, no leading/trailing spaces or uncollapsed whitespace (`TRIM` verified), no Excel formula errors (`#REF!`, `#ERROR!`, `#N/A`), and no repeated `|` segments.

### Stage 5 — Final Concise Numeric Summary Report (Mandatory Ending)
Always conclude with a compact, factual numeric summary derived strictly from verified script & verification outputs:
1. **Row Funnel:** Raw rows scanned -> Shifted/header rows recovered -> Duplicates merged -> Dropped unreachable rows -> Final output row counts per file.
2. **Repairs Counter:** Shifted/multi-schema rows realigned, phones/WhatsApps cleaned, misplaced emails/campaigns/countries/times/budgets relocated, repeated `|` cells cleaned, and total cells normalized with strict `TRIM`.
3. **Verification Proof:** Total phone numbers matched (`X / Y`) back to raw source files + `0` remaining formatting defects confirmed.
4. **Output Deliverables:** Exact clickable file links in the chosen format(s) and destination folder.

---

## Directory Structure

- `SKILL.md` — Primary cross-agent skill definition and execution contract.
- `AGENTS.md` — Universal agent instructions for tools that auto-load `AGENTS.md`.
- `references/patterns-and-rules.md` — Detailed pattern catalog and normalization rules.
- `scripts/` — Reusable, standalone Python utilities (`pandas` + `openpyxl`).
- `examples/` — Anonymized synthetic sample CSVs (zero real PII).

## Privacy Rule for Public / Team Sharing

Never commit real client Excel/CSV files, audit JSON dumps containing phone numbers/emails, or customer PII to version control. Use `.gitignore` rules to track only skill instructions, generic scripts, references, and synthetic anonymized examples.

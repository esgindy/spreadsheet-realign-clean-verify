# Spreadsheet Realign, Clean & Verify (`spreadsheet-realign-clean-verify`)

> **Two-agent semantic cleaning, mid-sheet schema & merged-cell recovery, big-data streaming (50k–1M+ rows), universal strict `TRIM`, CSV safety & error guards, interactive setup, and 100% phone-by-phone source cross-verification for Excel & CSV.**  
> **Created by:** Eslam Elgindy | **License:** MIT

[![Standard: Agent Skills](https://img.shields.io/badge/Standard-Agent%20Skills%20(Open)-blue.svg)](#-cross-platform-installation)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-green.svg)](#requirements--quick-start)
[![Scale: 50k-1M+ Rows](https://img.shields.io/badge/Scale-Big%20Data%20Streaming-purple.svg)](#-key-capabilities-14-real-world--big-data-problems-solved)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)

> **GitHub About Description (Copy-Paste Ready):**  
> Cross-platform AI Agent Skill for cleaning, semantically realigning, strictly TRIMming, deduplicating, and 100% source-verifying messy & large-scale Excel/CSV spreadsheets (`50k–1M+` rows) with interactive setup, `_needs_review.csv` quarantine, Arabic/English bilingual intelligence, and UTF-8 BOM exports.

---

## Why Naive Cleaning Fails (المشكلة التي تحلها المهارة)

When real-world CRM, Facebook Lead Ads, WhatsApp, and property inventory exports come from dozens of Excel/CSV sheets, traditional single-pass column mappers fail because:

- **Columns shift mid-sheet** (`Rows 1–13` follow one layout, while `Rows 14+` shift without a new header).
- **Stacked sub-schemas** hide budgets, submission dates, or emails inside legacy columns (`LANG`, `REGION`, `when`, `PHONE`).
- **Header-as-data exports** put the first customer directly inside Excel row 1 headers.
- **Free text leaks into strict columns** (customer questions inside `WhatsApp`, emails inside `Preferred Time`, campaign tags holding audience languages, `"نعم - City"` prefixes).
- **Repeated pipes & invisible characters** (`|`-collapsed multi-unit blobs, `\xa0`, Unicode bidi marks `\u202d`/`\u200f`, extra internal spaces) break deduplication and CRM imports.
- **Plain UTF-8 CSVs scramble Arabic** in Microsoft Excel unless saved with UTF-8 BOM (`utf-8-sig`).

---

## Solution Architecture (منهجية الوكيلين والـ TRIM الشامل)

```text
  [0. Pre-Flight Interview]
      └─ Confirm output format (Excel / CSV UTF-8 BOM / Both),
         destination folder, & quarantine review preference
           │
           ▼
  [1. Adaptive Profiler & Raw Indexer]
      ├─ Streams large sheets (openpyxl read_only=True / chunksize)
      ├─ Fixes merged-cell ffill gaps, phantom empty rows/200+ cols
      ├─ Detects mid-sheet schema shifts & header-as-data rows
      └─ Indexes every phone/email -> (file, sheet, excel_row)
           │
           ▼
  [2. Agent 1 — Cell Semantic Reader]
      ├─ Reads cell content + raw source context semantically (AR/EN)
      ├─ Reclassifies Phone/WA/Email/Campaign/Country/Time/Budget/Name
      ├─ Flags email domain typos & strips subtotals/section dividers
      └─ Emits deterministic JSON decisions (agent1_semantic_decisions.json)
           │
           ▼
  [3. Agent 2 — Row Realigner, Split, TRIM & Exporter]
      ├─ Rebuilds rows cleanly & preserves 1-row-per-unit for asset owners
      ├─ Runs Arabic compare-key normalization for deduplication
      ├─ Applies strict_trim + Excel error scrub + injection guard on EVERY cell
      ├─ Converts dates/Excel serials (YYYY-MM-DD) & enforces Text '@' on phones
      └─ Exports .xlsx + .csv (utf-8-sig) + optional _needs_review.csv
           │
           ▼
  [4. Cross-Verification Audit & Stage 5 Report]
      └─ Matches EVERY output phone back to raw sources + emits compact numeric summary
```

## Key Capabilities (14 Real-World & Big-Data Problems Solved)

- **Stage 0 Interactive Setup & Stage 5 Audit Summary:** Never guesses where or how to save files, and always finishes with a compact numeric funnel (`Scanned -> Recovered -> Merged -> Quarantined -> Final`), repairs counter, and verified match proof.
- **Big-Data Ready (50k–1M+ Rows):** Uses adaptive streaming (`openpyxl read_only=True` / `write_only=True`, chunked CSV reading, and pre-stripping phantom `""` rows & 200+ blank columns) instead of loading giant Excel DOM trees into RAM.
- **Data Integrity Guards:** Fixes merged cells (`ffill`), prevents `9.715E+11` scientific notation & stripped leading zeros, neutralizes CSV formula injection (`=`, `+`, `-`, `@`), scrubs `#REF!`/`#ERROR!`, and quarantines rejected rows in `_needs_review.csv`.

---

## Cross-Platform Installation (التوافق مع كافة برامج الذكاء الاصطناعي)

Built on the **Open Agent Skills standard** (`SKILL.md` + `AGENTS.md` + standalone Python scripts) with **zero IDE lock-in**:

| AI Tool / Platform | Where to place this folder / files |
| :--- | :--- |
| **Antigravity** | `.agents/skills/spreadsheet-realign-clean-verify/` (workspace) or `~/.gemini/antigravity/skills/spreadsheet-realign-clean-verify/` |
| **Claude Code** | `.claude/skills/spreadsheet-realign-clean-verify/` or `~/.claude/skills/spreadsheet-realign-clean-verify/` |
| **OpenAI Codex CLI** | `.agents/skills/spreadsheet-realign-clean-verify/` (reads `SKILL.md` & `AGENTS.md` automatically) |
| **Cursor** | `.cursor/skills/spreadsheet-realign-clean-verify/` or `.cursor/rules/` referencing `SKILL.md` |
| **Gemini CLI** | `.gemini/skills/spreadsheet-realign-clean-verify/` |
| **Windsurf / Copilot / Others** | Copy `AGENTS.md` + `SKILL.md` + `scripts/` into project root or `.github/skills/spreadsheet-realign-clean-verify/` |

---

## Repository Structure

```text
spreadsheet-realign-clean-verify/
├── SKILL.md                          # Core open-standard AI skill instructions
├── AGENTS.md                         # Universal operating contract for coding agents
├── README.md                         # Bilingual documentation & installation guide
├── LICENSE                           # MIT License (Eslam Elgindy)
├── .gitignore                        # Blocks accidental upload of real customer PII
├── references/
│   └── patterns-and-rules.md         # Failure taxonomy & strict TRIM specification
├── scripts/
│   ├── profile_and_extract.py        # Schema-shift detector & raw source indexer
│   ├── apply_clean_trim_export.py    # Deterministic realigner, TRIM & XLSX/CSV exporter
│   └── cross_verify_sources.py       # 100% phone-by-phone raw-to-output verifier
└── examples/
    ├── dirty_leads_sample.csv        # Anonymized synthetic dirty sample
    └── expected_clean_sample.csv     # Anonymized synthetic cleaned output
```

---

## Requirements & Quick Start

- Python **3.10+**
- `pandas`, `openpyxl`, and `lxml` *(optional C-level XML accelerator for fast big-file reads)*

```bash
pip install pandas openpyxl lxml
```

```bash
# 1. Profile, stream-ingest large sheets, and build raw source verification index
python scripts/profile_and_extract.py -i ./raw_folder -w ./_clean_work

# 2. (Optional interactive Stage 0 prompt) Realign, strict TRIM, deduplicate & export
python scripts/apply_clean_trim_export.py -b ./_clean_work/extracted_base.json --interactive

# 3. Verify 100% of phones back to raw sources + emit Stage 5 audit JSON
python scripts/cross_verify_sources.py -c ./cleaned_output/cleaned_dataset.csv -r ./_clean_work/raw_source_index.json -m ./cleaned_output/cleaned_dataset_clean_metrics.json
```

---

## Privacy & Security

Never commit production spreadsheets (`.xlsx`, `.xls`, `.csv`) or real phone/email JSON dumps to GitHub. This repository's `.gitignore` blocks all real data files and keeps only synthetic anonymized CSV examples under `examples/`.

# Common Spreadsheet Corruption Patterns, Big-Data Scaling & Cleaning Rules

This reference documents **all 21 real-world corruption, schema, encoding, and large-file bottlenecks** solved by `spreadsheet-realign-clean-verify`.

---

## 1. Structural & Schema Failure Modes

| # | Pattern | Symptom in Raw File | Automated + Semantic Resolution |
| :- | :--- | :--- | :--- |
| 1 | **Mid-Sheet Schema Shift** | Rows `1–13` map cleanly, then rows `14+` shift left/right (`PHONE` holds timestamps, `DATE` holds phone numbers, `EMAIL` holds Ad Name). | Detect per-row content signatures (`@`, date pattern, digit length) and realign rows dynamically before deduplication. |
| 2 | **Stacked Form Sub-Schemas** | Multiple ad forms share one sheet under static headers (`REGION` holds Budget in rows 206–214 and Submission Date in rows 215+). | Classify block/row type by cell pattern (`@` inside `PHONE`, ISO date inside `REGION`) rather than column title alone. |
| 3 | **Header-Row-as-Data** | Exporting without column names promotes client #1 into row-1 headers (appending `.1` on colliding numbers). | Scan header cells for `@`, phone runs, or dates; recover row 0 into records and strip `.1` suffix artifacts. |
| 4 | **Merged Cells & `ffill` Gaps** | Visual grouping merges `A2:A5`, leaving `A3:A5` blank (`NaN`), breaking sorting/filtering/`GROUP BY`. | Detect merged ranges via `openpyxl`, unmerge cleanly, and forward-fill (`ffill`) category columns when structured. |
| 5 | **Phantom Blank Rows & 200+ Empty Columns** | Ghost formatting or `""` formulas create 3,500+ empty rows and 200+ duplicate empty columns, blowing up RAM and file size. | Strip completely empty trailing columns and empty rows during streaming ingestion before building DataFrames. |
| 6 | **Embedded Subtotals & Section Dividers** | `Total`, `Grand Total`, `Subtotal`, `Studio`, `Tower ( A )` mixed directly inside data rows. | Separate or strip divider/summary rows via semantic name/phone pattern gates so they never pollute clean entities. |
| 7 | **Entity Collapse (`\|` Concatenation)** | Owners with multiple properties/orders are collapsed by phone into `Apt 101 \| Apt 407 \| ...`. | Preserve 1 row per unique asset/unit key (`Developer + Building + Unit + PhoneTail`) for asset inventories while merging prospect leads cleanly. |

---

## 2. Cell-Level, Bilingual & Encoding Pitfalls

| # | Pattern | Symptom in Raw Cell | Resolution Rule |
| :- | :--- | :--- | :--- |
| 8 | **Free Text in Structured Fields** | Questions in `WhatsApp`, emails/budgets in `Preferred Time`, campaign tags holding audience labels (`ARABIC`, `GCC`), `"نعم - City"` prefixes. | Agent 1 reads cell meaning semantically: relocate values to true columns, preserve notes, and clear keyboard junk. |
| 9 | **Composite, Bidi & Arabic-Indic Phones** | `nullI971\|50-1234567InullI...`, `p:+`, Arabic-Indic `٠٥٦...`, invisible bidi marks (`\u202d`, `\u200f`), multiple numbers joined by `/` or `;`. | Strip bidi marks, translate localized digits to ASCII, split delimiters carefully without corrupting non-local country codes (`+44`, `+1`), extract Primary/WhatsApp/Additional phones. |
| 10 | **Date Hell & Excel Serial Integers** | `DD/MM` vs `MM/DD` clashes, ISO timestamps attached to campaign names, or raw Excel serials (`45292`). | Convert Excel serials (1900 epoch), strip time suffixes where date-only, standardize to native `date` with `YYYY-MM-DD` (`yyyy-mm-dd`). |
| 11 | **Scientific Notation & Excel Errors** | Phones saved as `9.715E+11`, dropped leading `0`, or literal `#REF!`, `#ERROR!`, `#N/A`, `#VALUE!`. | Force `dtype=str` on ingestion, scrub Excel error strings to blank (logged), format phone/ID columns in Excel strictly as Text (`@`). |
| 12 | **Email Typos & Formula Injection** | Domains like `gamil.com`, `icloude.com`, `gmail.co35518m`, OR CSV formula triggers (`=`, `+`, `-`, `@`). | Flag suspect email domains in an audit note rather than mutating silently; neutralize unsafe leading formula characters on CSV export while preserving valid phones/emails. |
| 13 | **Arabic Key Normalization (Compare Only)** | `أبو ظبي` vs `ابوظبي`, `القاهرة` vs `القاهره`, `إمارات` vs `امارات` fail exact string matches. | Map `أ/إ/آ -> ا`, `ى -> ي`, `ة -> ه` on deduplication/grouping keys **only**; retain proper spelling in output cells. |
| 14 | **Mojibake, BOM & Invisible Whitespace** | CSV opened in Excel scrambles Arabic without UTF-8 BOM; `\xa0`, `\t`, double spaces break joins. | Multi-encoding ingest fallback (`utf-8-sig` -> `utf-8` -> `cp1256` -> `latin1`), run `strict_trim` on **100% of output cells**, and export `.csv` with `utf-8-sig`. |
| 15 | **Hidden Rows & Hidden Sheets** | Filtered-out, deleted, or hidden scratch tabs silently sneak back into cleaned files. | Skip `ws.sheet_state != "visible"` and `ws.row_dimensions[r].hidden` by default (`--include-hidden-rows` / `--include-hidden-sheets` overrides). |
| 16 | **Hyperlink Contacts & Cell Comments** | Display text says `"Chat"` or `"Email"` while real `wa.me/`, `tel:`, `mailto:` link or agent notes sit in hyperlinks/comments. | Extract `cell.hyperlink.target`, `=HYPERLINK(...)` formulas, and `cell.comment.text`, fold into clean fields, and index in `raw_source_index.json`. |
| 17 | **Macintosh 1904 Date Epoch** | Old Mac Excel / legacy exports shift every serial date backward by 1,462 days (4 years + 1 day). | Detect `wb.epoch` (and `xlrd datemode`) automatically (`--epoch auto|1900|1904`) and convert with the exact epoch base (`1904-01-01` vs `1899-12-30`). |
| 18 | **Fake & Sequential Placeholders** | Lead forms contain `000000000`, `11111111`, `123456789`, `987654321`, or `test@test.com` bypassing regex checks. | Detect repetitive/sequential phones and dummy email pairs and route unreachable records to `_needs_review.csv` (`fake_placeholder_contact`). |
| 19 | **Mixed Locale Currencies & K/M** | Workbooks mix `1.200,50 €`, `$1,200.50`, `١٬٥٠٠٬٠٠٠`, `2.5M AED`, and `750k`. | Normalize US/EU separators and Arabic digits cleanly (`1200.50 EUR`, `2500000 AED`) while preserving multi-number text ranges (`1.5M - 2M`). |
| 20 | **Stacked Headers & Title Banners** | Report title rows sit above headers, or 2-row merged category/sub-column headers (`header=[0,1]`) break mapping. | Skip sparse banner rows, forward-fill/flatten 2-tier stacked headers cleanly, and prioritize leaf column hints when inferring canonical slots. |
| 21 | **Non-Comma CSVs, Multi-Line & `.xls`/`.xlsb`** | European `;` or `\t` CSVs, `Alt+Enter` multiline headers/cells, and legacy binary `.xls`/`.xlsb` break standard parsers. | Sniff delimiter across `16 KB` (`,`, `;`, `\t`, `|`) preserving multi-line quoted cells, and read `.xls`/`.xlsb` via `xlrd`/`pyxlsb` with actionable hints. |

---

## 3. Large Spreadsheet Performance & Memory Scaling (50k–1M+ Rows)

Standard `pd.read_excel()` or default `openpyxl.load_workbook()` allocates full XML DOM + style objects in RAM (~20x–50x file size), causing slowdowns or Out-Of-Memory (OOM) crashes on large workbooks.

Use **Adaptive Streaming** automatically when raw inputs exceed ~25,000 rows or ~10 MB:

1. **Streaming Excel Reads:** `openpyxl.load_workbook(path, read_only=True, data_only=True)` + `ws.reset_dimensions()` + `ws.iter_rows(values_only=True)` (resetting stale XML dimension bounds prevents missing trailing rows while skipping phantom blank rows immediately).
2. **Pre-Conversion for Giant `.xlsx`:** Use `xlsx2csv` (streaming XML-to-CSV converter) when workbooks exceed ~100k rows, then iterate CSV in memory-bounded batches.
3. **Chunked CSV Processing:** `pd.read_csv(path, chunksize=50_000, dtype=str, keep_default_na=False)` with compact hash indexes (`tail9` phone & lowercase email) rather than caching entire raw DOM trees in RAM.
4. **Streaming Excel Writes:** `openpyxl.Workbook(write_only=True)` when output rows exceed ~50,000 rows (avoiding heavy per-cell style objects across massive sheets, applying header styles cleanly via `WriteOnlyCell`).
5. **Fast XML Acceleration:** Ensure `lxml` is available so `openpyxl` runs at peak C-level parsing speed.

---

## 4. Mandatory `strict_trim` Specification

Execute on **every** non-empty cell after semantic realignment and summarization:

```python
import re

INVISIBLE_CHARS_RE = re.compile(r"[\u200b-\u200f\u202a-\u202e\ufeff\xa0]")
EXCEL_ERR_RE = re.compile(r"^#(REF!|DIV/0!|VALUE!|NAME\?|N/A|NUM!|NULL!|ERROR!)$", re.I)

def strict_trim(val: str) -> str:
    if val is None:
        return ""
    s = INVISIBLE_CHARS_RE.sub(" ", str(val))
    s = s.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    s = re.sub(r"\s+", " ", s).strip()
    if s.lower() in ("nan", "none", "null") or EXCEL_ERR_RE.match(s):
        return ""
    return s
```

---

## 5. Quarantine Policy (`_needs_review.csv`)

Never discard questionable rows silently. When enabled (Stage 0 default), write every excluded row (e.g., lacking both valid >=7-digit phone and valid email, subtotal footers, or ambiguous corrupted lines) to `_needs_review.csv` (`utf-8-sig`) with an explicit `exclusion_reason` column so reviewers can audit 100% of dropped rows in seconds.

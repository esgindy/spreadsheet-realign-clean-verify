#!/usr/bin/env python3
"""
Adaptive Streaming Profiler, Structural Fixer & Raw Indexer
Part of `spreadsheet-realign-clean-verify` by Eslam Elgindy (MIT License).

Solves:
- Merged cell ffill recovery, multi-row stacked headers, & title banner skipping
- Hidden rows & hidden sheets filtering (`ws.row_dimensions[r].hidden`, `sheet_state`)
- Hyperlink contact rescue (`wa.me/`, `tel:`, `mailto:`, `=HYPERLINK`) & cell comments
- Macintosh 1904 date epoch detection (`wb.epoch`) alongside Windows 1900 epoch
- Header-row-as-data recovery & phantom empty rows / 200+ blank columns cleanup
- Resilient CSV delimiter sniffing (`,`, `;`, `\t`, `|`) + multi-line quoted cells + `.xls`/`.xlsb` checks
- Memory-bounded streaming (`read_only=True` / `chunksize`) & 100% raw phone/email indexing
"""

from __future__ import annotations

import argparse
import csv
import importlib.util
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import openpyxl

INVISIBLE_RE = re.compile(r"[\u200b-\u200f\u202a-\u202e\ufeff\xa0]")
EXCEL_ERR_RE = re.compile(r"^#(REF!|DIV/0!|VALUE!|NAME\?|N/A|NUM!|NULL!|ERROR!)$", re.I)
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,24}")
SCI_NUM_RE = re.compile(r"^\d+(\.\d+)?[eE][+\-]?\d+$")
HYPERLINK_FORMULA_RE = re.compile(r'=HYPERLINK\(\s*"([^"]+)"\s*[,;]', re.I)
WA_LINK_RE = re.compile(r"(?:wa\.me/|api\.whatsapp\.com/send\?phone=)\+?(\d{7,15})", re.I)
TEL_LINK_RE = re.compile(r"tel:\+?(\d{7,15})", re.I)
MAILTO_LINK_RE = re.compile(r"mailto:([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,24})", re.I)
FAKE_PLACEHOLDER_LOCAL = {"test", "fake", "dummy", "none", "noemail", "na", "null", "noreply", "aaa", "asdf"}
DATE_LIKE_RE = re.compile(
    r"^(\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4})"
    r"([T\s]\d{1,2}:\d{2}(:\d{2}(\.\d+)?)?Z?)?$"
)
SUBTOTAL_DIVIDER_RE = re.compile(
    r"^(total|grand\s*total|subtotal|sub-total|duplicae|duplicate|"
    r"studio|ssstudio|\d+-bedroom(\+study)?|tower\s*\(?.*\)?)$",
    re.I,
)
SUSPECT_EMAIL_DOMAINS = {
    "gamil.com", "gimil.com", "glail.com", "gmul.com", "gamli.com", "gamail.com",
    "gmal.com", "gmaill.com", "gmail.co", "gmail.con", "gmail.coma", "gmail.come",
    "gmail.comm", "gmail.coms", "gmail.comj", "gmail.com0", "gmeal.come",
    "hotmil.com", "hotmal.come", "hotmaol.com", "hotmail.con", "hotmail.coms",
    "hotmail.comi", "icloude.com", "icloud.con", "yahoo.vom", "yaho.com",
}

TRANS_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")

CANONICAL_HEADER_HINTS: Dict[str, Tuple[str, ...]] = {
    "project": ("project name", "project", "rooms", "room", "bedroom", "property", "building", "unit", "المشروع", "الوحدة", "البرج", "الشقة"),
    "campaign": ("campaign name", "campaign", "source", "ad name", "form", "حملة", "الحملة", "المصدر"),
    "whatsapp": ("whatsapp", "wtsapp", "wa", "واتساب", "واتس", "الواتس"),
    "extra_phone": ("extra phone", "secondary phone", "alt phone", "هاتف اضافي", "هاتف إضافي", "رقم اضافي"),
    "phone": ("phone", "mobile", "tel", "contact", "هاتف", "رقم الهاتف", "الجوال", "موبايل"),
    "email": ("email", "e-mail", "mail", "البريد", "الايميل", "إيميل", "بريد"),
    "budget": ("budget", "price", "amount", "cost", "السعر", "الميزانية", "القيمة"),
    "feedback": ("feedback", "feed back", "status", "result", "الحالة", "الفيدباك"),
    "notes": ("note", "notes", "comment", "comments", "remark", "ملاحظات", "تعليق", "تعليقات"),
    "pref_time": ("time", "when", "prefer", "contact time", "الوقت", "وقت التواصل"),
    "country": ("country", "city", "region", "location", "nationality", "الدولة", "المدينة", "الجنسية"),
    "date": ("date", "created", "submitted", "timestamp", "تاريخ", "تاريخ التسجيل"),
    "name": ("full name", "client name", "customer name", "name", "client", "lead", "customer", "الاسم", "اسم العميل", "اسم المالك"),
}


def basic_trim(val: Any) -> str:
    if val is None:
        return ""
    s = INVISIBLE_RE.sub(" ", str(val))
    s = s.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return "" if s.lower() in ("nan", "none", "null") else s


def digits_only(val: str) -> str:
    return re.sub(r"\D", "", val.translate(TRANS_DIGITS))


def is_phone_like(val: str) -> bool:
    if not val or EMAIL_RE.search(val) or DATE_LIKE_RE.match(val):
        return False
    d = digits_only(val)
    return 7 <= len(d) <= 15 and sum(c.isalpha() for c in val) <= 3


def is_fake_placeholder_phone(val: str) -> bool:
    d = digits_only(val)
    if len(d) < 7:
        return False
    if len(set(d)) == 1 or len(set(d[-7:])) == 1:
        return True
    if d in "0123456789012345" or d in "9876543210":
        return True
    return len(d) > 7 and d[-7:] in ("0000000", "1111111", "1234567", "7654321", "9999999")


def is_fake_placeholder_email(val: str) -> bool:
    m = EMAIL_RE.search(val or "")
    if not m:
        return False
    em = m.group(0).lower()
    local, _, domain = em.partition("@")
    dom_stem = domain.split(".")[0]
    return local in FAKE_PLACEHOLDER_LOCAL and (local == dom_stem or domain in ("test.com", "noemail.com", "none.com"))


def extract_contact_from_hyperlink(target: str) -> Optional[Tuple[str, str]]:
    """Extract recoverable contact (whatsapp/phone/email) from URL or =HYPERLINK(...) formula."""
    if not target:
        return None
    m_formula = HYPERLINK_FORMULA_RE.search(target)
    url = m_formula.group(1) if m_formula else target
    m_wa = WA_LINK_RE.search(url)
    if m_wa:
        return "whatsapp", m_wa.group(1)
    m_tel = TEL_LINK_RE.search(url)
    if m_tel:
        return "phone", m_tel.group(1)
    m_mail = MAILTO_LINK_RE.search(url)
    if m_mail:
        return "email", m_mail.group(1).lower()
    return None


def looks_like_header_as_data(headers: List[str]) -> bool:
    for h in headers:
        s = basic_trim(h)
        if not s:
            continue
        if EMAIL_RE.search(s) or is_phone_like(s) or DATE_LIKE_RE.match(s):
            return True
    return False


def _match_header_hint(header_text: str) -> Optional[str]:
    hl = re.sub(r"\.\d+$", "", basic_trim(header_text).lower())
    candidates = [hl.split(" - ", 1)[-1].strip(), hl] if " - " in hl else [hl]
    for cand in candidates:
        if not cand:
            continue
        for canon, hints in CANONICAL_HEADER_HINTS.items():
            if any(hint == cand or hint in cand for hint in hints):
                return canon
    return None


def _count_header_hints(cells: List[str]) -> int:
    return sum(1 for c in cells if c and _match_header_hint(c) is not None)


def resolve_headers_and_rows(
    indexed_entries: List[Tuple[int, List[str], List[Tuple[str, str]], List[str]]],
) -> Tuple[List[str], List[Tuple[int, List[str], List[Tuple[str, str]], List[str]]], bool, bool, int]:
    """Trim phantom trailing cols, skip title banners, flatten 2-row stacked headers, and detect header_as_data."""
    if not indexed_entries:
        return [], [], False, False, 0

    max_used = 0
    for _, r, _, _ in indexed_entries:
        for idx in range(len(r) - 1, -1, -1):
            if r[idx] != "":
                max_used = max(max_used, idx + 1)
                break
    if max_used == 0:
        return [], [], False, False, 0

    entries = [
        (erow, r[:max_used] + [""] * max(0, max_used - len(r)), links, comments)
        for erow, r, links, comments in indexed_entries
    ]

    banner_skipped = 0
    # Skip single-cell / sparse banner title above real header row
    if len(entries) >= 2:
        non_empty_0 = sum(1 for c in entries[0][1] if c)
        non_empty_1 = sum(1 for c in entries[1][1] if c)
        if non_empty_0 <= 2 and non_empty_1 >= 3 and not looks_like_header_as_data(entries[1][1]) and _count_header_hints(entries[1][1]) >= 2:
            entries = entries[1:]
            banner_skipped = 1

    raw_first = entries[0][1]
    header_as_data = looks_like_header_as_data(raw_first)
    if header_as_data:
        headers = [f"col_{i}" for i in range(max_used)]
        return headers, entries, True, False, banner_skipped

    # Detect two-row stacked / grouped headers (`header=[0, 1]`)
    stacked_used = False
    if len(entries) >= 2:
        r0 = entries[0][1]
        r1 = entries[1][1]
        if not looks_like_header_as_data(r1) and _count_header_hints(r1) >= 2:
            r0_non_empty = [c for c in r0 if c]
            r0_has_group_repeat = len(r0_non_empty) > len(set(r0_non_empty))
            r0_sparse_groups = 1 <= len(r0_non_empty) < sum(1 for c in r1 if c)
            r1_non_empty = [c for c in r1 if c]
            r0_has_diff = any(a != b and a and b for a, b in zip(r0, r1))
            if len(r1_non_empty) >= 2 and (r0_has_group_repeat or r0_sparse_groups) and r0_has_diff:
                # Forward-fill sparse group headers when CSV lacks merged cells
                ffill_r0: List[str] = []
                last_g = ""
                for g in r0:
                    if g:
                        last_g = g
                    ffill_r0.append(g or last_g)
                flattened: List[str] = []
                for g, sub in zip(ffill_r0 if r0_sparse_groups and not r0_has_group_repeat else r0, r1):
                    if g and sub and g.lower() != sub.lower():
                        flattened.append(f"{g} - {sub}")
                    else:
                        flattened.append(sub or g)
                return flattened, entries[2:], False, True, banner_skipped

    return raw_first, entries[1:], False, stacked_used, banner_skipped


def detect_csv_delimiter(sample: str) -> str:
    if not sample:
        return ","
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        return dialect.delimiter
    except csv.Error:
        header_line = sample.splitlines()[0] if sample.splitlines() else sample
        counts = {d: header_line.count(d) for d in (",", ";", "\t", "|")}
        best = max(counts, key=counts.get)
        return best if counts[best] > 0 else ","


def read_csv_rows(path: Path, chunk_size: int = 50_000, meta_out: Optional[Dict[str, Any]] = None) -> Iterable[List[str]]:
    for enc in ("utf-8-sig", "utf-8", "cp1256", "latin1"):
        try:
            with path.open("r", encoding=enc, newline="") as fh:
                sample = fh.read(16_384)
                fh.seek(0)
                delim = detect_csv_delimiter(sample)
                if meta_out is not None:
                    meta_out["delimiter"] = "\\t" if delim == "\t" else delim
                    meta_out["encoding"] = enc
                reader = csv.reader(fh, delimiter=delim)
                batch: List[List[str]] = []
                for row in reader:
                    cleaned = [basic_trim(c) for c in row]
                    if any(cleaned):
                        batch.append(cleaned)
                    if len(batch) >= chunk_size:
                        for item in batch:
                            yield item
                        batch.clear()
                for item in batch:
                    yield item
            return
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("utf-8", b"", 0, 1, f"Unable to decode CSV file: {path}")


def unmerge_and_ffill_xlsx(
    path: Path,
    include_hidden_rows: bool = False,
    include_hidden_sheets: bool = False,
) -> Tuple[Dict[str, List[Tuple[int, List[str], List[Tuple[str, str]], List[str]]]], Dict[str, Any]]:
    """Full load for rich metadata: merged cells, hidden rows/sheets, hyperlinks, comments & 1904 epoch."""
    wb = openpyxl.load_workbook(path, data_only=True)
    epoch = 1904 if (getattr(wb, "epoch", None) and wb.epoch.year == 1904) else 1900
    hidden_sheets_skipped = 0
    hidden_rows_skipped = 0
    hyperlinks_rescued = 0
    comments_rescued = 0
    result: Dict[str, List[Tuple[int, List[str], List[Tuple[str, str]], List[str]]]] = {}

    for name in wb.sheetnames:
        ws = wb[name]
        if not include_hidden_sheets and getattr(ws, "sheet_state", "visible") != "visible":
            hidden_sheets_skipped += 1
            continue
        merged_map: Dict[Tuple[int, int], Any] = {}
        for mrange in ws.merged_cells.ranges:
            top_val = ws.cell(mrange.min_row, mrange.min_col).value
            for r in range(mrange.min_row, mrange.max_row + 1):
                for c in range(mrange.min_col, mrange.max_col + 1):
                    merged_map[(r, c)] = top_val

        rows: List[Tuple[int, List[str], List[Tuple[str, str]], List[str]]] = []
        for r_idx in range(1, ws.max_row + 1):
            r_dim = ws.row_dimensions.get(r_idx)
            if not include_hidden_rows and r_dim is not None and bool(getattr(r_dim, "hidden", False)):
                hidden_rows_skipped += 1
                continue
            row_vals: List[str] = []
            row_links: List[Tuple[str, str]] = []
            row_comments: List[str] = []
            for c_idx in range(1, ws.max_column + 1):
                cell = ws.cell(r_idx, c_idx)
                raw_v = merged_map.get((r_idx, c_idx), cell.value)
                trimmed_v = basic_trim(raw_v)

                link_target = getattr(getattr(cell, "hyperlink", None), "target", None) or ""
                if not link_target and isinstance(raw_v, str) and "HYPERLINK" in raw_v.upper():
                    link_target = raw_v
                if link_target:
                    extracted = extract_contact_from_hyperlink(link_target)
                    if extracted and extracted[1] not in trimmed_v.lower():
                        row_links.append(extracted)
                        hyperlinks_rescued += 1

                if getattr(cell, "comment", None) and getattr(cell.comment, "text", None):
                    cmt = basic_trim(cell.comment.text)
                    if cmt:
                        row_comments.append(cmt)
                        comments_rescued += 1

                row_vals.append(trimmed_v)
            if any(row_vals) or row_links or row_comments:
                rows.append((r_idx, row_vals, row_links, row_comments))
        if rows:
            result[name] = rows
    wb.close()
    return result, {
        "epoch": epoch,
        "hidden_sheets_skipped": hidden_sheets_skipped,
        "hidden_rows_skipped": hidden_rows_skipped,
        "hyperlinks_rescued": hyperlinks_rescued,
        "comments_rescued": comments_rescued,
    }


def stream_xlsx_rows(
    path: Path,
    max_cols_guard: int = 250,
    include_hidden_sheets: bool = False,
    meta_out: Optional[Dict[str, Any]] = None,
) -> Iterable[Tuple[str, int, List[str]]]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    epoch = 1904 if (getattr(wb, "epoch", None) and wb.epoch.year == 1904) else 1900
    hidden_sheets_skipped = 0
    if meta_out is not None:
        meta_out["epoch"] = epoch
        meta_out["hidden_sheets_skipped"] = 0
    try:
        for name in wb.sheetnames:
            ws = wb[name]
            if not include_hidden_sheets and getattr(ws, "sheet_state", "visible") != "visible":
                hidden_sheets_skipped += 1
                if meta_out is not None:
                    meta_out["hidden_sheets_skipped"] = hidden_sheets_skipped
                continue
            ws.reset_dimensions()  # Avoid wrong XML dimension bounds clipping trailing rows
            for excel_row, row in enumerate(ws.iter_rows(values_only=True), start=1):
                if len(row) > max_cols_guard:
                    non_empty_last = 0
                    for idx in range(len(row) - 1, -1, -1):
                        if row[idx] is not None and str(row[idx]).strip() != "":
                            non_empty_last = idx + 1
                            break
                    row = row[:max(1, non_empty_last)]
                cleaned = [basic_trim(c) for c in row]
                if any(cleaned):
                    yield name, excel_row, cleaned
    finally:
        wb.close()


def read_legacy_excel_sheets(
    path: Path,
    include_hidden_sheets: bool = False,
) -> Tuple[Dict[str, List[Tuple[int, List[str], List[Tuple[str, str]], List[str]]]], int]:
    """Read legacy .xls / .xlsb when xlrd, pyxlsb, or python-calamine is available."""
    ext = path.suffix.lower()
    sheets: Dict[str, List[Tuple[int, List[str], List[Tuple[str, str]], List[str]]]] = {}
    epoch = 1900

    if ext == ".xls" and importlib.util.find_spec("xlrd"):
        import xlrd  # type: ignore

        book = xlrd.open_workbook(str(path))
        epoch = 1904 if getattr(book, "datemode", 0) == 1 else 1900
        for s in book.sheets():
            if not include_hidden_sheets and getattr(s, "visibility", 0) != 0:
                continue
            s_rows: List[Tuple[int, List[str], List[Tuple[str, str]], List[str]]] = []
            for r_idx in range(s.nrows):
                vals = [basic_trim(s.cell_value(r_idx, c_idx)) for c_idx in range(s.ncols)]
                if any(vals):
                    s_rows.append((r_idx + 1, vals, [], []))
            if s_rows:
                sheets[s.name] = s_rows
        return sheets, epoch

    if ext == ".xlsb" and importlib.util.find_spec("pyxlsb"):
        from pyxlsb import open_workbook  # type: ignore

        with open_workbook(str(path)) as book:
            for sname in book.sheets:
                with book.get_sheet(sname) as s:
                    s_rows = []
                    for r_idx, row in enumerate(s.rows(), start=1):
                        vals = [basic_trim(c.v) for c in row]
                        if any(vals):
                            s_rows.append((r_idx, vals, [], []))
                    if s_rows:
                        sheets[sname] = s_rows
        return sheets, epoch

    hint = "xlrd" if ext == ".xls" else "pyxlsb"
    raise ValueError(
        f"Legacy binary spreadsheet '{path.name}' ({ext}) requires `pip install {hint}` "
        "or conversion to `.xlsx` / `.csv` before ingestion."
    )


def infer_column_map(headers: List[str], sample_rows: List[List[str]]) -> Dict[int, str]:
    col_map: Dict[int, str] = {}
    for idx, h in enumerate(headers):
        matched = _match_header_hint(h)
        if not matched and sample_rows:
            vals = [r[idx] for r in sample_rows[:50] if idx < len(r) and r[idx]]
            if vals:
                emails = sum(1 for v in vals if EMAIL_RE.search(v))
                phones = sum(1 for v in vals if is_phone_like(v))
                dates = sum(1 for v in vals if DATE_LIKE_RE.match(v))
                if emails >= max(2, len(vals) * 0.4):
                    matched = "email"
                elif phones >= max(2, len(vals) * 0.4):
                    matched = "phone"
                elif dates >= max(2, len(vals) * 0.4):
                    matched = "date"
        col_map[idx] = matched or f"extra_{idx}_{basic_trim(h) or 'col'}"
    return col_map


def detect_row_anomalies(row_dict: Dict[str, str], raw_cells: List[str]) -> List[str]:
    flags: List[str] = []
    for key, val in row_dict.items():
        if not val:
            continue
        if EXCEL_ERR_RE.match(val):
            flags.append(f"EXCEL_ERR:{key}")
        if SCI_NUM_RE.match(val):
            flags.append(f"SCI_NOTATION:{key}")
        if "|" in val and len({basic_trim(p) for p in val.split("|") if basic_trim(p)}) > 0:
            if val.count("|") >= 1:
                flags.append(f"PIPE_MULTI:{key}")
        if key == "name" and (SUBTOTAL_DIVIDER_RE.match(val) or "?" in val or "؟" in val):
            flags.append("SUSPECT_NAME_OR_DIVIDER")
        if key in ("phone", "whatsapp"):
            if not is_phone_like(val):
                flags.append(f"NON_PHONE_IN_{key.upper()}")
            elif is_fake_placeholder_phone(val):
                flags.append(f"FAKE_PLACEHOLDER_PHONE:{key}")
        if key == "email":
            m = EMAIL_RE.search(val)
            if not m:
                flags.append("NON_EMAIL_IN_EMAIL")
            else:
                if is_fake_placeholder_email(val):
                    flags.append("FAKE_PLACEHOLDER_EMAIL")
                domain = m.group(0).split("@")[-1].lower()
                if domain in SUSPECT_EMAIL_DOMAINS:
                    flags.append(f"EMAIL_TYPO_SUSPECT:{domain}")
        if key not in ("email",) and EMAIL_RE.search(val):
            flags.append(f"SHIFTED_EMAIL_IN_{key}")
        if key not in ("phone", "whatsapp", "extra_phone", "budget") and is_phone_like(val):
            flags.append(f"SHIFTED_PHONE_IN_{key}")
        if key not in ("date", "campaign") and DATE_LIKE_RE.match(val):
            flags.append(f"SHIFTED_DATE_IN_{key}")

    if not flags:
        # Check raw cells for invisible directional chars if raw differs visually
        for rc in raw_cells:
            if INVISIBLE_RE.search(rc):
                flags.append("INVISIBLE_CHARS")
                break
    return sorted(set(flags))


def _extract_raw_text_from_cells(cells: List[str], row_links: List[Tuple[str, str]], row_comments: List[str]) -> str:
    parts = list(cells)
    parts.extend(v for _, v in row_links if v)
    parts.extend(c for c in row_comments if c)
    return " | ".join(p for p in parts if p)


def _index_raw_contacts(
    text: str,
    file_name: str,
    sheet_name: str,
    excel_row: int,
    raw_phone_index: Dict[str, List[Dict[str, Any]]],
    raw_email_index: Dict[str, List[Dict[str, Any]]],
) -> None:
    snippet = text[:280]
    for m in re.finditer(r"[\d٠-٩۰-۹][\d٠-٩۰-۹\s\-+().]{5,20}[\d٠-٩۰-۹]", text):
        d = digits_only(m.group(0))
        if 7 <= len(d) <= 15:
            tail = d[-9:] if len(d) >= 9 else d
            raw_phone_index.setdefault(tail, []).append({
                "file": file_name,
                "sheet": sheet_name,
                "row": excel_row,
                "digits": d,
                "snippet": snippet,
            })
    for em in EMAIL_RE.findall(text):
        raw_email_index.setdefault(em.lower(), []).append({
            "file": file_name,
            "sheet": sheet_name,
            "row": excel_row,
            "snippet": snippet,
        })


def ingest_files(
    inputs: List[Path],
    out_dir: Path,
    stream_mb_threshold: float = 10.0,
    unmerge_ffill: bool = True,
    epoch_override: str = "auto",
    include_hidden_rows: bool = False,
    include_hidden_sheets: bool = False,
) -> Dict[str, Any]:
    out_dir.mkdir(parents=True, exist_ok=True)
    records: List[Dict[str, Any]] = []
    audit_batch: List[Dict[str, Any]] = []
    raw_phone_index: Dict[str, List[Dict[str, Any]]] = {}
    raw_email_index: Dict[str, List[Dict[str, Any]]] = {}
    file_stats: List[Dict[str, Any]] = []

    for path in inputs:
        if not path.exists():
            raise FileNotFoundError(f"Input file not found: {path}")
        size_mb = path.stat().st_size / (1024 * 1024)
        ext = path.suffix.lower()

        sheet_bundles: List[Tuple[str, List[Tuple[int, List[str], List[Tuple[str, str]], List[str]]], Dict[str, Any]]] = []
        if ext == ".csv":
            csv_meta: Dict[str, Any] = {"epoch": 1900}
            rows = [
                (idx, r, [], [])
                for idx, r in enumerate(read_csv_rows(path, meta_out=csv_meta), start=1)
            ]
            mode = "csv_chunked_stream"
            sheet_bundles.append(("CSV", rows, csv_meta))
        elif ext in (".xlsx", ".xlsm"):
            if unmerge_ffill and size_mb < stream_mb_threshold:
                mode = "xlsx_unmerge_ffill"
                sheet_dict, wb_meta = unmerge_and_ffill_xlsx(
                    path,
                    include_hidden_rows=include_hidden_rows,
                    include_hidden_sheets=include_hidden_sheets,
                )
                for sname, srows in sheet_dict.items():
                    sheet_bundles.append((sname, srows, dict(wb_meta)))
            else:
                mode = "xlsx_read_only_stream"
                stream_meta: Dict[str, Any] = {"epoch": 1900, "hidden_sheets_skipped": 0}
                by_sheet: Dict[str, List[Tuple[int, List[str], List[Tuple[str, str]], List[str]]]] = {}
                for sname, excel_row, r in stream_xlsx_rows(
                    path, include_hidden_sheets=include_hidden_sheets, meta_out=stream_meta
                ):
                    by_sheet.setdefault(sname, []).append((excel_row, r, [], []))
                for sname, srows in by_sheet.items():
                    sheet_bundles.append((sname, srows, dict(stream_meta)))
        elif ext in (".xls", ".xlsb"):
            mode = f"legacy_{ext.lstrip('.')}"
            sheet_dict, leg_epoch = read_legacy_excel_sheets(path, include_hidden_sheets=include_hidden_sheets)
            for sname, srows in sheet_dict.items():
                sheet_bundles.append((sname, srows, {"epoch": leg_epoch}))
        else:
            raise ValueError(f"Unsupported extension '{ext}' for {path}. Supported: .xlsx, .xlsm, .csv, .xls, .xlsb.")

        for sname, indexed_entries, meta in sheet_bundles:
            if not indexed_entries:
                continue
            headers, data_entries, header_as_data, stacked_used, banner_skipped = resolve_headers_and_rows(indexed_entries)
            if not headers:
                continue

            eff_epoch = int(epoch_override) if epoch_override in ("1900", "1904") else int(meta.get("epoch", 1900))
            col_map = infer_column_map(headers, [r for _, r, _, _ in data_entries[:50]])
            sheet_count = 0
            csv_formula_links_rescued = 0

            for excel_row, cells, row_links, row_comments in data_entries:
                merged_links = list(row_links)
                for c_val in cells:
                    if "HYPERLINK" in c_val.upper():
                        ext_c = extract_contact_from_hyperlink(c_val)
                        if ext_c and ext_c not in merged_links:
                            merged_links.append(ext_c)
                            csv_formula_links_rescued += 1

                row_dict: Dict[str, str] = {}
                for c_idx, c_val in enumerate(cells):
                    canon = col_map.get(c_idx, f"col_{c_idx}")
                    if canon in row_dict and row_dict[canon] and c_val:
                        if c_val not in row_dict[canon]:
                            row_dict[canon] = f"{row_dict[canon]} | {c_val}"
                    elif c_val:
                        row_dict[canon] = c_val
                if not any(row_dict.values()) and not merged_links and not row_comments:
                    continue

                flags = detect_row_anomalies(row_dict, cells)
                rec: Dict[str, Any] = {
                    "record_id": len(records) + 1,
                    "source_file": path.name,
                    "source_sheet": sname,
                    "source_row": excel_row,
                    "header_as_data": header_as_data,
                    "date_epoch": eff_epoch,
                    "fields": row_dict,
                    "raw_cells": cells,
                    "flags": flags,
                }
                if merged_links:
                    rec["hyperlinks_rescued"] = [[slot, val] for slot, val in merged_links]
                if row_comments:
                    rec["cell_comments"] = row_comments

                records.append(rec)
                sheet_count += 1
                if flags:
                    audit_batch.append(rec)

                joined_raw = _extract_raw_text_from_cells(cells, merged_links, row_comments)
                _index_raw_contacts(joined_raw, path.name, sname, excel_row, raw_phone_index, raw_email_index)

            stat_item: Dict[str, Any] = {
                "file": path.name,
                "sheet": sname,
                "mode": mode,
                "size_mb": round(size_mb, 3),
                "rows_extracted": sheet_count,
                "header_as_data": header_as_data,
                "stacked_headers_flattened": stacked_used,
                "banner_rows_skipped": banner_skipped,
                "date_epoch": eff_epoch,
                "hidden_rows_skipped": meta.get("hidden_rows_skipped", 0),
                "hidden_sheets_skipped": meta.get("hidden_sheets_skipped", 0),
                "hyperlinks_rescued": meta.get("hyperlinks_rescued", 0) + csv_formula_links_rescued,
                "comments_rescued": meta.get("comments_rescued", 0),
            }
            if "delimiter" in meta:
                stat_item["csv_delimiter"] = meta["delimiter"]
            file_stats.append(stat_item)

    (out_dir / "extracted_base.json").write_text(
        json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (out_dir / "audit_batch.json").write_text(
        json.dumps(audit_batch, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (out_dir / "raw_source_index.json").write_text(
        json.dumps(
            {"phones": raw_phone_index, "emails": raw_email_index, "stats": file_stats},
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    return {
        "total_records": len(records),
        "flagged_rows": len(audit_batch),
        "indexed_phone_tails": len(raw_phone_index),
        "indexed_emails": len(raw_email_index),
        "file_stats": file_stats,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Profile, stream-ingest, and index messy Excel/CSV files.")
    parser.add_argument("--input", "-i", nargs="+", required=True, help="Input .xlsx/.xlsm/.csv/.xls/.xlsb file(s) or folders.")
    parser.add_argument("--work-dir", "-w", default="./_clean_work", help="Directory for extracted JSON artifacts.")
    parser.add_argument("--stream-mb", type=float, default=10.0, help="Size MB threshold to switch XLSX to read_only stream.")
    parser.add_argument("--no-unmerge", action="store_true", help="Always stream XLSX without merged-cell ffill pass.")
    parser.add_argument("--epoch", choices=("auto", "1900", "1904"), default="auto", help="Override Excel date serial epoch.")
    parser.add_argument("--include-hidden-rows", action="store_true", help="Keep hidden Excel rows instead of skipping them.")
    parser.add_argument("--include-hidden-sheets", action="store_true", help="Keep hidden Excel worksheets instead of skipping them.")
    args = parser.parse_args()

    files: List[Path] = []
    for item in args.input:
        p = Path(item)
        if p.is_dir():
            for ext in ("*.xlsx", "*.xlsm", "*.xls", "*.xlsb", "*.csv"):
                files.extend(sorted(p.rglob(ext)))
        else:
            files.append(p)

    summary = ingest_files(
        files,
        Path(args.work_dir),
        stream_mb_threshold=args.stream_mb,
        unmerge_ffill=not args.no_unmerge,
        epoch_override=args.epoch,
        include_hidden_rows=args.include_hidden_rows,
        include_hidden_sheets=args.include_hidden_sheets,
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())

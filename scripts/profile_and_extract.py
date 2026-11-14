#!/usr/bin/env python3
"""
Adaptive Streaming Profiler, Structural Fixer & Raw Indexer
Part of `spreadsheet-realign-clean-verify` by Eslam Elgindy (MIT License).

Solves:
- Merged cell ffill recovery & phantom empty rows / 200+ blank columns
- Header-row-as-data recovery (strips pandas `.1` suffixes)
- Mid-sheet schema shifts & stacked sub-schema anomaly flagging
- Multi-encoding CSV ingestion (utf-8-sig, utf-8, cp1256, latin1)
- Memory-bounded streaming (`read_only=True` / `chunksize`) for large files
- Raw phone & email indexing for 100% Stage 4 cross-verification
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import openpyxl

INVISIBLE_RE = re.compile(r"[\u200b-\u200f\u202a-\u202e\ufeff\xa0]")
EXCEL_ERR_RE = re.compile(r"^#(REF!|DIV/0!|VALUE!|NAME\?|N/A|NUM!|NULL!|ERROR!)$", re.I)
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,24}")
SCI_NUM_RE = re.compile(r"^\d+(\.\d+)?[eE][+\-]?\d+$")
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


def looks_like_header_as_data(headers: List[str]) -> bool:
    for h in headers:
        s = basic_trim(h)
        if not s:
            continue
        if EMAIL_RE.search(s) or is_phone_like(s) or DATE_LIKE_RE.match(s):
            return True
    return False


def strip_empty_trailing_cols(rows: List[List[str]]) -> Tuple[List[str], List[List[str]]]:
    if not rows:
        return [], []
    max_used = 0
    for r in rows:
        for idx in range(len(r) - 1, -1, -1):
            if r[idx] != "":
                max_used = max(max_used, idx + 1)
                break
    if max_used == 0:
        return [], []
    trimmed = [r[:max_used] + [""] * max(0, max_used - len(r)) for r in rows]
    return trimmed[0], trimmed[1:]


def read_csv_rows(path: Path, chunk_size: int = 50_000) -> Iterable[List[str]]:
    for enc in ("utf-8-sig", "utf-8", "cp1256", "latin1"):
        try:
            with path.open("r", encoding=enc, newline="") as fh:
                reader = csv.reader(fh)
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


def unmerge_and_ffill_xlsx(path: Path) -> Dict[str, List[List[str]]]:
    """Full load only for smaller files where merged-cell ffill is needed."""
    wb = openpyxl.load_workbook(path, data_only=True)
    result: Dict[str, List[List[str]]] = {}
    for name in wb.sheetnames:
        ws = wb[name]
        merged_map: Dict[Tuple[int, int], Any] = {}
        for mrange in ws.merged_cells.ranges:
            top_val = ws.cell(mrange.min_row, mrange.min_col).value
            for r in range(mrange.min_row, mrange.max_row + 1):
                for c in range(mrange.min_col, mrange.max_col + 1):
                    merged_map[(r, c)] = top_val
        rows: List[List[str]] = []
        for r_idx in range(1, ws.max_row + 1):
            row_vals: List[str] = []
            for c_idx in range(1, ws.max_column + 1):
                v = merged_map.get((r_idx, c_idx), ws.cell(r_idx, c_idx).value)
                row_vals.append(basic_trim(v))
            if any(row_vals):
                rows.append(row_vals)
        if rows:
            result[name] = rows
    wb.close()
    return result


def stream_xlsx_rows(path: Path, max_cols_guard: int = 250) -> Iterable[Tuple[str, int, List[str]]]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for name in wb.sheetnames:
            ws = wb[name]
            ws.reset_dimensions()  # Avoid wrong XML dimension bounds clipping trailing rows
            for excel_row, row in enumerate(ws.iter_rows(values_only=True), start=1):
                if len(row) > max_cols_guard:
                    # Trim excessive phantom trailing empty columns early
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


def infer_column_map(headers: List[str], sample_rows: List[List[str]]) -> Dict[int, str]:
    col_map: Dict[int, str] = {}
    for idx, h in enumerate(headers):
        hl = re.sub(r"\.\d+$", "", basic_trim(h).lower())
        matched: Optional[str] = None
        for canon, hints in CANONICAL_HEADER_HINTS.items():
            if any(hint == hl or hint in hl for hint in hints):
                matched = canon
                break
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
        if key in ("phone", "whatsapp") and not is_phone_like(val):
            flags.append(f"NON_PHONE_IN_{key.upper()}")
        if key == "email":
            m = EMAIL_RE.search(val)
            if not m:
                flags.append("NON_EMAIL_IN_EMAIL")
            else:
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


def ingest_files(
    inputs: List[Path],
    out_dir: Path,
    stream_mb_threshold: float = 10.0,
    unmerge_ffill: bool = True,
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

        sheet_streams: List[Tuple[str, List[Tuple[int, List[str]]]]] = []
        if ext == ".csv":
            rows = list(enumerate(read_csv_rows(path), start=1))
            sheet_streams.append(("CSV", rows))
            mode = "csv_chunked_stream"
        elif ext in (".xlsx", ".xlsm"):
            if unmerge_ffill and size_mb < stream_mb_threshold:
                mode = "xlsx_unmerge_ffill"
                sheet_dict = unmerge_and_ffill_xlsx(path)
                for sname, srows in sheet_dict.items():
                    sheet_streams.append((sname, [(i + 1, r) for i, r in enumerate(srows)]))
            else:
                mode = "xlsx_read_only_stream"
                by_sheet: Dict[str, List[Tuple[int, List[str]]]] = {}
                for sname, excel_row, r in stream_xlsx_rows(path):
                    by_sheet.setdefault(sname, []).append((excel_row, r))
                for sname, srows in by_sheet.items():
                    sheet_streams.append((sname, srows))
        else:
            raise ValueError(f"Unsupported extension '{ext}' for {path}. Convert .xls/.xlsb to .xlsx or .csv first.")

        for sname, indexed_rows in sheet_streams:
            if not indexed_rows:
                continue
            raw_only = [r for _, r in indexed_rows]
            headers, data_rows_plain = strip_empty_trailing_cols(raw_only)
            if not headers:
                continue
            header_as_data = looks_like_header_as_data(headers)
            if header_as_data:
                data_pairs = indexed_rows
                headers = [f"col_{i}" for i in range(len(headers))]
            else:
                # Re-align indices carefully
                col_count = len(headers)
                data_pairs = [
                    (erow, r[:col_count] + [""] * max(0, col_count - len(r)))
                    for erow, r in indexed_rows[1:]
                ]
            col_map = infer_column_map(headers, [r for _, r in data_pairs[:50]])
            sheet_count = 0

            for excel_row, cells in data_pairs:
                row_dict: Dict[str, str] = {}
                for c_idx, c_val in enumerate(cells):
                    canon = col_map.get(c_idx, f"col_{c_idx}")
                    if canon in row_dict and row_dict[canon] and c_val:
                        if c_val not in row_dict[canon]:
                            row_dict[canon] = f"{row_dict[canon]} | {c_val}"
                    elif c_val:
                        row_dict[canon] = c_val
                if not any(row_dict.values()):
                    continue

                flags = detect_row_anomalies(row_dict, cells)
                rec = {
                    "record_id": len(records) + 1,
                    "source_file": path.name,
                    "source_sheet": sname,
                    "source_row": excel_row,
                    "header_as_data": header_as_data,
                    "fields": row_dict,
                    "raw_cells": cells,
                    "flags": flags,
                }
                records.append(rec)
                sheet_count += 1
                if flags:
                    audit_batch.append(rec)

                # Index every >=7 digit run & email for Stage 4 cross-verification
                joined_raw = " | ".join(cells)
                for m in re.finditer(r"[\d٠-٩۰-۹][\d٠-٩۰-۹\s\-+().]{5,20}[\d٠-٩۰-۹]", joined_raw):
                    d = digits_only(m.group(0))
                    if 7 <= len(d) <= 15:
                        tail = d[-9:] if len(d) >= 9 else d
                        raw_phone_index.setdefault(tail, []).append({
                            "file": path.name,
                            "sheet": sname,
                            "row": excel_row,
                            "digits": d,
                            "snippet": joined_raw[:280],
                        })
                for em in EMAIL_RE.findall(joined_raw):
                    raw_email_index.setdefault(em.lower(), []).append({
                        "file": path.name,
                        "sheet": sname,
                        "row": excel_row,
                        "snippet": joined_raw[:280],
                    })

            file_stats.append({
                "file": path.name,
                "sheet": sname,
                "mode": mode,
                "size_mb": round(size_mb, 3),
                "rows_extracted": sheet_count,
                "header_as_data": header_as_data,
            })

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
    parser.add_argument("--input", "-i", nargs="+", required=True, help="Input .xlsx or .csv file(s) / folders.")
    parser.add_argument("--work-dir", "-w", default="./_clean_work", help="Directory for extracted JSON artifacts.")
    parser.add_argument("--stream-mb", type=float, default=10.0, help="Size MB threshold to switch XLSX to read_only stream.")
    parser.add_argument("--no-unmerge", action="store_true", help="Always stream XLSX without merged-cell ffill pass.")
    args = parser.parse_args()

    files: List[Path] = []
    for item in args.input:
        p = Path(item)
        if p.is_dir():
            for ext in ("*.xlsx", "*.xlsm", "*.csv"):
                files.extend(sorted(p.rglob(ext)))
        else:
            files.append(p)

    summary = ingest_files(
        files,
        Path(args.work_dir),
        stream_mb_threshold=args.stream_mb,
        unmerge_ffill=not args.no_unmerge,
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())

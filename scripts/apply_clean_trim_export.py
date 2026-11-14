#!/usr/bin/env python3
"""
Deterministic Realigner, Split, Strict TRIM, Safety Guard & Adaptive Exporter
Part of `spreadsheet-realign-clean-verify` by Eslam Elgindy (MIT License).

Solves:
- Stage 0 CLI / interactive pre-flight configuration (format, destination, quarantine)
- Applies Agent 1 semantic JSON overrides (`agent1_semantic_decisions.json`) + deterministic rules
- Arabic compare-key deduplication (أ/إ/آ -> ا, ى -> ي, ة -> ه) without altering cell spelling
- Safe phone normalization (Arabic-Indic digits, country codes intact, tail matching, text '@')
- Excel serial date & timestamp standardization -> YYYY-MM-DD (native Excel date)
- Mandatory `strict_trim`, Excel error scrubbing, pipe dedup, CSV formula-injection guard
- Reachability gate + `_needs_review.csv` quarantine export
- Adaptive Excel export (`Workbook(write_only=True)` > 50,000 rows) + CSV UTF-8 BOM (`utf-8-sig`)
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill

INVISIBLE_RE = re.compile(r"[\u200b-\u200f\u202a-\u202e\ufeff\xa0]")
EXCEL_ERR_RE = re.compile(r"^#(REF!|DIV/0!|VALUE!|NAME\?|N/A|NUM!|NULL!|ERROR!)$", re.I)
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,24}")
SCI_NUM_RE = re.compile(r"^(\d+(?:\.\d+)?)[eE][+\-]?\d+$")
TRANS_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")

UNIT_RE = re.compile(
    r"(studio|ssstudio|استديو|ستوديو|\b\d+\s*[_\-]?\s*bed(room)?s?\b|\b\d+\s*br\b|"
    r"غرفة|غرف|شقة|فيلا|تاون\s*هاوس|apartment|villa|townhouse|penthouse|duplex)",
    re.I,
)
AUDIENCE_TAG_RE = re.compile(r"^(arabic|gcc|europe|english|ar|en|all)$", re.I)
YES_CITY_RE = re.compile(r"^(نعم|yes)\s*[-–:]\s*(.+)$", re.I)
PRONOUN_NOISE = {"أنا", "انا", "me", "none", "null", "na", "n/a", "test", "duplicae", "duplicate"}
DIVIDER_NAME_RE = re.compile(
    r"^(total|grand\s*total|subtotal|sub-total|tower\s*\(?.*\)?|\d+-bedroom(\+study)?|studio|ssstudio)$",
    re.I,
)
SUSPECT_EMAIL_DOMAINS = {
    "gamil.com", "gimil.com", "glail.com", "gmul.com", "gamli.com", "gamail.com",
    "gmal.com", "gmaill.com", "gmail.co", "gmail.con", "gmail.coma", "gmail.come",
    "gmail.comm", "gmail.coms", "gmail.comj", "gmail.com0", "gmeal.come",
    "hotmil.com", "hotmal.come", "hotmaol.com", "hotmail.con", "hotmail.coms",
    "hotmail.comi", "icloude.com", "icloud.con", "yahoo.vom", "yaho.com",
}

STANDARD_HEADERS = [
    ("name", "الاسم بالكامل"),
    ("phone", "رقم الهاتف"),
    ("whatsapp", "رقم الواتساب"),
    ("extra_phone", "هاتف إضافي"),
    ("email", "البريد الإلكتروني"),
    ("campaign", "اسم الحملة والمصدر"),
    ("project", "المشروع والوحدة المطلوبة"),
    ("budget", "الميزانية والسعر"),
    ("feedback", "حالة الفيدباك (القائمة)"),
    ("notes", "ملاحظات وتعليقات الموظف"),
    ("pref_time", "وقت التواصل المفضل"),
    ("country", "الدولة والمدينة"),
    ("date", "تاريخ التسجيل"),
]


def strict_trim(val: Any) -> str:
    """Apply universal strict TRIM + invisible/bidi removal + Excel error scrubbing."""
    if val is None:
        return ""
    s = INVISIBLE_RE.sub(" ", str(val))
    s = s.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    s = re.sub(r"\s+", " ", s).strip()
    if s.lower() in ("nan", "none", "null") or EXCEL_ERR_RE.match(s):
        return ""
    return s


def sanitize_csv_formula_cell(val: str, is_phone_col: bool = False) -> str:
    """Prevent spreadsheet CSV formula injection while preserving valid phone/email tokens."""
    if not val or is_phone_col:
        return val
    if val[0] in ("=", "+", "@") or (val[0] == "-" and not re.match(r"^-\d+(\.\d+)?$", val)):
        return f"'{val}"
    return val


def dedup_pipe_tokens(val: str) -> str:
    val = strict_trim(val)
    if "|" not in val:
        return val
    seen: set[str] = set()
    kept: List[str] = []
    for part in val.split("|"):
        p = strict_trim(part)
        if p and p.lower() not in seen:
            seen.add(p.lower())
            kept.append(p)
    return " | ".join(kept)


def arabic_compare_key(text: str) -> str:
    """Normalize Arabic letters solely for matching/dedup keys (never mutate output cells)."""
    t = strict_trim(text).lower()
    t = re.sub(r"[\u064b-\u065f\u0670]", "", t)
    t = t.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    t = t.replace("ى", "ي").replace("ة", "ه").replace("_", " ")
    return re.sub(r"\s+", " ", t).strip()


COUNTRY_CALLING_CODES: Dict[str, str] = {
    "uae": "971", "united arab emirates": "971", "dubai": "971", "abu dhabi": "971", "sharjah": "971", "ajman": "971", "الامارات": "971", "دبي": "971", "ابوظبي": "971", "الشارقه": "971",
    "saudi": "966", "ksa": "966", "riyadh": "966", "jeddah": "966", "dammam": "966", "السعوديه": "966", "الرياض": "966", "جده": "966", "الدمام": "966",
    "egypt": "20", "cairo": "20", "alexandria": "20", "مصر": "20", "القاهره": "20", "الاسكندريه": "20",
    "kuwait": "965", "الكويت": "965", "qatar": "974", "قطر": "974", "bahrain": "973", "البحرين": "973",
    "oman": "968", "عمان": "968", "jordan": "962", "الاردن": "962", "iraq": "964", "العراق": "964",
}


def normalize_phone_digits(raw: str, country_hint: str = "") -> str:
    if not raw:
        return ""
    s = strict_trim(raw).translate(TRANS_DIGITS)
    if EXCEL_ERR_RE.match(s):
        return ""
    # Recover scientific notation if possible without throwing
    m_sci = SCI_NUM_RE.match(s)
    if m_sci:
        try:
            s = str(int(float(s)))
        except ValueError:
            return ""
    # Clean composite prefixes like "nullI971|50..."
    s = re.sub(r"^(p:|wa:|tel:|nulli?)", "", s, flags=re.I)
    d = re.sub(r"\D", "", s)
    if len(d) < 7 or len(d) > 15:
        return ""
    if d.startswith("00"):
        d = d[2:]
    # Restore local calling code ONLY when number clearly uses local '0'-prefix & matches country hint
    if d.startswith("0") and 9 <= len(d) <= 11 and country_hint:
        ckey = arabic_compare_key(country_hint)
        for token, code in COUNTRY_CALLING_CODES.items():
            if token in ckey:
                local_rest = d[1:]
                if code == "20" and local_rest.startswith("1") and len(local_rest) == 10:
                    return code + local_rest
                if code in ("971", "966") and local_rest.startswith("5") and len(local_rest) == 9:
                    return code + local_rest
                if code in ("964", "962") and local_rest.startswith("7") and len(local_rest) >= 9:
                    return code + local_rest
    return d


def phone_tail(d: str) -> str:
    return d[-9:] if len(d) >= 9 else d


def parse_date_iso(raw: str) -> str:
    s = strict_trim(raw).translate(TRANS_DIGITS)
    if not s:
        return ""
    # Excel serial integer (1900 epoch) between ~2000 and ~2040
    if re.fullmatch(r"\d{5}(?:\.0+)?", s):
        serial = int(float(s))
        if 36526 <= serial <= 55000:
            return (dt.date(1899, 12, 30) + dt.timedelta(days=serial)).strftime("%Y-%m-%d")
    # ISO datetime prefix
    m = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[T\s].*)?$", s)
    if m:
        try:
            return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3))).strftime("%Y-%m-%d")
        except ValueError:
            return ""
    m2 = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})(?:[T\s].*)?$", s)
    if m2:
        a, b, y = int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
        day, mon = (b, a) if b > 12 else (a, b)
        try:
            return dt.date(y, mon, day).strftime("%Y-%m-%d")
        except ValueError:
            return ""
    return ""


def deterministic_realign_row(rec: Dict[str, Any], decision: Optional[Dict[str, Any]] = None) -> Dict[str, str]:
    """Combine Agent 1 overrides (if supplied) with deterministic multi-pattern realignment."""
    if decision and isinstance(decision.get("cleaned"), dict):
        base = {k: dedup_pipe_tokens(v) for k, v in decision["cleaned"].items()}
        for k, _ in STANDARD_HEADERS:
            base.setdefault(k, "")
        return {k: strict_trim(base[k]) for k, _ in STANDARD_HEADERS}

    fields = dict(rec.get("fields", {}))
    out: Dict[str, str] = {k: "" for k, _ in STANDARD_HEADERS}
    notes_parts: List[str] = []
    phones_found: List[str] = []

    raw_phone = strict_trim(fields.get("phone", ""))
    raw_wa = strict_trim(fields.get("whatsapp", ""))
    country_hint = strict_trim(fields.get("country", ""))
    dup_phone_tails: set[str] = set()

    def absorb_phone(candidate: str, target_slot: Optional[str] = None) -> bool:
        d = normalize_phone_digits(candidate, country_hint=country_hint)
        if not d:
            return False
        t = phone_tail(d)
        for idx, existing in enumerate(phones_found):
            if phone_tail(existing) == t:
                dup_phone_tails.add(t)
                if len(d) > len(existing):
                    phones_found[idx] = d
                if target_slot == "whatsapp":
                    out["whatsapp"] = phones_found[idx]
                return True
        phones_found.append(d)
        if target_slot == "whatsapp" and not out["whatsapp"]:
            out["whatsapp"] = d
        return True

    # 1. Explicit Phone & WhatsApp columns
    for tok in ([strict_trim(x) for x in raw_phone.split("|")] if raw_phone else []):
        if EMAIL_RE.search(tok) and not out["email"]:
            out["email"] = EMAIL_RE.search(tok).group(0).lower()
        elif parse_date_iso(tok) and not out["date"]:
            out["date"] = parse_date_iso(tok)
        elif not absorb_phone(tok, "phone") and ("?" in tok or "؟" in tok):
            notes_parts.append(f"استفسار العميل: {tok}")

    for tok in ([strict_trim(x) for x in raw_wa.split("|")] if raw_wa else []):
        if UNIT_RE.search(tok) and not out["project"]:
            out["project"] = strict_trim(tok.replace("_", " "))
        elif not absorb_phone(tok, "whatsapp") and len(tok) > 2:
            notes_parts.append(tok)

    # 2. Email column check
    raw_email = strict_trim(fields.get("email", ""))
    if raw_email:
        m_em = EMAIL_RE.search(raw_email)
        if m_em:
            out["email"] = m_em.group(0).lower()
            domain = out["email"].split("@")[-1]
            if domain in SUSPECT_EMAIL_DOMAINS:
                notes_parts.append(f"EMAIL_TYPO_SUSPECT: {out['email']}")
        elif absorb_phone(raw_email):
            pass
        elif "campaign" in raw_email.lower() and not out["campaign"]:
            out["campaign"] = raw_email

    # 3. Date column check
    raw_date = strict_trim(fields.get("date", ""))
    if raw_date:
        iso = parse_date_iso(raw_date)
        if iso:
            out["date"] = iso
        elif absorb_phone(raw_date):
            pass

    # 4. Name column check (token-level for pipe-merged cells)
    raw_name = strict_trim(fields.get("name", ""))
    for tok in ([strict_trim(x) for x in raw_name.split("|")] if raw_name else []):
        if not tok:
            continue
        if "?" in tok or "؟" in tok:
            notes_parts.append(f"استفسار العميل: {tok}")
        elif absorb_phone(tok):
            pass
        elif tok.lower() in PRONOUN_NOISE or DIVIDER_NAME_RE.match(tok):
            pass
        elif len(tok) > 1 and not out["name"]:
            out["name"] = tok

    # 5. Project / Rooms (handle pipe-merged headers & shifted human names / budgets / yes-city)
    raw_proj = strict_trim(fields.get("project", ""))
    for token in ([strict_trim(x) for x in raw_proj.split("|")] if raw_proj else []):
        if not token or AUDIENCE_TAG_RE.match(token):
            continue
        m_yc = YES_CITY_RE.match(token)
        if m_yc:
            if not out["country"]:
                out["country"] = strict_trim(m_yc.group(2))
            continue
        if re.fullmatch(r"\d{2,10}", token):
            if not out["budget"]:
                out["budget"] = token
            continue
        if UNIT_RE.search(token):
            out["project"] = strict_trim(token.replace("_", " "))
        elif "demo " in token.lower() or "campaign" in token.lower():
            if "owner" in token.lower() and not out["name"]:
                out["name"] = token
            elif not out["campaign"]:
                out["campaign"] = token
        elif not out["campaign"] and not out["name"]:
            # Decide if client-style full name (3+ words) vs short uppercase project tag
            if token.isupper() or any(w in token.lower() for w in ("demo", "rivera", "azizi", "sobh", "venice", "milan")):
                out["campaign"] = token
            else:
                out["name"] = token
        elif not out["name"]:
            out["name"] = token
        elif not out["campaign"]:
            out["campaign"] = token

    # 6. Country / City
    raw_cty = strict_trim(fields.get("country", ""))
    if raw_cty:
        m_yc = YES_CITY_RE.match(raw_cty)
        if m_yc:
            out["country"] = strict_trim(m_yc.group(2))
        elif "campaign" in raw_cty.lower():
            out["campaign"] = raw_cty
        else:
            out["country"] = raw_cty

    # 7. Remaining straightforward fields
    for slot in ("campaign", "budget", "feedback", "pref_time"):
        val = dedup_pipe_tokens(fields.get(slot, ""))
        if val and not out[slot]:
            if EMAIL_RE.search(val) and not out["email"]:
                out["email"] = EMAIL_RE.search(val).group(0).lower()
            elif absorb_phone(val):
                pass
            elif slot == "campaign" and AUDIENCE_TAG_RE.match(val):
                continue
            else:
                out[slot] = val

    raw_notes = dedup_pipe_tokens(fields.get("notes", ""))
    if raw_notes:
        notes_parts.append(raw_notes)

    # Final phone slot assignment (align tails cleanly)
    if phones_found:
        out["phone"] = phones_found[0]
        if out["whatsapp"]:
            t_wa = phone_tail(out["whatsapp"])
            for p in phones_found:
                if phone_tail(p) == t_wa and len(p) >= len(out["whatsapp"]):
                    out["whatsapp"] = p
        elif phone_tail(out["phone"]) in dup_phone_tails:
            out["whatsapp"] = out["phone"]
        for p in phones_found[1:]:
            if phone_tail(p) != phone_tail(out["phone"]) and not out["extra_phone"]:
                out["extra_phone"] = p

    # If WA matched Phone on short tail after assignment
    if out["whatsapp"] and out["phone"] and phone_tail(out["whatsapp"]) == phone_tail(out["phone"]):
        if len(out["phone"]) >= len(out["whatsapp"]):
            out["whatsapp"] = out["phone"]

    out["notes"] = strict_trim(" | ".join(dict.fromkeys([n for n in notes_parts if strict_trim(n)])))
    return {k: strict_trim(out[k]) for k, _ in STANDARD_HEADERS}


def merge_row(existing: Dict[str, str], incoming: Dict[str, str]) -> Dict[str, str]:
    merged = dict(existing)
    for k, v in incoming.items():
        if not v:
            continue
        if not merged.get(k):
            merged[k] = v
        elif k in ("notes", "feedback") and arabic_compare_key(v) not in arabic_compare_key(merged[k]):
            merged[k] = strict_trim(f"{merged[k]} | {v}")
        elif k in ("phone", "whatsapp") and len(v) > len(merged[k]) and phone_tail(v) == phone_tail(merged[k]):
            merged[k] = v
    return merged


def write_csv_bom(path: Path, headers: List[Tuple[str, str]], rows: List[Dict[str, str]]) -> None:
    phone_keys = {"phone", "whatsapp", "extra_phone"}
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow([label for _, label in headers])
        for row in rows:
            writer.writerow([
                sanitize_csv_formula_cell(strict_trim(row.get(k, "")), is_phone_col=(k in phone_keys))
                for k, _ in headers
            ])


def write_xlsx_adaptive(
    path: Path,
    headers: List[Tuple[str, str]],
    rows: List[Dict[str, str]],
    write_only_threshold: int = 50_000,
) -> str:
    phone_keys = {"phone", "whatsapp", "extra_phone"}
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF")

    if len(rows) > write_only_threshold:
        wb = openpyxl.Workbook(write_only=True)
        ws = wb.create_sheet("Cleaned")
        hdr_cells = []
        for _, label in headers:
            c = WriteOnlyCell(ws, value=label)
            c.fill = header_fill
            c.font = header_font
            hdr_cells.append(c)
        ws.append(hdr_cells)
        for r in rows:
            out_cells = []
            for k, _ in headers:
                v = strict_trim(r.get(k, ""))
                if k == "date" and re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
                    c = WriteOnlyCell(ws, value=dt.date.fromisoformat(v))
                    c.number_format = "yyyy-mm-dd"
                else:
                    c = WriteOnlyCell(ws, value=v)
                    if k in phone_keys:
                        c.number_format = "@"
                out_cells.append(c)
            ws.append(out_cells)
        wb.save(path)
        return "xlsx_write_only_stream"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cleaned"
    ws.freeze_panes = "A2"
    for col_idx, (_, label) in enumerate(headers, start=1):
        c = ws.cell(1, col_idx, value=label)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")

    col_widths = [len(label) + 4 for _, label in headers]
    for row_idx, r in enumerate(rows, start=2):
        for col_idx, (k, _) in enumerate(headers, start=1):
            v = strict_trim(r.get(k, ""))
            if k == "date" and re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
                c = ws.cell(row_idx, col_idx, value=dt.date.fromisoformat(v))
                c.number_format = "yyyy-mm-dd"
            else:
                c = ws.cell(row_idx, col_idx, value=v)
                if k in phone_keys:
                    c.number_format = "@"
            col_widths[col_idx - 1] = min(max(col_widths[col_idx - 1], len(v) + 3), 45)

    for col_idx, w in enumerate(col_widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = w
    if rows:
        last_col = openpyxl.utils.get_column_letter(len(headers))
        ws.auto_filter.ref = f"A1:{last_col}{len(rows) + 1}"
    wb.save(path)
    return "xlsx_styled"


def prompt_stage0_if_interactive(args: argparse.Namespace) -> None:
    if not args.interactive:
        return
    fmt = input("Output format? [both / xlsx / csv] (default: both): ").strip().lower()
    if fmt in ("both", "xlsx", "csv"):
        args.format = fmt
    dest = input(f"Destination folder? (default: {args.out_dir}): ").strip()
    if dest:
        args.out_dir = dest
    dmode = input(f"Granularity / dedup mode? [lead / unit / none] (default: {args.dedup_mode}): ").strip().lower()
    if dmode in ("lead", "unit", "none"):
        args.dedup_mode = dmode
    quar = input("Save excluded rows to _needs_review.csv? [Y/n] (default: Y): ").strip().lower()
    if quar == "n":
        args.no_quarantine = True


def main() -> int:
    parser = argparse.ArgumentParser(description="Realign, clean, strict-TRIM, and export Excel/CSV.")
    parser.add_argument("--base", "-b", required=True, help="Path to extracted_base.json")
    parser.add_argument("--decisions", "-d", help="Optional path to agent1_semantic_decisions.json")
    parser.add_argument("--format", "-f", choices=("xlsx", "csv", "both"), default="both", help="Stage 0 output format.")
    parser.add_argument("--out-dir", "-o", default="./cleaned_output", help="Stage 0 destination folder.")
    parser.add_argument("--prefix", "-p", default="cleaned_dataset", help="Output filename prefix.")
    parser.add_argument("--dedup-mode", choices=("lead", "unit", "none"), default="lead", help="Deduplication granularity.")
    parser.add_argument("--no-quarantine", action="store_true", help="Disable _needs_review.csv quarantine export.")
    parser.add_argument("--write-only-rows", type=int, default=50_000, help="Excel write_only streaming row threshold.")
    parser.add_argument("--interactive", action="store_true", help="Run Stage 0 interactive terminal prompt.")
    args = parser.parse_args()

    prompt_stage0_if_interactive(args)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    base_records = json.loads(Path(args.base).read_text(encoding="utf-8"))
    decisions: Dict[str, Any] = {}
    if args.decisions and Path(args.decisions).exists():
        decisions = json.loads(Path(args.decisions).read_text(encoding="utf-8"))

    cleaned_rows: List[Dict[str, str]] = []
    quarantine_rows: List[Dict[str, str]] = []
    dedup_index: Dict[str, int] = {}
    stats = {
        "raw_records_input": len(base_records),
        "header_as_data_recovered": sum(1 for r in base_records if r.get("header_as_data")),
        "rows_with_anomalies": sum(1 for r in base_records if r.get("flags")),
        "merged_duplicates": 0,
        "quarantined_rows": 0,
        "cells_strict_trimmed": 0,
    }

    for rec in base_records:
        rid = str(rec.get("record_id", ""))
        dec = decisions.get(rid)
        if dec and dec.get("action") == "drop":
            stats["quarantined_rows"] += 1
            quarantine_rows.append({
                "source_file": strict_trim(rec.get("source_file", "")),
                "source_sheet": strict_trim(rec.get("source_sheet", "")),
                "source_row": str(rec.get("source_row", "")),
                "exclusion_reason": strict_trim(dec.get("reason", "dropped_by_agent1")),
                "raw_snapshot": strict_trim(" | ".join(rec.get("raw_cells", [])))[:300],
            })
            continue

        row = deterministic_realign_row(rec, dec)
        has_phone = bool(row["phone"])
        has_email = bool(EMAIL_RE.fullmatch(row["email"]))
        if not has_phone and not has_email:
            stats["quarantined_rows"] += 1
            quarantine_rows.append({
                "source_file": strict_trim(rec.get("source_file", "")),
                "source_sheet": strict_trim(rec.get("source_sheet", "")),
                "source_row": str(rec.get("source_row", "")),
                "exclusion_reason": "unreachable_no_phone_or_email",
                "raw_snapshot": strict_trim(" | ".join(rec.get("raw_cells", [])))[:300],
            })
            continue

        # Enforce strict_trim on every cell of kept row
        for k, _ in STANDARD_HEADERS:
            row[k] = strict_trim(row.get(k, ""))
            stats["cells_strict_trimmed"] += 1

        if args.dedup_mode == "none":
            cleaned_rows.append(row)
            continue

        if args.dedup_mode == "unit":
            key = f"U:{phone_tail(row['phone'])}|{arabic_compare_key(row['project'])}|{arabic_compare_key(row['name'])}"
        else:
            key = f"P:{phone_tail(row['phone'])}" if has_phone else f"E:{row['email'].lower()}"

        if key in dedup_index:
            idx = dedup_index[key]
            cleaned_rows[idx] = merge_row(cleaned_rows[idx], row)
            stats["merged_duplicates"] += 1
        else:
            dedup_index[key] = len(cleaned_rows)
            cleaned_rows.append(row)

    # Sort descending by date then name
    cleaned_rows.sort(key=lambda r: (r["date"] == "", r["date"], r["name"]))
    cleaned_rows.sort(key=lambda r: r["date"], reverse=True)
    stats["final_output_rows"] = len(cleaned_rows)

    outputs: List[str] = []
    if args.format in ("csv", "both"):
        csv_path = out_dir / f"{args.prefix}.csv"
        write_csv_bom(csv_path, STANDARD_HEADERS, cleaned_rows)
        outputs.append(str(csv_path))

    if args.format in ("xlsx", "both"):
        xlsx_path = out_dir / f"{args.prefix}.xlsx"
        xlsx_mode = write_xlsx_adaptive(xlsx_path, STANDARD_HEADERS, cleaned_rows, args.write_only_rows)
        stats["xlsx_mode"] = xlsx_mode
        outputs.append(str(xlsx_path))

    if not args.no_quarantine and quarantine_rows:
        q_path = out_dir / f"{args.prefix}_needs_review.csv"
        q_headers = [
            ("source_file", "source_file"),
            ("source_sheet", "source_sheet"),
            ("source_row", "source_row"),
            ("exclusion_reason", "exclusion_reason"),
            ("raw_snapshot", "raw_snapshot"),
        ]
        write_csv_bom(q_path, q_headers, quarantine_rows)
        outputs.append(str(q_path))

    stats["outputs"] = outputs
    (out_dir / f"{args.prefix}_clean_metrics.json").write_text(
        json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(stats, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())

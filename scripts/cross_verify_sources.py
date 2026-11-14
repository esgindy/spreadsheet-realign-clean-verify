#!/usr/bin/env python3
"""
100% Phone-by-Phone Source Cross-Verifier & Stage 5 Summary Generator
Part of `spreadsheet-realign-clean-verify` by Eslam Elgindy (MIT License).

Audits:
- Matches EVERY phone number in the cleaned CSV back against raw_source_index.json
- Verifies Name, Email, Country/City, Phone, and WhatsApp fidelity against source snippets
- Checks zero remaining Excel errors (#REF!, #ERROR!, #N/A), zero invisible/bidi chars,
  zero untrimmed whitespace, zero invalid date strings, zero unreachable rows
- Prints a compact Stage 5 Numeric Summary Report
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List

INVISIBLE_RE = re.compile(r"[\u200b-\u200f\u202a-\u202e\ufeff\xa0]")
EXCEL_ERR_RE = re.compile(r"^#(REF!|DIV/0!|VALUE!|NAME\?|N/A|NUM!|NULL!|ERROR!)$", re.I)
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,24}")


def phone_tail(val: str) -> str:
    d = re.sub(r"\D", "", val or "")
    return d[-9:] if len(d) >= 9 else d


def arabic_norm(text: str) -> str:
    t = (text or "").strip().lower()
    t = t.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    t = t.replace("ى", "ي").replace("ة", "ه")
    return re.sub(r"\s+", " ", t)


def verify_dataset(
    cleaned_csv: Path,
    raw_index_path: Path,
    metrics_path: Path | None = None,
) -> Dict[str, Any]:
    raw_idx = json.loads(raw_index_path.read_text(encoding="utf-8"))
    raw_phones: Dict[str, List[Dict[str, Any]]] = raw_idx.get("phones", {})
    raw_emails: Dict[str, List[Dict[str, Any]]] = raw_idx.get("emails", {})

    with cleaned_csv.open("r", encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.DictReader(fh))

    phones_checked = 0
    phones_matched = 0
    phone_mismatches: List[Dict[str, Any]] = []
    format_defects = {
        "untrimmed_or_double_space": 0,
        "invisible_or_bidi_chars": 0,
        "excel_error_tokens": 0,
        "invalid_date_format": 0,
        "unreachable_rows": 0,
        "field_cross_mismatches": 0,
    }

    for r_idx, row in enumerate(rows, start=2):
        for col, val in row.items():
            v = val or ""
            if v != v.strip() or re.search(r"  +", v):
                format_defects["untrimmed_or_double_space"] += 1
            if INVISIBLE_RE.search(v):
                format_defects["invisible_or_bidi_chars"] += 1
            if EXCEL_ERR_RE.match(v.strip()):
                format_defects["excel_error_tokens"] += 1

        date_val = (row.get("تاريخ التسجيل") or "").strip()
        if date_val and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", date_val):
            format_defects["invalid_date_format"] += 1

        ph = (row.get("رقم الهاتف") or "").strip()
        wa = (row.get("رقم الواتساب") or "").strip()
        ex = (row.get("هاتف إضافي") or "").strip()
        em = (row.get("البريد الإلكتروني") or "").strip().lower()
        name = (row.get("الاسم بالكامل") or "").strip()
        cty = (row.get("الدولة والمدينة") or "").strip()

        if len(re.sub(r"\D", "", ph)) < 7 and not EMAIL_RE.fullmatch(em):
            format_defects["unreachable_rows"] += 1

        row_matched_hits: List[Dict[str, Any]] = []
        for pcol, pval in (("phone", ph), ("whatsapp", wa), ("extra_phone", ex)):
            if not pval:
                continue
            phones_checked += 1
            tail = phone_tail(pval)
            hits = raw_phones.get(tail, [])
            if hits:
                phones_matched += 1
                row_matched_hits.extend(hits)
            else:
                phone_mismatches.append({"row": r_idx, "col": pcol, "value": pval, "reason": "tail_not_in_raw"})

        if row_matched_hits:
            combined_raw = " ".join(arabic_norm(h.get("snippet", "")) for h in row_matched_hits)
            if em and em not in combined_raw and em not in raw_emails:
                format_defects["field_cross_mismatches"] += 1
            if name and arabic_norm(name) not in combined_raw:
                # Check token overlap for Arabic/English names
                tokens = [t for t in arabic_norm(name).split() if len(t) >= 2]
                if tokens and not any(t in combined_raw for t in tokens):
                    format_defects["field_cross_mismatches"] += 1
            if cty and arabic_norm(cty) not in combined_raw:
                cty_tokens = [t for t in arabic_norm(cty).split() if len(t) >= 2]
                if cty_tokens and not any(t in combined_raw for t in cty_tokens):
                    format_defects["field_cross_mismatches"] += 1

    metrics = {}
    if metrics_path and metrics_path.exists():
        metrics = json.loads(metrics_path.read_text(encoding="utf-8"))

    total_defects = sum(format_defects.values()) + len(phone_mismatches)
    return {
        "cleaned_file": str(cleaned_csv),
        "total_output_rows": len(rows),
        "phones_checked": phones_checked,
        "phones_matched_raw": phones_matched,
        "phone_match_rate": round(100.0 * phones_matched / phones_checked, 2) if phones_checked else 100.0,
        "phone_mismatches": phone_mismatches[:20],
        "format_defects": format_defects,
        "all_zero_defects_passed": total_defects == 0,
        "stage3_metrics": metrics,
        "raw_source_stats": raw_idx.get("stats", []),
    }


def format_stage5_summary(report: Dict[str, Any]) -> str:
    m = report.get("stage3_metrics", {})
    raw_stats = report.get("raw_source_stats", [])
    defects = report.get("format_defects", {})
    hidden_rows = sum(s.get("hidden_rows_skipped", 0) for s in raw_stats)
    hidden_sheets = sum(s.get("hidden_sheets_skipped", 0) for s in raw_stats)
    stacked_cnt = sum(1 for s in raw_stats if s.get("stacked_headers_flattened"))
    lines = [
        "## Stage 5 — Concise Numeric Summary Report",
        f"- **File Verified:** `{report['cleaned_file']}`",
        f"- **Row Funnel:** Raw `{m.get('raw_records_input', '?')}` -> Recovered header rows `{m.get('header_as_data_recovered', 0)}` -> Merged `{m.get('merged_duplicates', 0)}` -> Quarantined `{m.get('quarantined_rows', 0)}` (fake placeholders: `{m.get('fake_placeholders_quarantined', 0)}`) -> **Final `{report['total_output_rows']}` rows**",
        f"- **Strict TRIM & Rescues:** `{m.get('cells_strict_trimmed', 0)}` cells (`{m.get('xlsx_mode', 'n/a')}`), hyperlinks/comments applied `{m.get('hyperlinks_and_comments_applied', 0)}`, stacked headers flattened `{stacked_cnt}`, hidden skipped `{hidden_rows}r/{hidden_sheets}s`, 1904-epoch `{m.get('epoch_1904_records', 0)}`",
        f"- **Raw Phone Cross-Check:** `{report['phones_matched_raw']} / {report['phones_checked']}` (**{report['phone_match_rate']}%**)",
        f"- **Zero-Defect Gate:** `all_zero_defects_passed = {report['all_zero_defects_passed']}` (untrimmed={defects.get('untrimmed_or_double_space', 0)}, bidi={defects.get('invisible_or_bidi_chars', 0)}, excel_err={defects.get('excel_error_tokens', 0)}, bad_dates={defects.get('invalid_date_format', 0)}, unreachable={defects.get('unreachable_rows', 0)}, field_mismatch={defects.get('field_cross_mismatches', 0)})",
    ]
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description="100% phone-by-phone cross-verification & Stage 5 report.")
    parser.add_argument("--cleaned-csv", "-c", required=True, help="Path to cleaned CSV (utf-8-sig).")
    parser.add_argument("--raw-index", "-r", required=True, help="Path to raw_source_index.json.")
    parser.add_argument("--metrics", "-m", help="Optional path to *_clean_metrics.json.")
    parser.add_argument("--summary-md", action="store_true", help="Print Stage 5 Markdown summary alongside JSON.")
    args = parser.parse_args()

    report = verify_dataset(
        Path(args.cleaned_csv),
        Path(args.raw_index),
        Path(args.metrics) if args.metrics else None,
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if args.summary_md:
        print("\n" + format_stage5_summary(report))
    return 0 if report["all_zero_defects_passed"] else 2


if __name__ == "__main__":
    sys.exit(main())

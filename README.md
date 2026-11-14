# `spreadsheet-realign-clean-verify`

**Cross-Platform AI Agent Skill for Messy & Large-Scale Spreadsheets (`.xlsx`, `.xlsm`, `.csv`, `.xls`, `.xlsb`) — `v1.2.0`**  
**مهارة ذكاء اصطناعي شاملة لتنقيح وإعادة اصطفاف وتنظيف والتحقق الكامل لملفات الإكسل والبيانات الضخمة**

[![Version: v1.2.0](https://img.shields.io/badge/Version-v1.2.0-success.svg)](#english-documentation)
[![Problems Solved: 21](https://img.shields.io/badge/Problems%20Solved-21-blue.svg)](#all-21-solved-problems--production-capabilities)
[![Scale: 50k-1M+ Rows](https://img.shields.io/badge/Scale-50k--1M%2B%20Rows-purple.svg)](#architecture--workflow)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

> **Choose Language / اختر اللغة:**  
> [English Documentation](#english-documentation) · [الدليل الكامل باللغة العربية](#الدليل-الكامل-باللغة-العربية)

---

# English Documentation

## About

`spreadsheet-realign-clean-verify` is an industrial-grade, cross-agent **AI Agent Skill** designed to salvage corrupted, mid-sheet-shifted, merged, multi-encoding, or massive (`50k–1M+` rows) Excel and CSV workbooks. Before modifying or writing any deliverable, it conducts a brief **interactive pre-flight interview (Stage 0)**, orchestrates two dedicated subagents alongside memory-bounded deterministic Python pipelines, enforces universal strict `TRIM` on every output cell, isolates junk/fake rows into `_needs_review.csv`, cross-verifies **100% of phone numbers** back against raw sources, and concludes with a **concise numeric summary report (Stage 5)**.

---

## Architecture & Workflow

```text
  [Stage 0: Pre-Flight Interview]
   Output format (.xlsx / .csv / both) · Save destination · Split granularity (lead / unit) · Quarantine toggle · Hidden rows rule
                │
                ▼
  [Stage 1: Profile & Stream Extraction (`scripts/profile_and_extract.py`)]
   Memory-bounded streaming (`read_only` / `chunksize`) · Reset dimensions · Unmerge & ffill (<10 MB)
   Skip hidden rows/sheets · Rescued hyperlinks (`wa.me/`, `tel:`, `mailto:`) & cell comments
   Mac 1904 vs 1900 epoch detection · Auto CSV delimiter sniffing (`,` `;` `\t` `|`) · 2-tier header flattening
                │
        ┌───────┴───────┐
        ▼               ▼
  [Stage 2: Agent 1]   [Stage 3: Agent 2 (`scripts/apply_clean_trim_export.py`)]
  `Cell_Semantic_Reader` Strict `TRIM` · Fake placeholder & junk quarantine · Locale budget normalization
  Inspects anomaly flags Protected calling codes · Arabic compare-key dedup · CSV formula-injection guard
                         Adaptive `write_only=True` XLSX + `utf-8-sig` BOM CSV + `_needs_review.csv`
                │
                ▼
  [Stage 4 & 5: Cross-Verification & Concise Numeric Summary (`scripts/cross_verify_sources.py`)]
   100% phone tail audit vs `raw_source_index.json` · 0-defect gate · Compact numeric audit report
```

---

## All 21 Solved Problems & Production Capabilities

### Core Structural, Semantic & Big-Data Failures (`#1–#14`)
1. **Interactive Stage 0 Setup:** Confirms export format (`.xlsx`, `.csv`, or Both), destination folder (same folder vs. custom directory — never overwrites raw inputs), split granularity (`lead` vs. `unit`), and `_needs_review.csv` quarantine log before any file writes.
2. **Stacked Form Sub-Schemas & Column Shifts:** Classifies shifted cells by semantic value pattern (email in phone column, phone in project column, date in campaign column) instead of blindly trusting headers.
3. **Header-Row-As-Data Recovery:** Detects CRM/Excel sheets whose first row contains actual lead data (`col@domain.com`, `+971...`), promotes synthetic headers, and strips pandas `.1` suffixes without losing row 1.
4. **Merged Cells (`ffill`) & 200+ Phantom Columns:** Unmerges and forward-fills (`ffill`) multi-row merged groups and trims infinite phantom trailing columns (`Column200+`) automatically.
5. **Dual-Metric Workbook Streaming (`50k–1M+` Rows):** Calls `ws.reset_dimensions()` in `openpyxl(..., read_only=True)` for giant files, streams exports with `Workbook(write_only=True)`, and chunks huge CSVs to prevent Out-Of-Memory crashes.
6. **CSV Formula-Injection Guard:** Neutralizes dangerous leading `=`, `+`, `@`, and non-numeric `-` cells on non-phone columns using `'`, while keeping phone numbers cleanly uncorrupted.
7. **Contact Quarantine Log (`_needs_review.csv`):** Routes unreachable, junk, divider (`Total` / `Subtotal`), and fake contact rows into `_needs_review.csv` with explicit `exclusion_reason` and `raw_snapshot` traceability instead of silent deletion.
8. **Universal Strict `TRIM`:** Removes leading/trailing/repeated spaces, tabs, carriage returns, newlines, NBSP (`\xa0`), BOM (`\ufeff`), and zero-width BiDi marks across **100% of output cells**.
9. **Protected Country Calling Codes:** Strips junk prefixes (`p:`, `+`) and normalizes Arabic-Indic numerals (`٠١٢٣٤٥٦٧٨٩` / `۰۱۲۳۴۵۶۷۸۹`) back to ASCII digits while preserving country prefixes (`971`, `20`, `966`, `44`, etc.) via `COUNTRY_CALLING_CODES`.
10. **Smart 3-Slot Phone Assignment:** Dedupes identical tail numbers into primary Phone + WhatsApp + Extra Phone while guarding cells with Excel Text (`@`) format to eliminate `E+11` scientific notation corruption.
11. **Excel Serial Date Standardization:** Converts both integer serial dates (`45292 -> 2024-01-01`) and ISO timestamps into clean native Excel dates formatted `YYYY-MM-DD`.
12. **Excel Error Token Scrubbing & Pipe Dedup:** Wipes `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A`, `#NUM!`, `#NULL!`, `#ERROR!`, and collapses duplicate `"No Answer | No Answer"` notes cleanly.
13. **Arabic Compare-Key Normalization & Typo Flagging:** Dedupes `أ/إ/آ` -> `ا`, `ى` -> `ي`, `ة` -> `ه` for duplicate matching without mutating original Arabic spelling, and flags email typos (`gamil.com`, `hotmil.com`).
14. **Multi-Encoding Ingestion & Windows Excel BOM:** Auto-fallback across `utf-8-sig`, `utf-8`, `cp1256` (Windows Arabic), and `latin1`, exporting `.csv` in `utf-8-sig` BOM so Excel renders Arabic cleanly without mojibake.

### Newly Added Enterprise Edge Cases in `v1.2.0` (`#15–#21`)
15. **Hidden Rows & Hidden Sheets Filtering:** Automatically skips manually hidden/filtered rows (`ws.row_dimensions[r].hidden`) and hidden internal tabs (`ws.sheet_state != "visible"`) by default (overridable via `--include-hidden-rows` / `--include-hidden-sheets`).
16. **Hyperlink-Embedded Contacts & Cell Comments Rescue:** Extracts real WhatsApp/phone/email targets from Excel hyperlinks (`cell.hyperlink.target` for `wa.me/`, `tel:`, `mailto:`), `=HYPERLINK(...)` formulas, and team notes stored inside `cell.comment.text`, and indexes them for Stage 4 cross-verification.
17. **Macintosh 1904 Date Epoch Detection (`1,462`-Day Shift):** Inspects `wb.epoch` (and legacy `.xls` `datemode`) automatically (`--epoch auto|1900|1904`) so Mac Excel serials (`1904-01-01` base) convert accurately without a 4-year offset error.
18. **Fake & Sequential Placeholder Filtering:** Flags and quarantines synthetic/test junk such as `000000000`, `11111111`, `123456789`, `987654321`, `test@test.com`, and `noemail@noemail.com` under `fake_placeholder_contact`.
19. **Mixed Locale Currency & Budget Normalization:** Unifies European comma decimals (`1.200,50 €`), US thousands formatting (`$1,200.50`), Arabic numerals (`١٬٥٠٠٬٠٠٠`), and `K`/`M` shorthand (`2.5M AED` -> `2500000 AED`, `750k` -> `750000`) while keeping text ranges (`1.5M - 2M`) intact.
20. **Multi-Row Stacked Headers & Banner Row Skipping:** Skips top report banner titles and flattens merged 2-tier group headers (`Category / Subheader`, `header=[0, 1]`) automatically, prioritizing leaf column hints during semantic mapping.
21. **Auto CSV Delimiter Sniffing, Multi-Line Cells & Legacy `.xls`/`.xlsb` Check:** Uses `16 KB` `csv.Sniffer` fallback detection across `,`, `;`, `\t`, and `|` while handling `Alt+Enter` multi-line quoted cells seamlessly, and reads legacy `.xls`/`.xlsb` when `xlrd` or `pyxlsb` is installed (or emits actionable install guidance).

---

## Cross-Platform Compatibility

Built to follow the open [Agent Skills standard](https://agentskills.io/):

| Platform / Agent | Skill Discovery Location | Notes |
| --- | --- | --- |
| **Antigravity IDE** | `.agents/skills/spreadsheet-realign-clean-verify/` | Native `SKILL.md` + optional dual subagents (`Cell_Semantic_Reader`, `Row_Realigner_Verifier`) |
| **OpenAI Codex CLI** | `.agents/skills/spreadsheet-realign-clean-verify/` | Auto-loads `AGENTS.md` + `SKILL.md` |
| **Claude Code** | `.claude/skills/spreadsheet-realign-clean-verify/` | Symlink or copy into `.claude/skills/` |
| **Cursor / Windsurf** | `.cursor/rules/` or `.windsurf/rules/` | Direct reference to `SKILL.md` & `AGENTS.md` |
| **Universal CLI** | Any coding agent | `npx openskills install <repo-url>` |

---

## Repository Structure

```text
spreadsheet-realign-clean-verify/
├── SKILL.md                          # Master 6-stage agent workflow (v1.2.0)
├── AGENTS.md                         # Core non-negotiable engineering rules
├── README.md                         # Full bilingual documentation (EN + AR)
├── LICENSE                           # MIT License
├── .gitignore                        # Strict block against pushing private spreadsheets/PII
├── references/
│   └── patterns-and-rules.md         # Catalog of all 21 corruption & scaling rules
├── scripts/
│   ├── profile_and_extract.py        # Stage 1 profiler, streaming extractor & raw indexer
│   ├── apply_clean_trim_export.py    # Stage 3 deterministic realigner, strict TRIM & exporter
│   └── cross_verify_sources.py       # Stage 4/5 100% phone cross-verifier & audit summary
└── examples/
    ├── dirty_leads_sample.csv        # 100% synthetic sample dirty CSV fixture
    └── expected_clean_sample.csv     # Synthetic clean output fixture
```

---

## Requirements & Quick Start

- **Python 3.10+**
- Required libraries: `pip install pandas openpyxl lxml`
- Optional legacy binary Excel support: `pip install xlrd` (for `.xls`) / `pip install pyxlsb` (for `.xlsb`)

```bash
# 1. Profile, stream-extract, rescue hyperlinks/comments & build raw contact index
python scripts/profile_and_extract.py -i examples/dirty_leads_sample.csv -w ./_clean_work

# 2. Interactive Stage 0 prompt (or pass CLI flags directly): Realign, strict TRIM, deduplicate & export
python scripts/apply_clean_trim_export.py -b ./_clean_work/extracted_base.json -o ./cleaned_output -p sample_clean --interactive

# 3. Verify 100% of phones back to raw sources + output Stage 5 concise numeric report
python scripts/cross_verify_sources.py -c ./cleaned_output/sample_clean.csv -r ./_clean_work/raw_source_index.json -m ./cleaned_output/sample_clean_clean_metrics.json --summary-md
```

---

## Privacy & Security

Never commit client spreadsheets (`.xlsx`, `.xlsm`, `.xls`, `.xlsb`, `.csv`), extracted JSON workdirs (`_clean_work/`), or real PII to GitHub. This repository's `.gitignore` explicitly blocks all spreadsheet binaries, customer CSVs outside `examples/`, and runtime extraction/audit dumps.

---
---

# الدليل الكامل باللغة العربية

## نبذة عن المهارة

مهارة **`spreadsheet-realign-clean-verify`** هي مهارة ذكاء اصطناعي احترافية متوافقة مع كافة برامج ووكلاء البرمجة (`Antigravity`، `Claude Code`، `Codex CLI`، `Cursor`، `Windsurf`) ومصممة لإنقاذ وتنقيح جداول الإكسل وملفات CSV المبعثرة أو ضخمة الحجم (`50 ألف إلى أكثر من مليون صف`). تبدأ المهارة دائماً **بمقابلة استباقية قصيرة للمستخدم (المرحلة 0)** لتحديد صيغة الملفات ومكان الحفظ وطريقة الفصل، ثم تنفّذ عملية التنقيح وإعادة اصطفاف الأعمدة وتطبيق دالة `TRIM` الصارمة على كل خلية، وعزل الصفوف المريبة في ملف `_needs_review.csv`، وتختتم **بفحص ومطابقة 100% من أرقام الهواتف** مع المصادر الأصلية وعرض **تقرير رقمي ختامي مختصر وواضح (المرحلة 5)**.

---

## المخطط المعماري ومراحل العمل

```text
  [المرحلة 0: المقابلة الاستباقية قبل بدء التنفيذ]
   تحديد صيغة الحفظ (Excel / CSV / الاثنين معاً) · مسار الحفظ · طريقة الفصل (حسب العميل أو الوحدة) · تفعيل سجل المراجعة · معالجة الصفوف المخفية
                │
                ▼
  [المرحلة 1: التحليل والفحص والبث الخفيف للبيانات (`scripts/profile_and_extract.py`)]
   قراءة متدفقة قليلة الذاكرة (`read_only` / `chunksize`) · إعادة ضبط أبعاد الورقة `reset_dimensions` · فك الدمج (<10 ميجا)
   استبعاد الصفوف والشيتات المخفية · استخراج أرقام الواتساب والإيميلات من الروابط والتعليقات (`wa.me/`, `tel:`, `cell.comment`)
   اكتشاف نظام تاريخ ماك (1904) أو ويندوز (1900) · كشف فاصل الـ CSV تلقائياً (`,` `;` `\t` `|`) · دمج العناوين المركبة ذات الصفين
                │
        ┌───────┴───────┐
        ▼               ▼
  [المرحلة 2: الوكيل 1] [المرحلة 3: الوكيل 2 (`scripts/apply_clean_trim_export.py`)]
  `Cell_Semantic_Reader` تطبيق `TRIM` الشامل · استبعاد الأرقام والإيميلات الوهمية · توحيد العملات والميزانيات
  يفحص الصفوف الشاذة دلالياً حماية رموز الدول · توحيد الهمزات والياء والتاء للمطابقة فقط · منع حقن معادلات الـ CSV
                         تصدير تكيّفي (`write_only=True` > 50 ألف صف) + `CSV UTF-8 BOM` + ملف `_needs_review.csv`
                │
                ▼
  [المرحلة 4 و 5: المطابقة الشاملة والتقرير الرقمي الختامي (`scripts/cross_verify_sources.py`)]
   مطابقة 100% من كل رقم هاتف أمام المصدر الأصلي `raw_source_index.json` · بوابة خلو العيوب (0 defects) · عرض الملخص الرقمي المختصر
```

---

## قائمة الـ 21 مشكلة التي تحلها المهارة بالكامل

### المشاكل الأساسية لهيكلة الجداول والبيانات الضخمة (`#1–#14`)
1. **المقابلة الاستباقية في المرحلة 0:** تسأل المستخدم أولاً عن صيغة التصدير (`.xlsx` أم `.csv` أم الاثنين معاً)، ومجلد الحفظ (نفس المجلد أو مجلد جديد مع عدم المساس بالملف الأصلي أبداً)، وطريقة الدمج (صف لكل عميل أو صف لكل وحدة)، وهل يُنشأ ملف `_needs_review.csv`.
2. **تداخل نماذج الإعلانات وانزياح الأعمدة في منتصف الشيت:** تصنّف كل خلية حسب معناها الفعلي (بريد داخل خانة الهاتف، هاتف داخل خانة المشروع، تاريخ داخل خانة الحملة) بدلاً من الاعتماد الأعمى على عنوان العمود.
3. **إنقاذ صف العناوين الذي يحتوي على بيانات عميل فعلية (`header-as-data`):** تكتشف الشيتات التي يوضع فيها عميل حقيقي داخل الصف الأول وتنقذه وتُزيل لاحقة `.1` دون فقدان أي بيانات.
4. **معالجة الخلايا المدمجة (`ffill`) وحذف الأعمدة الفارغة الوهمية (+200 عمود):** تفك الخلايا المدمجة وتملأ قيمها لأسفل وتقص الأعمدة الفارغة الممتدة تلقائياً.
5. **المعالجة المتدفقة للملفات الكبيرة (`50 ألف – +1 مليون صف`):** تستدعي `ws.reset_dimensions()` مع وضع `read_only=True` عند القراءة، ووضع `Workbook(write_only=True)` عند التصدير لتفادي استهلاك الذاكرة أو توقف الجهاز.
6. **الحماية من حقن المعادلات الخبيثة في ملفات الـ CSV:** تؤمّن أي نص يبدأ بـ `=` أو `+` أو `@` أو `-` غير الرقمية بوضع علامة `'`، مع إبقاء أرقام الهواتف نظيفة وصالحة للاستخدام.
7. **عزل الصفوف المستبعدة في سجل المراجعة (`_needs_review.csv`):** بدلاً من الحذف الصامت، يتم نقل الصفوف غير القابلة للتواصل أو فواصل المجاميع (`Total` / `Subtotal`) والأرقام الوهمية إلى `_needs_review.csv` مع ذكر سبب الاستبعاد الدقيق ومصدر الصف الأصلي.
8. **تطبيق `TRIM` الشامل والإجباري على 100% من الخلايا:** إزالة المسافات الزائدة في البداية والنهاية وبين الكلمات، وتنظيف رموز `NBSP`، و`BOM`، وفواصل الأسطر، وعلامات اتجاه النص غير المرئية (`BiDi`).
9. **تحويل الأرقام العربية مع حماية أكواد الدول كاملة:** تحويل الأرقام الهندية/الفارسية (`٠١٢٣٤٥٦٧٨٩` / `۰۱۲۳۴۵۶۷۸۹`) إلى إنجليزية وحذف بادئة `p:` مع الحفاظ على رموز الدول (`971`، `20`، `966`، وغيرها) وحماية الرقم التراكمي عند القراءة.
10. **الاصطفاف الذكي للهاتف والواتساب والهاتف الإضافي:** مطابقة أواخر الأرقام لتجميع الهاتف الرئيسي والواتساب والهاتف الإضافي وتخزينها بصيغة نصية (`@`) في الإكسل لمنع ظهورها كأرقام علمية مثل `9.715E+11`.
11. **توحيد صيغ التواريخ والأرقام التسلسلية:** تحويل الأرقام التسلسلية للإكسل (مثل `45292 -> 2024-01-01`) ونصوص الـ ISO المعقدة إلى تواريخ إكسل أصلية بصيغة `YYYY-MM-DD`.
12. **تنظيف أخطاء الإكسل وتكرار الفواصل الرأسية (`|`):** إزالة `#REF!` و`#DIV/0!` و`#VALUE!` و`#NAME?` و`#N/A` و`#ERROR!`، وتنقية تكرار الجمل المنسوخة مثل `"No Answer | No Answer"`.
13. **المطابقة الذكية لأسماء العربية وكشف أخطاء البريد الشائعة:** توحيد (`أ/إ/آ` -> `ا`، `ى` -> `ي`، `ة` -> `ه`) في مفتاح المقارنة لمنع التكرار مع الحفاظ على إملاء الاسم الأصلي، واكتشاف أخطاء نطاقات الإيميل مثل `gamil.com` و`hotmil.com`.
14. **دعم كافة الترميزات ومنع تشويه العربية في إكسل:** قراءة تلقائية بالترتيب (`utf-8-sig` ثم `utf-8` ثم `cp1256` ثم `latin1`) مع تصدير ملفات `.csv` بترميز `UTF-8 with BOM` لتفتح الكلمات العربية سليمة تماماً في برنامج Excel.

### المشاكل الاحترافية المضافة حديثاً في الإصدار `v1.2.0` (`#15–#21`)
15. **استبعاد الصفوف المخفية والشيتات المخفية تلقائياً:** تجاهل الصفوف المخفية أو المستبعدة بالفلتر (`ws.row_dimensions[r].hidden`) والشيتات الداخلية المخفية (`sheet_state != "visible"`) افتراضياً حتى لا تعود بيانات محذوفة أو ملغاة إلى الملف النظيف (مع إمكانية تضمينها اختيارياً عبر `--include-hidden-rows`).
16. **إنقاذ بيانات الاتصال المخفية داخل الروابط (`Hyperlinks`) والتعليقات (`Comments`):** استخراج أرقام الواتساب والهواتف والإيميلات المدفونة خلف نص مثل `"اضغط هنا"` أو معادلة `=HYPERLINK(...)` لروابط (`wa.me/`، `tel:`، `mailto:`)، بالإضافة لاستخراج ملاحظات الموظفين من تعليقات الخلايا (`cell.comment`) وإدراجها ضمن فهرس التحقق.
17. **كشف نظام تاريخ ماك (Mac 1904 Date Epoch) ومنع إزاحة 1,462 يوماً:** فحص خصائص الملف تلقائياً (`wb.epoch` / `xlrd datemode`) لتمييز تاريخ الأساس `1904-01-01` عن `1899-12-30` ومنع تأخير جميع التواريخ بـ **4 سنوات ويوماً واحداً** عند التحويل.
18. **اكتشاف وعزل الأرقام والإيميلات الوهمية والمتسلسلة:** رصد وعزل البيانات الصورية من نماذج التسجيل مثل `000000000`، `11111111`، `123456789`، `987654321`، و`test@test.com` أو `noemail@noemail.com` ونقلها إلى `_needs_review.csv` تحت سبب `fake_placeholder_contact`.
19. **توحيد صيغ العملات المختلفة والميزانيات المختصرة (`K`/`M`):** توحيد الأرقام المكتوبة بالنظام الأوروبي (`1.200,50 €`) أو الأمريكي (`$1,200.50`) والأرقام العربية (`١٬٥٠٠٬٠٠٠`) والاختصارات (`2.5M AED` -> `2500000 AED`، `750k` -> `750000`) دون المساس بنطاقات الأسعار النصية مثل `1.5M - 2M`.
20. **معالجة العناوين المركّبة ذات الصفّين وتخطي صف العنوان الرئيسي (`Banner Row`):** تجاهل صف العنوان التسويقي الفردي بأعلى الورقة، ودمج وتسطيح العناوين الهرمية ذات الصفّين (`الفئة الرئيسية / العمود الفرعي` — مثل `header=[0, 1]`) مع إعطاء الأولوية لاسم العمود الفرعي عند التصنيف.
21. **الكشف الذكي لفاصل الـ `CSV` متعدد الأنماط والأسطر ودعم `.xls`/`.xlsb`:** فحص أول `16 KB` باستخدام `csv.Sniffer` لدعم الفواصل (`,` و`;` و`\t` و`|`) والتعامل السليم مع الخلايا متعددة الأسطر المقتبسة (`Alt+Enter`)، مع دعم مباشر لملفات `.xls` / `.xlsb` عند توفر `xlrd` أو `pyxlsb` أو عرض رسالة تثبيت واضحة.

---

## التوافق مع جميع برامج ووكلاء الذكاء الاصطناعي

تم بناء المهارة وفق المعيار المفتوح [Agent Skills](https://agentskills.io/):

| المنصة أو البرنامج | مسار تثبيت المهارة | ملاحظات التوافق |
| --- | --- | --- |
| **Antigravity IDE** | `.agents/skills/spreadsheet-realign-clean-verify/` | دعم أصلي كامل لـ `SKILL.md` مع الوكيلين الفرعيين (`Cell_Semantic_Reader` و`Row_Realigner_Verifier`) |
| **OpenAI Codex CLI** | `.agents/skills/spreadsheet-realign-clean-verify/` | تحميل تلقائي لملفات `AGENTS.md` و`SKILL.md` |
| **Claude Code** | `.claude/skills/spreadsheet-realign-clean-verify/` | يعمل مباشرة بالنسخ أو الربط الرمزي داخل `.claude/skills/` |
| **Cursor / Windsurf** | `.cursor/rules/` أو `.windsurf/rules/` | تضمين مباشر لقواعد `SKILL.md` و`AGENTS.md` |
| **تثبيت عام (`CLI`)** | أي أداة ذكاء اصطناعي متوافقة | عبر الأمر `npx openskills install <repo-url>` |

---

## هيكل المجلد والملفات

```text
spreadsheet-realign-clean-verify/
├── SKILL.md                          # الدليل الرئيسي ذو الـ 6 مراحل للذكاء الاصطناعي (الإصدار v1.2.0)
├── AGENTS.md                         # القواعد الصارمة الملزمة لأي وكيل يعمل على الملفات
├── README.md                         # التوثيق الكامل ثنائي اللغة (إنجليزي + عربي بالكامل)
├── LICENSE                           # رخصة الاستخدام الحرة MIT
├── .gitignore                        # حماية تمنع رفع أي ملفات عملاء أو بيانات شخصية على GitHub
├── references/
│   └── patterns-and-rules.md         # المرجع الشامل لجميع الـ 21 مشكلة وقواعد معالجتها
├── scripts/
│   ├── profile_and_extract.py        # سكربت المرحلة 1: التحليل والاستخراج المتدفق وفهرسة الأرقام
│   ├── apply_clean_trim_export.py    # سكربت المرحلة 3: إعادة الاصطفاف، TRIM، الفحص والتصدير التكيّفي
│   └── cross_verify_sources.py       # سكربت المرحلة 4 و5: المطابقة بنسبة 100% وإخراج التقرير المختصر
└── examples/
    ├── dirty_leads_sample.csv        # مثال تجريبي عشوائي (خالي تماماً من أي بيانات حقيقية)
    └── expected_clean_sample.csv     # النتيجة المتوقعة بعد التنقيح للمثال التجريبي
```

---

## المتطلبات وطريقة التشغيل السريع

- **Python 3.10 فأحدث**
- المكتبات المطلوبة: `pip install pandas openpyxl lxml`
- لدعم ملفات `.xls` / `.xlsb` القديمة اختيارياً: `pip install xlrd` أو `pip install pyxlsb`

```bash
# 1. تحليل الملف واستخراج البيانات وفهرسة جميع الأرقام والروابط والتعليقات
python scripts/profile_and_extract.py -i examples/dirty_leads_sample.csv -w ./_clean_work

# 2. تشغيل سؤال المرحلة 0 التفاعلي (أو تمرير الصيغة والمسار مباشرة) لإعادة الاصطفاف وتطبيق TRIM والتصدير
python scripts/apply_clean_trim_export.py -b ./_clean_work/extracted_base.json -o ./cleaned_output -p sample_clean --interactive

# 3. التحقق بنسبة 100% من كل رقم هاتف أمام المصدر الأصلي وطباعة التقرير الرقمي النهائي المختصر
python scripts/cross_verify_sources.py -c ./cleaned_output/sample_clean.csv -r ./_clean_work/raw_source_index.json -m ./cleaned_output/sample_clean_clean_metrics.json --summary-md
```

---

## الخصوصية والأمان

لا تقم مطلقاً برفع ملفات العملاء الحقيقية (`.xlsx`، `.xlsm`، `.xls`، `.xlsb`، `.csv`) أو ملفات JSON المؤقتة (`_clean_work/`) على GitHub. يقوم ملف `.gitignore` في هذا المستودع بمنع رفع أي جداول بيانات حقيقية أو ملفات تدقيق خارج أمثلة الاختبار الوهمية داخل مجلد `examples/`.

---

**Author / المطوّر:** Eslam Elgindy (`esgindy`) · [MIT License](LICENSE)

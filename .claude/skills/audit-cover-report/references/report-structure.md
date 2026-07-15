# Cover & report structure (KSK)

Derived from the ground-truth docx in `5_ตัวอย่างไฟล์ผลลัพธ์/`. Two documents per client, each with two variants selected by entity/auditor type.

## Template selection

| CONTEXT signal | kind | cover asset | report asset |
|---|---|---|---|
| juristic_status = liquidated (เลิก/ชำระบัญชี/ร้าง) OR "งบเลิก" in folder | `cpa_liq` | `cover_cpa.docx` | `report_cpa_liquidation.docx` |
| entity_type = บจ. (and auditor CPA) | `cpa` | `cover_cpa.docx` | `report_cpa.docx` |
| entity_type = หจก. OR auditor_type = TA | `ta` | `cover_ta.docx` | `report_ta.docx` |

`render_cover_report.py` auto-selects (liquidation checked first); override with `--kind cpa|cpa_liq|ta`. The choice is consequential (different documents) — confirm it against the client when unsure.

### หน้ารายงาน — งบเลิก (`report_cpa_liquidation.docx`)

CPA report for a dissolved company. Same fixed boilerplate spine as the CPA report but: addressed to `เสนอ กรรมการของ`, adds an emphasis-of-matter paragraph (`ข้อมูลและเหตุการณ์ที่เน้น`) about the dissolution, and uses `ผู้ชำระบัญชี` in place of `ผู้บริหาร`. Tokens: `«COMPANY»` · `«LIQ_DATE»` · `«PERIOD_START»` · `«AUDITOR_NAME»` · `«LICENSE»` · `«SIGN_DATE»` · `«CITY»`.

`«LIQ_DATE»` (dissolution date, Thai long form incl. พ.ศ.) is in no DB — pass `--liquidation-date` (the render fails asking for it if omitted). `«PERIOD_START»` defaults to `1 มกราคม พ.ศ. <liq year>`; override with `--period-start`. The cover reuses `cover_cpa` with its date set to `«LIQ_DATE» (วันที่จดทะเบียนเลิกบริษัท)`.

## ใบปะหน้า (cover) — 1 page, 3 lines

| line | cpa | ta |
|---|---|---|
| company | «COMPANY» | «COMPANY» |
| title | รายงานของผู้สอบบัญชีรับอนุญาตและรายงานทางการเงิน | รายงานการตรวจสอบและรับรองบัญชี / รายงานทางการเงิน |
| date | ณ วันที่ «STMT_DATE» | สำหรับรอบปี สิ้นสุด วันที่ «STMT_DATE» |

## หน้ารายงาน — CPA (`report_cpa.docx`)

ISA-style auditor's report. Body paragraphs are fixed boilerplate, identical across all CPA cases. Tokens:

`«COMPANY»` (×3) · `«STMT_DATE»` (×2) · `«AUDITOR_NAME»` · `«LICENSE»` · `«SIGN_DATE»` · `«CITY»` (sign city, default ขอนแก่น = firm office)

**Optional prior-auditor block** — the "เรื่องอื่นๆ" paragraph appears only when the prior year was audited by someone else. Driven by CONTEXT `has_prior_auditor_note`:
- `true` → block kept, `«PRIOR_AUDITOR_TEXT»` filled (uses `prior_auditor_text` from CONTEXT, else a default sentence with `«PRIOR_STMT_DATE»`).
- absent/`false` → both block paragraphs removed.

A CPA report with ANY residual `«TOKEN»` after render is a bug — the script hard-fails.

## หน้ารายงาน — TA (`report_ta.docx`)

RD-style "รายงานการตรวจสอบและรับรองบัญชี" form (different document, not a title swap). Form placeholders like "(อธิบายข้อยกเว้นที่สำคัญ ถ้ามี)" are literal and stay. Tokens:

`«COMPANY»` · `«STMT_DATE»` · `«CUR_YEAR»` · `«PRIOR_YEAR»` · `«AUDITOR_NAME»` · `«LICENSE»` · `«NATIONAL_ID»` · `«OFFICE»` · `«SIGN_DATE»` · `«PRIOR_SIGN_DATE»`

`«NATIONAL_ID»` is read from Database ผู้สอบ by auditor name (rendered double-spaced per the form). `«OFFICE»` = the auditor's office in Database ผู้สอบ, or the firm office (`9/38 ถ.กลางเมือง ต.เมืองเก่า อ.เมือง จ.ขอนแก่น`) when blank; override with `--office`. `«PRIOR_SIGN_DATE»` (prior-year report date) is in no DB — a per-client historical fact; pass `--prior-sign-date`, and if omitted the token is left in place and reported as a warning so the agent asks the user. The script never renders a blank in place of a missing value.

## Dates and sign city

All dates trace to Database งาน: `รอบปีบัญชี` → statement date (`«STMT_DATE»`, converted to Thai long form), `วันที่หน้ารายงาน` → `«SIGN_DATE»` (already Thai long form). Prior statement date is derived (year − 1). Sign city (`«CITY»`) is constant `ขอนแก่น` (the KSK office) across all 15 sample cases; override via a `sign_city` row in CONTEXT if ever needed.

## Templates are generated, not hand-authored

`build_templates.py` (DEV, run once) reads the ground-truth docx, collapses each variable paragraph to a single run, and replaces concrete values with tokens — then asserts no source client name/number leaked. Runs are heavily fragmented (7–20 runs/paragraph), which is why collapsing is required. The generated `assets/*.docx` are client-neutral boilerplate and are committed; regenerate them only when the firm's standard wording changes.

## Whitespace

Ground-truth docs have inconsistent whitespace (stray double spaces around company names and "ณ"). `build_templates.py` collapses 2+ spaces to one in the templates, and `clean()` does the same to values pulled from the DB, so rendered output is uniformly single-spaced — cleaner than some ground-truth docs but content-identical. The national ID is the one intentional exception (digits are double-spaced to match the form).

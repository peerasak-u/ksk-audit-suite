---
name: audit-cover-report
description: Generate the KSK cover page (ใบปะหน้างบการเงิน) and auditor's report (หน้ารายงาน) .docx from a client CONTEXT.md. Use when asked to "make the cover page", "generate the audit report page", "render หน้ารายงาน / ใบปะหน้า", or after audit-ingest has produced a CONTEXT.md. Fills tokenized .docx templates deterministically via a CLI; picks the CPA or TA template by entity type. Does not write free-form report text.
allowed-tools: Bash, Read
---

Render ใบปะหน้า + หน้ารายงาน by filling locked `.docx` templates from `CONTEXT.md`. Deterministic token substitution only — never draft or edit report prose by hand.

## Hard rules

1. **Only a script writes the .docx.** Do not open Word/docx and type report text. Run the CLI.
2. **Input is CONTEXT.md**, produced by `audit-ingest`. If it is missing, run audit-ingest first.
3. **Never invent values.** The script refuses to render if a required field is still `⚠`/missing, and flags any TA field it cannot fill. Report those — do not fill them by guessing.
4. **Write only under `6_ผลจากสกิล/`.** Never touch folders `4_` or `5_`.
5. If the CONTEXT verdict is `REVIEW` / `AMBIGUOUS` / `JOB NOT IN DB`, stop and tell the user to resolve it first.

## Step 1: Render

From the repository root:

```bash
uv run ${CLAUDE_SKILL_ROOT}/scripts/render_cover_report.py "PATH/TO/6_ผลจากสกิล/CLIENT/CONTEXT.md"
```

The script parses the CONTEXT Client Profile table, auto-selects the CPA or TA template by entity type, fills the tokens, and writes `ใบปะหน้างบการเงิน.docx` + `2 หน้ารายงาน ....docx` into the client's `WP/` dir. Add `--kind cpa|ta` only to override the auto-selection.

Read the JSON result:
- `ok:false` → report the error (usually an unresolved required field) and stop.
- `ok:false` with a งบเลิก message → the client is dissolved and `--liquidation-date` is required. **Ask the user** for the dissolution date (Thai long form, e.g. `26 ธันวาคม พ.ศ. 2568`) and re-run with `--liquidation-date "..."` (add `--period-start` only if the period did not start 1 มกราคม).
- `warnings` non-empty → a token could not be resolved from any DB. **Ask the user for that value, then re-run with the matching flag** — never invent it:
  - `«PRIOR_SIGN_DATE»` (TA prior-year report date, not in any DB) → ask, then `--prior-sign-date "28 พฤษภาคม 2568"`.
  - `«OFFICE»` (only if both the auditor's Database ผู้สอบ office and the firm default are unusable) → ask, then `--office "..."`.

Liquidation (งบเลิก) is auto-detected from `juristic_status` or a "งบเลิก" folder name; it uses `report_cpa_liquidation.docx` and addresses กรรมการ / ผู้ชำระบัญชี.

Dates come from Database งาน (รอบปีบัญชี → statement date, วันที่หน้ารายงาน → sign date). Sign city is the firm office (ขอนแก่น), constant across all cases.

For structure, tokens, the CPA prior-auditor block, and the selection rule, read `${CLAUDE_SKILL_ROOT}/references/report-structure.md`.

## Step 2: Report

Tell the user, in a few lines:
- which `kind` (cpa/ta) was chosen and the two output paths
- any `warnings` (unresolved tokens to fill by hand)
- that the documents are ready for human review against the firm's standard

## Regenerating templates (rare)

If the firm's standard wording changes, regenerate the tokenized assets from ground truth (needs `5_ตัวอย่างไฟล์ผลลัพธ์/` present):

```bash
uv run ${CLAUDE_SKILL_ROOT}/scripts/build_templates.py
```

This is a dev-only step, not part of a normal run.

## Done when

- `ใบปะหน้างบการเงิน.docx` and `2 หน้ารายงาน ....docx` exist in the client's `WP/`.
- No required field was missing (script returned `ok:true`).
- Any TA data-gap warnings were relayed to the user.

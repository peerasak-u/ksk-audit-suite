---
name: audit-ingest
description: Ingest a KSK audit client folder into a single CONTEXT.md working paper. Use when asked to "ingest a client", "scaffold a job", "prepare CONTEXT", "start an audit job", or when pointed at a client folder under 4_ตัวอย่างไฟล์ลูกค้า. Joins the three root Database files by job number, extracts archives, reads scanned บอจ.5 by vision, and writes CONTEXT.md plus the WP skeleton under 6_ผลจากสกิล. Does NOT generate any output document.
allowed-tools: Bash, Read, Edit, Glob
---

Ingest one client folder into a verified `CONTEXT.md`. Extraction only — never generate report/planning/CIT files here.

## Hard rules

1. **Never invent a value.** If a field has no source, leave `—` and keep `⚠`. No guessing company facts, directors, dates, fees.
2. **Never edit DB-sourced rows** (the `{{...}}` values the script wrote). Only replace `⟨FILL: ...⟩` markers.
3. **Do not touch `4_ตัวอย่างไฟล์ลูกค้า/` or `5_ตัวอย่างไฟล์ผลลัพธ์/`.** Write only under `6_ผลจากสกิล/`.
4. **Stop at CONTEXT.md.** Downstream skills run later, only after a human clears the ⚠ rows.
5. If validation verdict is `REVIEW` or `JOB NOT IN DB`, **report it and stop** — do not fill markers or proceed.

## Step 1: Run the DB-join

From the repository root:

```bash
uv run ${CLAUDE_SKILL_ROOT}/scripts/join_databases.py "PATH_TO_CLIENT_FOLDER_UNDER_4"
```

The script parses the job number from the folder name, joins all three Database files, **looks up the DBD OpenAPI by tax id** for the authoritative full legal name (plus EN name, type, status, registered capital, incorporation date, business type), writes `6_ผลจากสกิล/FOLDER/CONTEXT.md` from the pinned template (`${CLAUDE_SKILL_ROOT}/references/context-template.md`) with every value filled, creates `WP/`, and prints JSON.

The DBD lookup needs network and retries on throttling; if it cannot resolve the entity, those rows are left `⚠` with source `DBD (not found)` and `company_legal_name` falls back to the (possibly abbreviated) DB name.

Read the JSON. If `ok` is false, report the error and stop. Check `validation_verdict`:
- `PASS` → continue to Step 2.
- `REVIEW` / `JOB NOT IN DB` → report to the user and **stop** (Hard rule 5).

## Step 2: Extract archives, complete the inventory

For every archive (`.rar`/`.zip`/`.7z`) in the client folder, extract to a temp dir and compare against loose files:

```bash
tar -xf "<archive>" -C "<tempdir>"
```

Compare basenames. Files present ONLY inside the archive are new — append them to the Document Inventory table in CONTEXT.md, tagged `(new from archive)`. Most archives are duplicates (0 new); some carry the real financials (e.g. a zip with งบทดลอง / กระดาษทำการ / GLTRIAL). Never skip extraction.

## Step 3: Fill vision-extracted fields from บอจ.5

Read the บอจ.5 PDF with vision (it is a scan — `pypdf` returns empty text). Replace these markers in CONTEXT.md, verbatim from the document:

| marker | value |
|---|---|
| `registered_capital` | "ทุนจดทะเบียน" amount, digits only |
| `address` | shareholder #1 full address line |
| `director_1` | shareholder #1 name — keep `⚠` (role unconfirmed without หนังสือรับรอง) |
| Shareholders table | one row per shareholder: `# \| ชื่อ \| สัญชาติ \| หุ้น` |

If บอจ.5 is missing or unreadable, write `— ไม่พบ บอจ.5` and leave the rows `⚠`. Set each filled row's status to `✔` only when the value came directly from the document; otherwise leave `⚠`.

Leave `authorized_signatory`, `incorporation_date`, `business_type`, `director_1` role, `audit_team`, and `materiality_base` as `⚠` — they need หนังสือรับรอง / DBD / financials / human input, not available here. (`has_prior_auditor_note` is handled in Step 4.)

## Step 4: Detect the prior-year auditor (เรื่องอื่นๆ block)

The CPA report carries an optional เรื่องอื่นๆ paragraph when the **prior** year was signed by a **different** auditor than this year's (a first-year engagement for the signing auditor). A different individual counts even within KSK. The two values it needs live in the client's own prior-year report — do NOT ask the user unless that document is absent.

1. **Find the prior-year report** in the client folder (search recursively, prior year = period year − 1):
   - a subfolder like `งบ 25<prior>/` holding `หน้ารายงาน…signed.pdf`, or
   - a prior-year financial-statements PDF (`งบการเงินปี 25<prior> …`) — the report is on its first pages.
   If none exists, set `has_prior_auditor_note` to `false`, add a Judgement note "no prior-year report in folder — confirm with human", and skip to Step 5.
2. **Read it with vision** (these are scans — `pypdf` returns empty). From the signature block read the prior auditor's **name + license number**, the **opinion phrase**, and the **sign date**.
3. **Decide the flag.** If the prior signer's name/license ≠ this year's `auditor_name`/`auditor_license`, set `has_prior_auditor_note` = `true`; otherwise `false` (continuing auditor — no block).
4. **When `true`, fill the two atoms** in the section-4 table (cover-report composes the sentence — do not write the full paragraph yourself):
   - `prior_auditor_opinion` — the opinion phrase, e.g. `อย่างไม่มีเงื่อนไข` (unqualified). Set `✔`.
   - `prior_auditor_sign_date` — the report date in Thai long form, e.g. `25 เมษายน 2568`, verbatim from the document. Set `✔`.
   Leave both `—`/`⚠` when the flag is `false`.

## Step 5: Report

Tell the user, in 3–5 lines:
- CONTEXT.md path and validation verdict
- how many archive files were new
- which `⚠` rows remain and what each needs (หนังสือรับรอง, DBD, human, financials)

Do not proceed to any downstream skill.

## Done when

- `6_ผลจากสกิล/<folder>/CONTEXT.md` exists, every `{{...}}` resolved, DB rows untouched.
- Archives extracted and inventory reflects new files.
- บอจ.5 fields filled or explicitly marked absent.
- `has_prior_auditor_note` set from the prior-year report (with its two atoms when `true`), or `false` with a note when no prior-year report is in the folder.
- Remaining `⚠` rows listed for the human.

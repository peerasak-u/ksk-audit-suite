---
name: audit-planning
description: Generate the "1 Planning" audit workpaper .xlsx for a client from CONTEXT.md. Use when asked to "make the planning file", "generate the audit plan", "fill Planning.xlsx", or after audit-ingest has produced a CONTEXT.md. Fills sheet ข้อมูลลูกค้า deterministically via a CLI (the other 9 sheets cascade off it through formulas already in the firm's template) and selects one of 3 known boilerplate variants for the risk-assessment sheets. Does not write free-form audit narrative.
allowed-tools: Bash, Read
---

Render `1 Planning <company>.xlsx` by filling one sheet (`ข้อมูลลูกค้า`) of a locked,
client-neutral template from `CONTEXT.md`. Deterministic field-fill only — never type
values directly into the spreadsheet, and never touch any sheet/cell other than
`ข้อมูลลูกค้า`, `301!A28:E28`, and `203 TB (2)!`rows 5-10.

## Hard rules

1. **Only the script writes the .xlsx.** Do not open the file and type values by hand.
2. **Input is CONTEXT.md**, produced by `audit-ingest`. If it is missing, run audit-ingest first.
3. **Never invent values.** The script refuses to render if `company_name` / `tax_id` /
   `period_end` / `auditor_name` / `auditor_license` / `audit_fee` is still `⚠` in
   CONTEXT. Every other unresolved field (director, audit team, engagement dates,
   hours, materiality base) is left blank in the output and reported as a warning —
   **ask the user for it**, don't guess. See `references/planning-structure.md` for
   the full field list and which CLI flag fills which row.
3b. **`--business-type` is a judgment call, not something to guess.** Unless CONTEXT
   shows a clear liquidation signal (auto-detected as `dormant`), the script *refuses*
   to render without `--business-type` explicitly given — confirmed against all 16
   ground-truth files that business_type/revenue/juristic_status text has no reliable
   correlation with which of `default`/`construction`/`dormant` the firm actually used.
   When the script fails on this, **stop and ask the user** — show them CONTEXT's
   `business_type` text and `revenue`, and let them pick. Do not pattern-match keywords
   yourself and pass a guess.
4. **Write only under `6_ผลจากสกิล/`.** Never touch folders `4_` or `5_` — folder 5 is
   ground truth for QA, read-only, never edit it.
5. **Sheets 301 / 203 TB (2) are boilerplate selection, not narrative writing.** Ground
   truth shows the firm reuses ~3 fixed text blocks across unrelated industries almost
   verbatim — this skill replicates that behavior via `--business-type`, it does not
   draft bespoke risk analysis. See the "boilerplate, not bespoke narrative" section of
   `references/planning-structure.md` before assuming a client needs custom prose.
6. **งบเลิก (liquidation) clients are not fully supported yet.** The script fills the
   correct cells, but sheets 103/601/608 need liquidation-specific wording
   ("ผู้ชำระบัญชี" not "กรรมการ") that the current template doesn't have — flag this to
   the user rather than presenting a liquidation Planning.xlsx as final. See the "Known
   gap" section of `references/planning-structure.md`.

## Step 1: Gather the fields not already in CONTEXT

CONTEXT.md (from audit-ingest) does not carry director names beyond `director_1`,
audit team assignment, engagement scheduling dates/hours, or materiality figures —
these are genuinely not in any DB yet. Before rendering, check CONTEXT's `⚠` rows and
**ask the user** for whatever the client folder's own documents (หนังสือรับรอง, engagement
letter, prior workpapers) don't already answer. Do not block on this — the script will
still render with sensible warnings for anything left unresolved.

## Step 2: Render

From the repository root:

```bash
uv run ${CLAUDE_SKILL_ROOT}/scripts/render_planning.py "PATH/TO/6_ผลจากสกิล/CLIENT/CONTEXT.md" [options]
```

Common options (full list and defaults in `references/planning-structure.md`):

- `--business-type default|construction|dormant` — picks the 301/203 boilerplate
  variant. Auto-detects `dormant` from a liquidation signal in CONTEXT; otherwise
  **required** — the script fails and tells you to ask the user rather than guessing.
  There is no reliable signal (from business_type text, revenue, or anything else in
  CONTEXT) that predicts `default` vs `construction`, so don't infer it yourself —
  surface CONTEXT's business_type/revenue to the user and let them decide.
- `--director NAME` (repeatable, up to 3), `--director-authority "..."`
- `--audit-team NAME` (repeatable) — the lead auditor is NOT included in this list
- `--audit-start-date`, `--audit-hours`, `--rep-letter-date`, `--audit-expense` — no DB
  source for these, always ask the user (audit_expense auto-estimates at 30% of
  audit_fee if omitted, flagged for review)
- `--materiality-basis "สินทรัพย์รวม"|"รายได้รวม"`, `--materiality-reason`,
  `--materiality-base` — materiality figures need งบการเงิน (a later-phase skill);
  leave unset until then
- `--business-desc "..."` — a tightened one-line ประเภทธุรกิจ phrase; without it, the
  raw (often verbose) DBD business_type text from CONTEXT is used verbatim
- `--biz-environment`, `--tb-note`, `--prior-adjustment-note` — override the
  business_type boilerplate text directly when a client doesn't fit any of the 3
  variants cleanly (this happens — see the "outlier" note in the reference doc)

Read the JSON result:
- `ok:false` → report the error (an unresolved required field) and stop.
- `warnings` non-empty → tell the user exactly what was left blank or auto-estimated
  and which flag resolves it. Never re-run guessing a value — ask first.

## Step 3: Report

Tell the user, in a few lines:
- the output path and which `business_type` variant was selected
- any `warnings` (fields left blank, or auto-estimated ประมาณการค่าใช้จ่าย) — these need
  human input or confirmation before the plan is final
- that only `ข้อมูลลูกค้า` + the two boilerplate rows were touched; the other 9 sheets
  cascade automatically via the firm's own existing formulas — do not open the file and
  "fix" anything there

## Regenerating the template (rare)

If the firm's own Planning.xlsx master template changes shape, regenerate the
client-neutral template from ground truth (needs `5_ตัวอย่างไฟล์ผลลัพธ์/` present):

```bash
uv run ${CLAUDE_SKILL_ROOT}/scripts/build_template.py
```

Dev-only, not part of a normal run. Re-run `references/planning-structure.md`'s
analysis (re-open several ground-truth files and diff) before trusting a rebuilt
template — see that file for the row map and boilerplate bank that must stay in sync.

## Done when

- `1 Planning <company> <period>.xlsx` exists in the client's `WP/`.
- No required field was missing (script returned `ok:true`).
- All warnings were relayed to the user with the flag that resolves each one.
- If the client is งบเลิก, the user was told sheets 103/601/608 need manual review.

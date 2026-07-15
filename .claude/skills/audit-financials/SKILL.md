---
name: audit-financials
description: Finalize a client's financial-statement working paper after the auditor has adjusted it — run the QA gate and expand the note detail. Use when asked to "finalize the งบการเงิน", "validate the financial statements", "QA the workpaper", "check the งบ balances", or "แตกหมายเหตุ / expand note detail", after audit-workpaper (skill 5) scaffolded the workbook and a human has adjusted the TB, classified every account in Mapping!H, and posted the AJEs. Validates that the งบ is submission-ready (balance sheet balances, no #REF!, all accounts classified, company/year match, profit ties) and expands the one-row-per-caption notes into itemized per-account detail. Keeps the workbook a live, editable Excel — it never freezes formulas to values, never renders a PDF, and never invents numbers. This is skill 5.5 of the pipeline. v1 covers บจ. going-concern only.
allowed-tools: Bash, Read
---

Finalize the financial-statement working paper that `audit-workpaper` (skill 5) scaffolded,
**after** a human has done the audit judgment: adjusted the trial balance, classified every
account in `Mapping!H` from the dropdown, and posted the adjusting entries. Two jobs:

1. **QA gate** — check the finished งบ is submission-ready and emit an exception report.
2. **Note-detail expansion** — turn the one-summary-row-per-caption notes into the itemized
   per-account breakdown a real งบ needs.

The workbook stays a **live, editable Excel** throughout. This skill does NOT freeze formulas
to values, render a PDF, or produce the cover/report .docx (that is `audit-cover-report`,
skill 2). The shared structure it reads is `docs/financials-contract.md`.

## Hard rules

1. **Only the script touches the workbook.** Do not edit cells by hand in a viewer.
2. **Never invent numbers or classifications.** The QA gate only reads; the expansion only
   re-arranges existing live formulas. Anything wrong is reported for the human to fix.
3. **Never modify the auditor's working file in place.** Note-detail expansion writes a NEW
   copy (`… (final).xlsx`); the input WP is left untouched.
4. **The input must be an Excel-saved workbook.** openpyxl cannot evaluate formulas, so the WP
   must have been opened once in Excel/Google Sheets/LibreOffice and **saved** (which caches
   the computed values) before it can be validated. The script detects an unsaved file (all
   formula cells read empty) and refuses with that instruction — this is correct: a WP never
   opened in a spreadsheet app is by definition not finished.
5. **v1 = บจ. going-concern only.** หจก. / งบเลิก / ปีแรก variants (contract §6) not covered.

## Step 1: Run the QA gate first

From the repository root:

```bash
uv run ${CLAUDE_SKILL_ROOT}/scripts/finalize_financials.py \
    "PATH/TO/6_ผลจากสกิล/<client>/CONTEXT.md" [--wp PATH] --validate-only
```

- `--wp PATH` — the WP workbook (default: newest `4 งบการเงิน*.xlsx` in `<ctx>/WP`).
- Read the JSON `validation` block:
  - `errors` (hard — must be fixed): unsaved file, `#REF!`/error cells, balance sheet does
    not balance, unclassified balance-bearing accounts, captions outside the NPAE vocabulary,
    profit not tying to the tax sheet.
  - `warnings` (soft — relay, don't block): CIT50 net-profit reconciliation is a manual check;
    stale company name / year labels in a sheet.
  - `checks` — the numbers behind the checks (BS differences, PBT, net profit, counts).
- **If there are errors, stop and report them.** The human fixes the WP in Excel, saves, and
  you re-run. Do not proceed to expansion on a WP that fails the gate.

## Step 2: Finalize (expand note detail) once the gate passes

```bash
uv run ${CLAUDE_SKILL_ROOT}/scripts/finalize_financials.py \
    "PATH/TO/6_ผลจากสกิล/<client>/CONTEXT.md" [--wp PATH] [--out PATH]
```

Runs the same gate, then — only if it passes — writes a finalized copy with itemized notes.
- `--out PATH` — output path (default: `<wp> (final).xlsx`).
- Read the JSON: `output` (the new file), `notes_expanded` (how many captions got itemized).

The finalized copy: every statement line is a self-contained live `SUMIF`-by-caption; each
note (4 onward) lists one live row per contributing account (linked to its `Mapping` row) with
a subtotal. All formulas — nothing is frozen. Because openpyxl writes formulas without cached
values, the human opens the copy in Excel once to recompute (the script says so in a warning).

## Step 3: Report

Tell the user, in a few lines:
- the gate result (passed / the exact exceptions to fix)
- for a finalize run: the output path and how many notes were itemized
- the manual CIT50 net-profit reconciliation still to do
- any stale-sheet / year warnings
- that the file is still a live editable Excel (open once in Excel to recompute)

## What this skill checks (QA gate)

| check | severity | meaning |
|---|---|---|
| workbook has cached values | error | else it was never saved in a spreadsheet app |
| no `#REF!` / error cells | error | refuse to finalize a งบ containing errors |
| balance sheet balances (CY & PY) | error | `รวมสินทรัพย์` = `รวมหนี้สินและส่วนของผู้ถือหุ้น` |
| tie-out row = 0 | error | the scaffold's built-in balance check |
| every balance-bearing account classified | error | `Mapping!H` filled for accounts with a balance |
| captions in NPAE vocabulary | error | `Mapping!H` values are valid captions |
| profit ties to tax sheet | error | IS profit-before-tax = `ภาษีเงินได้` starting profit |
| CIT50 net-profit reconciliation | warning | manual — compare `FS_NET_PROFIT_CY` to ภ.ง.ด.50 |
| company name / year match CONTEXT | warning | stale-sheet guard (contract §8) |

Details of the formula chain, defined names, and the note-expansion design are in
`references/finalization.md` and `docs/financials-contract.md`.

## Done when

- The QA gate reports `ok:true` (no errors), and all warnings were relayed.
- For a finalize run: `… (final).xlsx` exists with `notes_expanded` > 0, and the user was told
  to open it once in Excel to recompute and to do the CIT50 reconciliation by hand.

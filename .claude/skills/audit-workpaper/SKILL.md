---
name: audit-workpaper
description: Scaffold the audit working-paper workbook (4 งบการเงิน <client>.xlsx) for a client from CONTEXT.md. Use when asked to "scaffold the workpaper", "set up the financial workpaper", "prepare the งบการเงิน template/structure", or after audit-ingest. Builds a formula-linked งบการเงิน template + Mapping (account→caption) + adjusted-TB layout + empty AJE grid + tax shell, so the statements recompute themselves once the auditor classifies accounts and adjusts the TB. Never invents numbers or account classifications. This is skill 5 of the pipeline; audit-financials (skill 5.5) later finalizes the งบ. v1 covers บจ. going-concern only.
allowed-tools: Bash, Read
---

Prepare the **structure** of a client's financial working-paper workbook before the numbers
and judgment exist. The workbook is wired so that once the auditor classifies each account in
`Mapping!H` (from a controlled dropdown) and adjusts the trial balance, every statement line
and subtotal recomputes itself — "approach A" in `docs/financials-contract.md`.

This skill does NOT produce finished financial statements — that is `audit-financials`
(skill 5.5), which runs after a human has done the audit judgment. This skill only lays the
scaffold and the live formula chain.

## Hard rules

1. **Only the script writes the .xlsx.** Do not build the workbook by hand in a viewer.
2. **Input is CONTEXT.md** (from `audit-ingest`) plus, optionally, the client's raw
   **TB/GL export** (`--client-tb`). If CONTEXT is missing, run `audit-ingest` first.
3. **Never invent numbers or classifications.** The script pastes the client's TB balances
   as-is and leaves `Mapping!H` (which caption each account belongs to) BLANK for the human —
   it is a judgment call. Adjusting entries are 100% human. Anything not derivable stays a ⚠.
4. **Write only under `6_ผลจากสกิล/`.** Never touch folders `4_` or `5_`.
5. **The shared structure is a contract.** Sheet names, column layouts, defined names, and the
   caption vocabulary are specified in `docs/financials-contract.md` and read back by
   `audit-financials`. Do not diverge from it — change the contract first if needed.
6. **v1 = บจ. going-concern only.** หจก. (partnership), งบเลิก (liquidation), and ปีแรก
   (first year) need their own variants (contract §6); the script warns if it detects หจก.

## Step 1: Locate the client TB (optional but recommended)

Look in the client's folder-4 documents (listed in CONTEXT §3) for a trial-balance / GL
export `.xlsx` — names vary (`งบทดลอง`, `GLTRIAL`, `TB…`, `Mapping ผังบัญชี…`). Passing it
pre-loads every account into `TB` and `Mapping` so the human only classifies, not retypes.
Without it, the script still produces the full empty structure to paste into by hand.

## Step 2: Scaffold

From the repository root:

```bash
uv run ${CLAUDE_SKILL_ROOT}/scripts/scaffold_workpaper.py \
    "PATH/TO/6_ผลจากสกิล/<client>/CONTEXT.md" [--client-tb PATH [--tb-sheet NAME]]
```

- `--client-tb PATH` — the client's TB/GL export to pre-load accounts (best-effort parse;
  the script warns to verify the Dr/Cr columns landed correctly since client exports vary).
- `--tb-sheet NAME` — which sheet inside `--client-tb` (default: first sheet).

Read the JSON result:
- `ok:false` → report the error and stop.
- `accounts_imported` → how many TB rows were loaded (0 = empty layout to fill by hand).
- `warnings` → relay each one: unclassified Mapping, no-TB, entity-variant, import-verify.

## Step 3: Report

Tell the user, in a few lines:
- the output path
- how many accounts were imported (or that TB/Mapping are empty to fill by hand)
- the key next human action: **open in Excel, classify every account in `Mapping!H` from the
  dropdown** — until then the งบ reads zero (this is expected, not a bug)
- that this is the scaffold only; the finished งบ comes from `audit-financials` (skill 5.5)
  after the audit judgment is done
- any variant warning (หจก./งบเลิก/ปีแรก not covered by v1)

## What the workbook contains

| sheet | what the human does |
|---|---|
| `งบการเงิน` | nothing — auto-computes (BS + IS + equity + note-detail via formulas) |
| `TB` | paste/adjust the trial balance; post AJEs into cols E/F |
| `Mapping` | **classify every account in col H** from the dropdown (the one judgment cell) |
| `ปรับปรุง` | enter adjusting entries (never auto-generated) |
| `ภาษีเงินได้` | fill the ⚠ add-back / exemption / credit cells; the rest computes |
| `0000`, `_captions` | hidden — placeholder and the controlled caption vocabulary |

Details of the formula chain, defined names, caption vocabulary, and variants are in
`references/workpaper-structure.md` and `docs/financials-contract.md`.

## Done when

- `4 งบการเงิน <client> <period>.xlsx` exists in the client's `WP/` under `6_ผลจากสกิล/`.
- The script returned `ok:true`.
- All warnings were relayed, especially the `Mapping!H` classification step.

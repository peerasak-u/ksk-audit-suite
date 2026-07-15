---
name: audit-cit50
description: Fill the CIT50 (ภ.ง.ด.50, corporate income tax return) AcroForm PDF for a client from CONTEXT.md. Use when asked to "fill CIT50", "generate ภงด.50", "fill the tax return", or after audit-ingest/audit-planning have produced CONTEXT.md and the client's own director/team data. Fills the 925-field government AcroForm deterministically via pypdf. Tax-computation numbers are extracted from the client's 4 งบการเงิน*.xlsx (the ภาษีเงินได้/คำนวณภาษี sheet, or the accounting net profit as fallback) — never computed by an agent guessing.
allowed-tools: Bash, Read
---

Fill `3 CIT50 <company>.pdf` by writing values into a locked, client-neutral AcroForm
template from `CONTEXT.md` plus numbers extracted from the client's own financial
workpaper. Deterministic field-fill only — never open the PDF and type into it by hand.

## Hard rules

1. **Only the script writes the .pdf.** Do not open the PDF and fill it in a viewer.
2. **Input is CONTEXT.md** (from `audit-ingest`) plus **numbers extracted from
   `4 งบการเงิน*.xlsx`** in the same client's `WP/` folder. If CONTEXT is missing, run
   `audit-ingest` first.
3. **Never compute tax.** Revenue, COGS, net profit, tax base, tax computed, and
   withholding/advance-tax credits are read out of the client's own workpaper (the
   `ภาษีเงินได้`/`คำนวณภาษี` sheet if it exists, else the accounting net profit with
   the explicit assumption of zero adjustments — see
   `references/cit50-field-map.md`) and passed to the script via CLI flags. The
   script is pure passthrough for these numbers — it does not apply tax rates or
   exemption brackets itself. **Before extracting, confirm the sheet's own date/year
   header matches CONTEXT's `period_end`** — a stale prior-year sheet produces a
   silently wrong filing (see the "check which year's column" caveat in the reference doc).
4. **Never invent values.** ISIC code, attachment page counts, bookkeeping fee, and
   the auditor's office tax ID (`aud.3`) are not in any DB — leave them blank and
   report the warning, ask the user, don't guess.
5. **Write only under `6_ผลจากสกิล/`.** Never touch folders `4_` or `5_`.
6. **Director/team data is shared with `audit-planning`.** Gather it once (director
   name, title) and reuse for both skills instead of asking twice.

## Step 1: Locate the tax numbers

Open the client's `4 งบการเงิน*.xlsx` (in folder 5 for reference cases, or the
client's own `6_ผลจากสกิล/<client>/WP/` once `audit-financials` produces one).
Look for a sheet named `ภาษีเงินได้` or `คำนวณภาษี` — 7 of 8 ground-truth cases have
one, with a fixed structure (accounting profit → add disallowed expenses → less loss
carryforward → less SME exemption → × rate → less WHT/ภงด.51 credits → net payable).
If it doesn't exist, use the `งบการเงิน` sheet's กำไร(ขาดทุน)สุทธิ line directly and
assume zero adjustments (documented as an explicit warning, not silent).

## Step 2: Render

From the repository root:

```bash
uv run ${CLAUDE_SKILL_ROOT}/scripts/render_cit50.py "PATH/TO/6_ผลจากสกิล/CLIENT/CONTEXT.md" [options]
```

Key options (full list and field mapping in `references/cit50-field-map.md`):

- `--director NAME` (reuse the same value given to `audit-planning`), `--director-title`
  (defaults to กรรมการ, or ผู้ชำระบัญชี if the client is liquidating)
- `--filing-date dd/mm/bbbb` — defaults to CONTEXT `sign_date` with a warning; verify
  against the actual filing date
- `--revenue N [--revenue-is-principal]` — classify as line 1 (principal business
  revenue) or line 4 (other income) based on whether it matches the client's
  registered `business_type`; default is "other income" (line 4) which also mirrors
  into the P4.5 detail schedule
- `--cogs N`, `--other-income N` (interest etc., additive on top of `--revenue`),
  `--sga-total N`
- `--net-profit-accounting N`, `--net-profit-tax N` (defaults to the accounting
  figure — Path B fallback), `--tax-base N`, `--tax-computed N`, `--wht-credit N`,
  `--pnd51-credit N`
- `--cash N --other-current-assets N --net-fixed-assets N --trade-payables N
  --other-current-liabilities N --retained-earnings N --total-assets N` — balance
  sheet section; `--total-assets` is the same figure as `audit-planning`'s
  `--materiality-base`, reuse it
- `--isic CODE`, `--bookkeeping-fee N`, `--office-tax-id ID` — no DB source, ask if needed

Read the JSON result:
- `ok:false` → report the error (an unresolved required field) and stop.
- `warnings` non-empty → tell the user exactly what was left blank or assumed and
  which flag resolves it.

## Step 3: Report

Tell the user, in a few lines:
- the output path
- any `warnings` — fields left blank (ISIC, attachment counts, itemized cost
  breakdown), and any Path-B assumption (net profit = accounting profit, no
  adjustments) that needs confirming
- that the days-until-deadline text on the form (`nDate1`/`txtMSG`/etc.) is an Adobe
  JavaScript widget that recalculates from the system clock when opened — this is
  normal, not a bug, don't try to set it

## Regenerating the template (rare)

If the government reissues the CIT50 AcroForm (new tax year revision), regenerate
the client-neutral template from ground truth:

```bash
uv run ${CLAUDE_SKILL_ROOT}/scripts/build_template.py
```

Dev-only. Re-verify `references/cit50-field-map.md` against the new form afterward —
field names/tooltips may have changed.

## Done when

- `3 CIT50 <company> <period>.pdf` exists in the client's `WP/`.
- No required field was missing (script returned `ok:true`).
- All warnings were relayed to the user with the flag/action that resolves each one.
- The user was told which year's `ภาษีเงินได้` sheet column the tax numbers came from.

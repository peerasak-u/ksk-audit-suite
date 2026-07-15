# audit-workpaper — workbook structure reference

Companion to `docs/financials-contract.md` (the authoritative shared spec). This file
documents what `scaffold_workpaper.py` actually emits, for maintenance.

## Sheets emitted (fixed order)

`0000` (veryHidden placeholder) · `งบการเงิน` · `TB` · `Mapping` · `ปรับปรุง` ·
`ภาษีเงินได้` · `_captions` (hidden vocabulary).

## The formula chain (approach A)

```
TB!G/H            ending balance  = C+E-D-F  (opening ± adjustments)
   ↓ SUMIF by account code
Mapping!C/D       C = =SUMIF(TB!$A:$A, A<r>, TB!$G:$G)   (Dr)
                  D = =SUMIF(TB!$A:$A, A<r>, TB!$H:$H)   (Cr)
                  H = caption  ← HUMAN picks from dropdown (controlled vocab)
   ↓ SUMIF by caption
งบการเงิน note-detail row   =SUMIF(Mapping!$H:$H, "<caption>", Mapping!$C:$C)
                            (prior year uses Mapping!$E / $F)
   ↓ direct ref
งบการเงิน statement line    =+E<note_row>
```

Assets/expenses take the debit column (`C` current, `E` prior); liabilities/equity/revenue
take the credit column (`D` current, `F` prior). See `SIDE` in the script.

Subtotals are `=SUM(range)` over caption rows; the liab+equity total re-sums the underlying
caption rows (not the subtotal rows) to avoid double counting. Retained earnings on the BS
references the equity statement's closing row (`opening RE from Mapping + net profit`).

## Defined names (read back by audit-financials)

`FS_TOTAL_ASSETS_CY/PY` · `FS_TOTAL_LIAB_EQUITY_CY/PY` · `FS_TOTAL_EQUITY_CY/PY` ·
`FS_NET_PROFIT_CY/PY` · `TAX_NET_PROFIT` (IS profit before tax) · `TAX_EXPENSE`
(ภาษีเงินได้ computed cell, used by the IS tax line) · `TAX_PAYABLE`.

Built-in tie-out: the row under `รวมหนี้สินและส่วนของผู้ถือหุ้น` shows
`=FS_TOTAL_ASSETS − FS_TOTAL_LIAB_EQUITY` and must read 0.

## Caption vocabulary

Defined in the script (`CUR_ASSETS … EXPENSES`), mirrored into the hidden `_captions` sheet,
and enforced on `Mapping!H` with a list data-validation dropdown. This is the controlled NPAE
vocabulary from `docs/financials-contract.md §4` — extend both places together, and keep the
strings byte-identical (the SUMIF criteria match on the literal text).

## Client-TB import (`--client-tb`)

Best-effort heuristic (`import_client_tb`): finds a header row mentioning
เลขที่บัญชี/รหัส/ผังบัญชี/บัญชี/account, then takes rows whose first cell is a numeric account
code, the first non-numeric cell as the name, and the last two numeric cells as ending Dr/Cr.
Client exports vary a lot (contract §2/§8) so it always warns to verify column alignment and
never fabricates balances. Imported balances land in TB cols C/D (opening); adjustments (E/F)
start empty; ending (G/H) is a formula.

## Accounting-policy notes

Note 2 (basis) and the note-3 policy intros for cash / inventory / PP&E / income tax are
emitted verbatim as fixed boilerplate. The revenue-recognition line and note 1 (general info)
are left as ⚠ for `audit-financials` / the human (contract §7).

## Known v1 limitations / handoff to audit-financials (skill 5.5)

- **openpyxl does not evaluate formulas.** The scaffold's cells compute live in Excel/
  LibreOffice, but reading the file back with openpyxl `data_only=True` returns `None` until a
  real spreadsheet app has opened & saved it. `audit-financials` must recalc (via a calc
  engine or by requiring the human's saved copy) before reading FS values.
- **Note-detail is one summary row per caption.** The itemized sub-account breakdown under
  each note is `audit-financials`' job, not the scaffold's.
- **บจ. going-concern only.** หจก. / งบเลิก / ปีแรก variants (contract §6) are not yet built;
  the script warns when it detects a หจก. entity_type.
- **Tax add-backs, exemptions, and credits are human ⚠ cells**; the script only wires the
  arithmetic around them (`N()` treats the blanks as 0 so the sheet computes meanwhile).

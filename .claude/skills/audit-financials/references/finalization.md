# audit-financials — finalization reference

Companion to `docs/financials-contract.md` (the authoritative shared spec) and to
`audit-workpaper` (skill 5), whose scaffold this skill consumes. Documents what
`finalize_financials.py` actually does, for maintenance.

## Why a QA gate at all, and why it needs a saved file

openpyxl writes formulas but never evaluates them. The skill-5 scaffold is a live formula
chain (approach A) that computes only inside a real spreadsheet app. So reading the workbook
back with openpyxl `data_only=True` returns `None` for every formula cell **until** a human
has opened the file in Excel/Google Sheets/LibreOffice and saved it (which caches the computed
values into the file).

This is not a limitation to route around — it *is* the gate. The premise of skill 5.5 is "the
auditor has finished adjusting the WP", and adjusting the WP means opening it in a spreadsheet
app. A workbook with no cached values was never opened, so it is not finished, so the QA gate
correctly refuses it (`validate()` → `unsaved: true`). The script loads the workbook twice:
`data_only=True` for the numbers the gate checks, `data_only=False` for the formulas the
expansion rewrites.

## QA-gate checks (`validate`)

Reads cached values only; never mutates. Hard errors block finalization; warnings are relayed.

- **unsaved guard** — `FS_TOTAL_ASSETS_CY` and `FS_TOTAL_LIAB_EQUITY_CY` both `None` → refuse.
- **error cells** — any cell whose cached value is an Excel error literal (`#REF!`, `#DIV/0!`,
  `#VALUE!`, `#NAME?`, broken external `[n]workbook` links) → refuse (contract §8).
- **balance sheet balances** — `FS_TOTAL_ASSETS_{CY,PY}` vs `FS_TOTAL_LIAB_EQUITY_{CY,PY}`
  within `TOL` (0.01 baht); also re-reads the scaffold's built-in tie-out row (`= assets −
  liab+equity`, must be 0).
- **Mapping classification** — every `Mapping` account row with a non-zero cached balance in
  C or D must have `H` filled, and every filled `H` must be in the controlled NPAE vocabulary
  (`CAPSET`).
- **profit tie** — IS profit-before-tax (`TAX_NET_PROFIT`) must equal the `ภาษีเงินได้` sheet's
  starting profit cell `C4`.
- **CIT50 reconciliation** — emitted as a *warning*: comparing `FS_NET_PROFIT_CY` to the ภ.ง.ด.50
  page-3 net profit is a manual check (reading a filled AcroForm return is out of v1 scope).
- **stale-sheet guard** — warns if a key sheet's `A1` company name does not match CONTEXT, or
  if header rows carry a 25xx year other than the CONTEXT current/prior year (contract §8).

Defined names are read via `wb.defined_names.get(name).destinations` (contract §3.3). They must
exist in the scaffold; the gate does not guess row numbers.

## Note-detail expansion (`expand_notes`)

The scaffold reserves a "รายละเอียดประกอบ" section with one summary row per caption. The finished
งบ needs the itemized per-account breakdown. openpyxl does **not** adjust formula references or
defined names when rows move, so the expansion is designed to avoid mid-sheet insertion:

1. **Re-point the statement lines.** Every statement caption line that referenced the old note
   area (`=+E<n>` with `n` ≥ the note-section start) is rewritten in place to a self-contained
   `=SUMIF(Mapping!$H:$H,"<caption>",Mapping!$C:$C)` (prior year uses `$E`/`$F`; credit-side
   captions use `$D`/`$F`). Now no statement line depends on a note-row position. The note
   number is written into column C for every caption line, including retained earnings on the
   BS (whose amount still comes from the equity statement, not a note ref).
2. **Delete the old note rows** from the section marker to the end of the sheet. All defined
   names live at rows above the marker (BS/IS/equity), so deletion cannot disturb them.
3. **Rebuild** downward from the marker: for each caption that has classified accounts, a
   numbered header (`4.`, `5.`, … — notes 1–3 are general/basis/policy) + one row per account
   linked live to its `Mapping` row (`=+Mapping!C<r>` current, `=+Mapping!E<r>` prior; credit
   side uses `D`/`F`) + a `=SUM(...)` subtotal. The subtotal equals the statement line by
   construction (both aggregate the same accounts).

Everything stays a live formula, so the workbook remains editable and self-updating. Because
openpyxl saves formulas without cached values, the output must be opened once in Excel to
recompute — the script warns about this. Re-running on the output hits the unsaved guard until
it is re-saved, and the marker-based detection makes a re-run idempotent (same row count).

## Output

A NEW file (`<wp> (final).xlsx` by default); the auditor's working file is never modified in
place. `--validate-only` runs the gate without writing anything.

## Known v1 limitations

- **บจ. going-concern only.** หจก. / งบเลิก / ปีแรก variants (contract §6) not handled.
- **CIT50 tie is a manual warning**, not an automated numeric check.
- **Single-account captions** still get a header + row + subtotal (three rows); not collapsed.
- Note text (policy notes 1–3, general-info note 1) is the scaffold's boilerplate; this skill
  expands the *numeric* note detail (notes 4+), not the narrative notes.

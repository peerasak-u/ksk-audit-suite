# audit-orchestrate — pipeline reference

Companion to `docs/OVERVIEW.md` (authoritative). Documents what `orchestrate.py` computes, for
maintenance. This skill is a read-only dashboard over the other five skills' output; it never
writes a deliverable and re-implements none of their logic.

## Stage machine

Derived per client from CONTEXT + the files in `6_ผลจากสกิล/<client>/WP/`:

| stage | condition | next |
|---|---|---|
| `0-not-ingested` | no `CONTEXT.md` | run `audit-ingest` |
| `1-phase-A-scaffolding` | CONTEXT exists, some of {planning, report, cover, workpaper} missing | run the missing skill(s) |
| `2-awaiting-human-WP-adjust` | all phase-A files present, but the งบ has no cached values | HUMAN GATE: adjust + save WP in Excel |
| `3-phase-B-finalizing` | WP adjusted, but financials/cit50 missing | run `audit-financials`, `audit-cit50` |
| `4-complete` | all 6 targets present | final human review vs folder 5 |

The **phase-B gate** ("has the human adjusted the WP?") is detected structurally: load the งบ
with `data_only=True` and read the `FS_TOTAL_ASSETS_CY` defined-name cell. openpyxl cannot
evaluate formulas, so a `None` there means the workbook was never opened & saved in a
spreadsheet app — i.e. the auditor has not done the adjustment yet. A real number means it has
been saved (values cached). This is the same signal `audit-financials`' unsaved-guard uses.

## Deliverable classification

Files are matched to one of the six targets by filename keyword (priority order in
`classify()`), because names vary between our output and the firm's (`ใบปะหน้า`, `หน้ารายงาน`
/ leading "2 ", `Planning` / "1 ", `CIT50`/`ภ.ง.ด`/"3 ", งบการเงิน/ร่างงบ/"4 "; a งบ name with
`(final)` is the financials copy). The WP folder is any subdir starting `WP` (5_ uses
`WP <year>`, 6_ uses `WP`).

## Format-QA invariants (locked format, not content)

Checked against each file type's contract, not against a ground-truth sample (content differs
per client by design):

- **Planning** — the 14 fixed sheets (`PLANNING_SHEETS`) are all present.
- **งบการเงิน / financials** — the core FS sheets (`FS_CORE_SHEETS`) present; plus the
  opened-in-Excel signal above.
- **docx** (report, cover) — opens and has ≥1 non-empty paragraph.
- **CIT50** — parses as an AcroForm with > 900 fields (the government form has ~925).

Deep งบ validation (balances, ties, classification completeness) is out of scope here — that is
`audit-financials`' QA gate. This skill only answers "is each deliverable present and in the
right shape, and what's the next step".

## Ground-truth cross-check

The matching `5_ตัวอย่างไฟล์ผลลัพธ์/` folder is found by exact folder name, falling back to the
`[NNN]` job number. Its deliverable set is reported so a missing file the real job had is
visible. It is a presence cross-check only — not a content or byte comparison.

## Known limitations

- Presence + format only; it does not open the งบ to re-check balances (that's skill 5.5).
- The stage machine assumes the standard บจ. going-concern flow; งบเลิก/หจก./ปีแรก variants
  still map onto the same stages but their extra artifacts are not separately tracked.
- `--all` reports every client folder under `6_ผลจากสกิล/`; it does not scan folder 4/5 for
  clients not yet ingested.

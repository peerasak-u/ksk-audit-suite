# CIT50 (ภ.ง.ด.50) AcroForm — field map

Derived by opening 9 usable ground-truth `3 CIT50*.pdf` files (of 10 total — 1 has no
AcroForm fields at all, likely a flattened/printed-then-scanned copy, skip it as a
source) under `5_ตัวอย่างไฟล์ผลลัพธ์/` (read-only) with `pypdf`, reading each field's
`/TU` tooltip (the Revenue Department's own Thai field description — official, not
guessed) and diffing values across files. Do not re-derive — update this file if a
re-check finds new variance.

**925 total AcroForm fields** (one file has 849 — an older form revision, don't use it
as the template source), of which **only ~140-200 are ever filled** in a real SME
return. The rest cover scenarios this firm's clients don't have (BOI promotion,
multi-currency, foreign-business specific schedules, large-corporation detail).
**~180 fields appear in 3+ of the 9 files** — treat that set as "the real form"; the
skill only needs a mapping for those.

The form is **owner-password encrypted** (standard for RD-issued forms) — `pypdf`
needs the `cryptography` package to open it (`AlgV5`/AES). Both scripts declare it as
a PEP 723 dependency.

## Field groups (by prefix)

| prefix | section | classification |
|---|---|---|
| `NID.*` | tax ID, split into fixed-width segments | CONTEXT (derived, deterministic split — see below) |
| `TXP.*` | address (เลขที่/หมู่ที่/ถนน/ตำบล/อำเภอ/จังหวัด), `TXP.1`=company name | CONTEXT (same data + same `expand_address` logic as `audit-planning`) |
| `23.1` | ประเภทกิจการ (one-line business description) | CONTEXT (business_type) — reuse audit-planning's `--business-desc` pattern |
| `24` | ISIC (Thai Standard Industrial Classification code) | **not in any DB** — no lookup table exists yet. Leave blank, warn, ask the user (never invent a code) |
| `17`-`22` | accounting period start/end (day/month/BE-year as separate fields) | CONTEXT (`period_start`/`period_end`, split — no BE/CE conversion needed, RD forms use BE) |
| `46`-`48` | filing date (day/month/BE-year) | CLI (no DB source — ask the user for the actual filing date; defaults to `sign_date` from CONTEXT if not given, flagged as an assumption) |
| `aud.0`, `aud.3` | เลขประจำตัวผู้เสียภาษีอากร ของผู้ตรวจสอบ / ของสำนักงานสอบบัญชี, spaced 1-4-5-2-1 | `aud.3` (firm's own tax ID) is a **firm-wide constant**, baked into the template. `aud.0` (the individual auditor's personal tax ID) is **not in any DB** — same gap as `audit-cover-report`'s TA `national_id` lookup (`Database ข้อมูลผู้สอบ.csv.xlsx`); reuse that lookup here too |
| `aud.1`, `aud.2` | ชื่อผู้ตรวจสอบและรับรองบัญชี, ทะเบียนเลขที่ (zero-padded to 8 digits) | CONTEXT (`auditor_name`, `auditor_license`) |
| `CEO_NAME1`, `CEO1` | ลงชื่อ (director name), ตำแหน่ง (title: กรรมการ / ผู้ชำระบัญชี for liquidation) | **same data as `audit-planning`'s `--director`** — gather once, use for both skills |
| `P3.2.2.*` | ปีปัจจุบัน (current-year) profit computation: revenue, COGS, other income, SG&A deduction, net profit (cascades through several unlabeled subtotal lines) | tax computation extraction — see below |
| `P3.2.3.*` | รวม (aggregate) column — mirrors `P3.2.2.*`; identical in every ground-truth case (single-segment SME clients have nothing to aggregate) | mirror `P3.2.2.*` unless the client has multiple business segments (none in ground truth) |
| `P3.2.1.4`, `P3.3.1.*`, `P4.4.*` (fine-grained cost/expense breakdown lines) | itemized sub-categories (ซื้อวัตถุดิบ, etc.) | **left blank in most ground-truth cases too** — the firm only fills the aggregate lines unless a client genuinely needs the detail. Don't try to force a breakdown that isn't in the source data |
| `P4.5.2.3`/`P4.5.3.3` | ดอกเบี้ยรับ (interest income) | tax computation extraction (only if the client has interest income; else blank) |
| `P5.7.2.13`/`.3.13` | ค่าทำบัญชี (bookkeeping fee) | CLI (`--bookkeeping-fee`) — not in CONTEXT |
| `P5.7.2.14`/`.3.14` | ค่าสอบบัญชี (audit fee) | **CONTEXT `audit_fee`** — same value `audit-planning` already uses |
| `P5.7.2.24`/`.3.24` | รวมรายจ่ายในการขายและบริหาร (SG&A total) | tax computation extraction |
| `P6.9.*` | balance sheet summary: เงินสด, สินทรัพย์หมุนเวียนอื่น, ทรัพย์สินสุทธิ, เจ้าหนี้การค้า, หนี้สินหมุนเวียนอื่น, ทุนจดทะเบียน (=CONTEXT `registered_capital`), ทุนที่ชำระแล้ว, กำไรสะสม | tax computation extraction (from the balance sheet, same source file) |
| `P6.10.*` | attachment index: which financial-statement pages are attached, page-range references | **file-specific page counts — not derivable without knowing the actual attached PDF's page numbers.** CLI override, else leave blank and warn |
| `Cit1` | ฐานในการคำนวณภาษี (final taxable base, after SME 300k exemption) | tax computation extraction |
| `Cit2`, `Cit5`, `Cit6`, `Cit11`, `TaxToPay`, `404`-`409` | ภาษีที่คำนวณได้ / เครดิตภาษีหัก ณ ที่จ่าย / ภาษี ภงด.51 ที่จ่ายล่วงหน้า / เงินเพิ่ม / ภาษีที่ชำระเพิ่มเติม (final payable, baht/satang split) | tax computation extraction — in every ground-truth case these come straight from the same "ภาษีเงินได้" sheet section that already computes them (see below); `404`-`409` are `-` (zero) whenever no additional tax is due, which is the norm for these clients |
| `Group1`-`Group6`, `Group91`-`Group995`, `P3.rdo2`, `P3.rdo3`, `cboExCode` | filing-type / juristic-status / related-party / currency / SME-rate-election / attachment-presence checkboxes | **firm-wide constants, baked into the template** (see below for the liquidation/loss exceptions) |
| `nDate1`, `txtMSG`, `strDate`, `calDate`, `newDueDate`, `MonthLate` | days-until-deadline countdown text | **Adobe JavaScript widgets that recompute from the system clock when the PDF is opened** — this is why they differ file-to-file even for clients with the same due date (each was opened on a different day). **Do not fill these** — leave them exactly as the template has them; Acrobat regenerates them live. |

## NID.* — tax ID split (deterministic, no lookup needed)

`NID.0`-`NID.5` are official form segments (confirmed via tooltip), not arbitrary
splitting — they are fixed-width substrings of the 13-digit `tax_id`:

| field | tooltip | width | example (`0845564003601`) |
|---|---|---|---|
| `NID.0` | จดทะเบียนที่กรมพัฒนาธุรกิจการค้า (รหัส 0) | 1 | `0` |
| `NID.1` | รหัสจังหวัดที่จดทะเบียน | 2 | `84` |
| `NID.2` | ประเภทการจดทะเบียน | 1 | `5` |
| `NID.3` | ปี พ.ศ.ที่จดทะเบียน (3 หลักสุดท้าย) | 3 | `564` |
| `NID.4` | ลำดับที่จดทะเบียน | 5 | `00360` |
| `NID.5` | Check Digit | 1 | `1` |

`tax_id[0:1]`, `[1:3]`, `[3:4]`, `[4:7]`, `[7:12]`, `[12:13]` — verified against the
[103] ground-truth case exactly (incorporation year 2564 BE → `NID.3` = `564`).

## Tax computation extraction — two paths, neither needs the agent to compute tax

Checked all 8 available `4 งบการเงิน*.xlsx` ground-truth files: **7/8 have a
`ภาษีเงินได้` or `คำนวณภาษี` sheet** with a fixed structure:

```
กำไร(ขาดทุน)สุทธิทางบัญชี
บวก รายจ่ายที่ไม่ถือเป็นรายจ่ายตามประมวลรัษฎากร (itemized: ภาษีซื้อต้องห้าม, เบี้ยปรับเงินเพิ่ม, ค่าใช้จ่ายต้องห้าม, ...)
บวก รายได้ที่ให้ถือเป็นรายได้ตามประมวลรัษฎากร
= กำไร(ขาดทุน)สุทธิทางภาษีอากร
หัก ขาดทุน(สะสม)ยกมาไม่เกิน 5 รอบบัญชี
= คงเหลือกำไรสุทธิเพื่อคำนวณภาษี
หัก ยกเว้น (SME first-300,000-THB bracket)
= กำไร(ขาดทุน)สุทธิ  ← this is Cit1 (ฐานในการคำนวณภาษี)
× อัตราภาษี (0.15 in every case seen)
= ค่าใช้จ่ายภาษีเงินได้  ← this is Cit2 (ภาษีที่คำนวณได้)
หัก ภาษีเงินได้ถูกหัก ณ ที่จ่าย  ← Cit5
หัก ภาษีเงินได้จ่ายล่วงหน้า ภงด.51  ← Cit6
= ภาษีเงินได้นิติบุคคลค้างจ่าย(ชำระเกิน)  ← TaxToPay / 404-409
```

**Path A — sheet exists:** the agent reads this sheet's cells (coordinates vary by
file — locate by label match, not fixed cell refs) and passes the values via CLI
flags (`--net-profit-accounting`, `--net-profit-tax`, `--tax-base`, `--tax-computed`,
`--wht-credit`, `--pnd51-credit`, ...). Pure extraction, no computation by the agent.

**Path B — sheet doesn't exist (verified on the one case that lacks it, [103]):**
the accounting net profit **from the `งบการเงิน` sheet's กำไร(ขาดทุน)สุทธิ line**
equals the CIT50 numbers exactly, with zero adjustments — the firm skips building a
separate worksheet when there's nothing to adjust. **Default assumption: no
adjustments** (`net_profit_tax = net_profit_accounting`, tax base = net profit minus
the 300k SME exemption if positive) unless the agent has read the client's actual
documents and found a specific add-back item — flag this assumption in the warnings,
don't silently apply it without telling the user.

**Revenue line classification (P3.2.2.1 vs P3.2.2.4) is genuinely case-by-case:**
principal trading/service revenue goes in `P3.2.2.1` (1.รายได้โดยตรงจากการประกอบกิจการ);
non-principal revenue (e.g. [103]'s real-estate rental, which is not their registered
principal activity) goes in `P3.2.2.4` (4.บวก รายได้อื่น) instead. There's no
mechanical rule for this — ask the agent to classify based on whether the revenue
matches the client's `business_type` description, default to `P3.2.2.1` when unsure.

**Checkbox flip on a loss:** `P3.rdo2`, `P3.rdo3`, `Group91`, `Group5` are `/1` in
every profit case and `/2` in the one loss case seen (KT2Power, Cit1 negative) — set
`/2` when the computed tax base is negative, else leave the template's `/1` default.

## Caveat: check which year's column you're reading

The `ภาษีเงินได้` sheet can have **more than one year's computation side by side**
(e.g. a prior-year reference column next to the current filing year), or the sheet
found in a client's `4 งบการเงิน*.xlsx` may itself be a **prior year's leftover
worksheet** carried over from an earlier engagement, not yet updated for the current
period. Always confirm the sheet's own "ณ วันที่ ..." header (or column header year)
matches CONTEXT's `period_end` before extracting numbers from it — a mismatch here is
silent and produces a completely wrong filing. `Cit1` (ฐานในการคำนวณภาษี) is **not
reliably "net profit minus the SME 300k exemption"** across cases — in at least one
verified case (เน็กซ์) it equals the net profit figure with the exemption applied later
inside the rate computation instead. `render_cit50.py` never computes this itself —
`--tax-base`/`--tax-computed` are pure passthrough — but the agent extracting them must
read whatever computation the source sheet actually shows, not assume the [103]-style
"subtract 300k before rate" pattern applies universally.

## Known gap: liquidation period-end date

Verified against the KT2Power ground-truth case: the loss-state checkbox flip
(`P3.rdo2`/`P3.rdo3`/`Group91`/`Group5` → `/2`) and `CEO1` → `ผู้ชำระบัญชี` both work
correctly, but `render_cit50.py` has **no `--period-end` override** (unlike
`audit-planning`) — fields `20`/`21`/`22` (period end) currently come straight from
CONTEXT's `period_end`, which for a liquidation client is the DB's *nominal* fiscal
year end (e.g. 31/12), not the actual dissolution cutoff date (e.g. 26/12, from the
folder name / จดทะเบียนเลิก document) that `audit-planning` already handles via its own
`--period-end` flag. Pass the correct date by hand if this matters for a specific
filing until the flag is added.

## Known gaps (ask the user, never invent)

- **ISIC code (`24`)** — no lookup table in this repo. Ask, or leave blank (it's not
  one of the fields that blocks a valid filing).
- **Attachment page counts (`P6.10.*`)** — depend on the actual page count of the
  financial-statement PDF that gets attached; CLI override or leave blank + warn.
- **`aud.0`** (auditor's personal tax ID) — same lookup gap `audit-cover-report`
  already documented for the TA `national_id` field; reuse `Database ข้อมูลผู้สอบ.csv.xlsx`.
- **`P5.7.2.13` bookkeeping fee** — not in CONTEXT, CLI override only.
- **Itemized cost/expense/balance-sheet sub-line breakdown** (`P3.2.2.3`, `P4.4.*` detail
  categories, `P5.7.2.10` and other individual SG&A lines beyond audit/bookkeeping fee,
  `P6.9.2.1.1` related-party loans, `P6.9.6.1` total liabilities) — v1 only fills the
  aggregate/total fields (revenue, COGS, SG&A total, net profit, tax base/computed,
  the balance-sheet lines with a clear single source, and the grand total via
  `--total-assets`). Breaking every aggregate into its full itemized detail needs the
  client's actual TB, not just CONTEXT — left blank and out of scope for v1, same as
  ground truth leaves most of these blank for simple clients anyway ([103] only had 3-4
  of the ~20 possible expense line items filled).

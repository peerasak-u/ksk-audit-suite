# Financials pipeline contract — skill 5 (WP scaffold) ↔ skill 5.5 (generate financials)

Status: **design spec, not yet implemented.** This is the shared contract the two skills
must both honor. Skill 5 *writes* this structure; skill 5.5 *reads* it. Neither skill is
useful if they disagree, so change this file first, both skills second.

Ground-truth basis: analyzed all 16 `4 งบการเงิน*` workbooks in `5_ตัวอย่างไฟล์ผลลัพธ์/`.
The mature exemplar for the mapping pipeline is `S [312]` (has
`Mapping ผังบัญชี` + cross-check sheets). Cell-level anatomy taken from [312],
[364], and [448] (all .xlsx with visible formulas).

---

## 0. What each skill does

```
Skill 5  — WP scaffold (before numbers/judgment)
  in : CONTEXT.md, client TB/GL export, doc inventory (from audit-ingest)
  out: 4 งบการเงิน <client>.xlsx — a scaffold workbook with:
       - correct sheet set for THIS client (skeleton + optional schedules + variant)
       - งบการเงิน sheet as a FULLY formula-linked template (see §3, approach A)
       - Mapping sheet pre-loaded with the client's accounts, caption column BLANK
       - empty AJE, tax-calc shell
       - ⚠ markers on every cell that needs human judgment
  human then: pastes/confirms TB, classifies accounts in Mapping, posts AJEs,
              exercises judgment. As they do, the งบ recomputes itself.

Skill 5.5 — finalize financials (after human adjust is done)
  in : the FINISHED scaffold workbook (TB adjusted, Mapping classified, AJE posted)
  out: - validation / QA report (BS balances, no #REF!, all accounts classified, captions
         valid, profit ties to the tax sheet; CIT50 tie is a manual warning)
       - a finalized COPY with the note detail expanded from one summary row per caption to
         itemized per-account rows (kept as LIVE formulas), '<wp> (final).xlsx'
  deterministic: given a finished WP, the gate + note expansion are mechanical.

  NOTE (revised 2026-07-14, user-approved): the งบ stays a LIVE, EDITABLE Excel. Skill 5.5
  does NOT freeze formulas to values and does NOT render a PDF — the auditor needs the
  editable workbook, and freezing would throw away the approach-A formula chain. The cover
  page (ใบปะหน้างบการเงิน.docx) is skill 2 (audit-cover-report), not 5.5.
```

> **Design choice (approach A):** we build a FULLER formula chain than the firm's real
> files do. In ground truth only the statement→note-detail link is a live formula; the
> Mapping amounts and note-detail numbers are hand-typed pastes. We make all layers live
> (§3) so the งบ self-updates and skill 5.5 has almost nothing that can go wrong. Output is
> intentionally cleaner/more automated than the reference files — not a byte-for-byte clone.

---

## 1. Fixed sheet set (names are the contract — do not rename per client)

| sheet name | role | skill 5 writes | notes |
|---|---|---|---|
| `0000` | placeholder | yes (hidden) | template artifact, keep hidden, empty |
| `TB` | adjusted trial balance | yes (layout only; client pastes data) | §2 column layout |
| `Mapping` | account → FS caption | yes (accounts + amount formulas; caption blank) | §2 |
| `ปรับปรุง` | adjusting entries (AJE) | yes (empty grid) | §2 |
| `ภาษีเงินได้` | tax computation shell | yes (shell) | name standardized to `ภาษีเงินได้` (ground truth also uses `คำนวณภาษี` — we pick one) |
| `งบการเงิน` | FS mega-sheet (BS+IS+equity+notes+note-detail) | yes (full template) | §3, §4 |
| optional schedules | supporting WP | yes (only those the client needs) | §5 library |

Row/column positions inside `TB`, `Mapping`, `ปรับปรุง` are fixed by this contract.
Inside `งบการเงิน` positions are anchored by **defined names** (§3.3), NOT absolute rows,
because variants (§6) and optional captions shift rows.

**Hard rule for skill 5.5:** never trust a sheet by name/position alone. Validate the
company name + fiscal year embedded in the sheet before consuming it — ground truth is
littered with stale leftover sheets from other clients and prior years (see §7).

---

## 2. Flat-sheet layouts (TB, Mapping, AJE)

### 2.1 `TB` — adjusted trial balance
Column order (row 1 = company, row 2 = "งบทดลอง ณ <date>", row 3 = headers):

| col | header | meaning |
|---|---|---|
| A | เลขที่บัญชี | GL account code (join key) |
| B | ชื่อบัญชี | GL account name |
| C | ยอดยกมา เดบิต | opening Dr |
| D | ยอดยกมา เครดิต | opening Cr |
| E | ปรับปรุง เดบิต | adjustment Dr (linked to `ปรับปรุง`) |
| F | ปรับปรุง เครดิต | adjustment Cr |
| G | คงเหลือ เดบิต | ending Dr = `=C+E-D-F` if Dr-natural |
| H | คงเหลือ เครดิต | ending Cr = `=D+F-C-E` if Cr-natural |

Client GL exports come in many shapes (raw `งบทดลอง` from software, or firm-reformatted
`กระดาษทำการ`). Skill 5 normalizes the client export INTO this layout; the "ending"
columns G/H are the single source of truth the Mapping layer reads.

### 2.2 `Mapping` — account → FS caption (the classification layer)
Header row: `A`=ผังบัญชี, `B`=ชื่อบัญชี, `C`=ยอด เดบิต, `D`=ยอด เครดิต,
`E`=ยอดปีก่อน เดบิต, `F`=ยอดปีก่อน เครดิต, `G`=หมวดตามหมายเหตุ, `H`=รายการในงบการเงิน (caption)

- Skill 5 fills A/B from the client chart of accounts, and sets **C/D as formulas**
  pulling the ending balance from `TB` by account code:
  `C = =SUMIF(TB!$A:$A, A7, TB!$G:$G)` (Dr), `D = =SUMIF(TB!$A:$A, A7, TB!$H:$H)` (Cr).
  *(This is the layer the firm types by hand; we make it live.)*
- Columns **G/H (the caption classification) are left BLANK with a ⚠** — this is the
  human's judgment call. `H` must be chosen from the controlled caption vocabulary (§4).
- Prior-year E/F: pasted from last year's finished งบ (or blank for ปีแรก, §6).

### 2.3 `ปรับปรุง` — adjusting entries (AJE)
Columns: `เลขที่ | คำอธิบาย | Ref | เดบิต(งบดุล) | เครดิต(งบดุล) | เดบิต(กำไรขาดทุน) | เครดิต(กำไรขาดทุน)`.
Skill 5 emits an empty numbered grid. Each posted AJE feeds `TB` cols E/F by account.
AJE content is 100% human judgment — skill 5 never invents entries.

---

## 3. `งบการเงิน` sheet — the formula-linked FS template (approach A core)

### 3.1 Vertical layout (canonical, our template controls the rows)
Ground truth varies ([312] BS rows 7–21, [364] 7–20; notes start 161 vs 169), so we
fix OUR own canonical layout and anchor by defined names, not absolute rows.

The sheet is **a stack of printed pages, not a continuous table** — see §10 for the print
contract that governs it. One statement per page, on a fixed `ROWS_PER_PAGE` grid:

```
page 1   งบแสดงฐานะการเงิน — สินทรัพย์  (ends 'รวมสินทรัพย์') + director sign-off block
page 2   งบแสดงฐานะการเงิน — หนี้สินและส่วนของผู้ถือหุ้น  (ends 'รวมหนี้สินและส่วนของผู้ถือหุ้น')
page 3   งบกำไรขาดทุน  (ends 'กำไร(ขาดทุน)สุทธิ')
page 4   งบแสดงการเปลี่ยนแปลงส่วนของผู้ถือหุ้น
page 5+  หมายเหตุประกอบงบการเงิน  (note 1 ข้อมูลทั่วไป → note 2 เกณฑ์ → note 3 นโยบาย → notes 4+)
         then the note-detail area (the rows the statement lines pull from), which
         skill 5.5 rebuilds and which flows across as many pages as it needs.
```
Every page opens with the same 6-row furniture:
```
+0  company name          (page 1 literal; every later page '=+A1')
+1  statement title       (BS page 2 '=+A2'; other statements literal)
+2  period line           (BS page 2 '=+A3'; equity/notes '=+A<IS period row>')
+3  (blank)
+4  E: 'หน่วย : บาท'
+5  C: 'หมายเหตุ'  E: CY  G: PY   (later pages ref the first column header)
```
Columns everywhere: **A**=caption, **C**=note#, **E**=current year, **G**=prior year.
**B/D/F carry no data** — D and F are thin print gutters (§10), B is caption spill-over.
**H onward is off-page**: the BS tie-out check formula lives in **column I**, outside the
`$A:$G` print area, so it never prints on a client deliverable but stays visible to the
auditor and readable by skill 5.5.

### 3.2 The 4-layer formula chain (this is what "approach A" means)
```
Layer 1  TB!G/H            ending balances (client pastes / AJE-adjusted)
   ↓  SUMIF by account code
Layer 2  Mapping!C/D       per-account balance (formula from TB)  + human sets caption H
   ↓  SUMIF by caption
Layer 3  note-detail row   e.g. E253 เงินสด = =SUMIF(Mapping!$H:$H,"<caption>",Mapping!$C:$C)
   ↓  direct ref
Layer 4  statement line    e.g. E10 (เงินสดฯ in BS) = =+E253
```
Firm files today: only Layer 4←3 is a formula; 3 and 2 are hand-typed. **We make all live.**
Net effect: once the human sets the caption in `Mapping!H` and the TB is adjusted, every
statement number and subtotal recomputes with zero manual typing.

Subtotals stay `=SUM(range)` (e.g. รวมสินทรัพย์หมุนเวียน). Tax line in IS references the
`ภาษีเงินได้` sheet result. Retained-earnings on BS references the equity statement row.

### 3.3 Defined names (the anchor skill 5.5 reads — required)
Skill 5 must create these workbook-level defined names so 5.5 can validate without guessing
row numbers. Minimum set:
`FS_TOTAL_ASSETS_CY`, `FS_TOTAL_ASSETS_PY`, `FS_TOTAL_LIAB_EQUITY_CY`,
`FS_TOTAL_LIAB_EQUITY_PY`, `FS_NET_PROFIT_CY`, `FS_NET_PROFIT_PY`,
`FS_TOTAL_EQUITY_CY`, `TAX_NET_PROFIT`, `TAX_PAYABLE`.
Skill 5.5 checks: `FS_TOTAL_ASSETS_CY == FS_TOTAL_LIAB_EQUITY_CY` (BS balances),
`FS_NET_PROFIT_CY` ties to `ภาษีเงินได้` starting profit and to CIT50 P3 net profit.

---

## 4. Controlled caption vocabulary (NPAE รายการย่อ)

`Mapping!H` and the statement captions must come from ONE controlled list (like the
assertion-code legend in audit-planning). Authority: ประกาศกรมพัฒนาธุรกิจการค้า
27 ต.ค. 2566 "รายการย่อที่ต้องมีในงบการเงิน (NPAE)". Common captions seen in ground truth
(to be completed against the full NPAE schedule before build):

**สินทรัพย์:** เงินสดและรายการเทียบเท่าเงินสด · ลูกหนี้การค้าและลูกหนี้หมุนเวียนอื่น ·
เงินให้กู้ยืมระยะสั้น · สินค้าคงเหลือ · สินทรัพย์หมุนเวียนอื่น · ที่ดิน อาคารและอุปกรณ์ ·
สินทรัพย์ไม่มีตัวตน · สินทรัพย์ไม่หมุนเวียนอื่น
**หนี้สิน:** เจ้าหนี้การค้าและเจ้าหนี้หมุนเวียนอื่น · เงินกู้ยืมระยะสั้น · ภาษีเงินได้ค้างจ่าย ·
เงินกู้ยืมระยะยาว · หนี้สินไม่หมุนเวียนอื่น
**ส่วนของผู้ถือหุ้น:** ทุนที่ออกและชำระแล้ว · กำไร(ขาดทุน)สะสม
**งบกำไรขาดทุน:** รายได้จากการขายหรือบริการ · รายได้อื่น · ต้นทุนขายหรือบริการ ·
ค่าใช้จ่ายในการขายและบริหาร · ต้นทุนทางการเงิน · ค่าใช้จ่ายภาษีเงินได้

Each caption maps to a fixed note number. Note numbering is generated programmatically
(ground truth numbering is unreliable — duplicated "3.2", etc.).

---

## 5. Optional supporting-schedule library (skill 5 picks by account type)

Trigger a schedule when the client's chart of accounts contains the matching account:

| schedule sheet | trigger account | seen in |
|---|---|---|
| `ค่าเสื่อม` (depreciation) | ที่ดิน อาคาร อุปกรณ์ | most |
| `สินค้าคงเหลือ` (inventory) | สินค้าคงเหลือ | [312], [487] |
| `ทะเบียนคุมที่ดิน` / `อาคารระหว่างก่อสร้าง` | ที่ดิน / งานระหว่างก่อสร้าง | [103], [202], [78] |
| `รายได้ค้างรับ` (progress billing) | รายได้ค้างรับ / งานตามสัญญา | [312], [งาน33] |
| `เช่าซื้อ` (hire-purchase) | เจ้าหนี้เช่าซื้อ | [งาน33] |
| `เงินเดือน` (payroll) | เงินเดือน / ประกันสังคม | [งาน33] |
| `ดอกเบี้ยค้างรับ` | ดอกเบี้ยค้างรับ | [278], [งาน4] |
| `wp_เงินสด` (cash count) | always | all |
| `ยืนยันยอด_เงินให้กู้` / `wp_เงินให้กู้` / `สัญญาเงินให้กู้` | เงินให้กู้ยืม | ~13/16 |
| `wpเจ้าหนี้การค้าและเจ้าหนี้อื่น` | เจ้าหนี้ | ~12/16 |
| `ภพ.30` (VAT recon) | client is VAT-registered | [312], [278] |

---

## 6. Entity / engagement variants (each needs a template delta)

Focus v1 on **บจ. going-concern**; add the others as explicit variants.

- **บจ. going-concern** (default) — share capital, sign-off "กรรมการ".
- **หจก. (partnership)** — equity section = per-partner capital lines
  (`ทุน - นาย…`), extra "เงินส่วนแบ่งของกำไร" line, no หุ้นสามัญ/ทุนจดทะเบียน;
  sign-off "หุ้นส่วนผู้จัดการ". (เจอที่: [485])
- **งบเลิก (liquidation)** — BS dated at dissolution date (not fiscal year-end);
  add "งบแสดงการเปลี่ยนแปลงสินทรัพย์สุทธิเพื่อการชำระบัญชี"; accounting-policy note
  switches to liquidation basis (เกณฑ์ชำระบัญชี); often no TB (replaced by multi-year
  catch-up `บันทึกบัญชี##`). (เจอที่: KT2Power)
- **ปีแรก (first year)** — no prior-year WP; prior-year column (G) blank or from client's
  opening TB / prior auditor; no internal prior-year TB sheet. (เจอที่: [งาน4])

---

## 7. Accounting-policy notes — boilerplate vs client-specific

~70–80% of the policy note block is verbatim-reusable boilerplate. Skill 5.5 generates the
boilerplate and inserts client-specific bits from CONTEXT, leaving ⚠ where judgment is needed.

**Fixed boilerplate (generate verbatim):**
- Note 2 เกณฑ์การจัดทำงบการเงิน — 100% fixed (references มาตรฐาน NPAE, ประกาศ กรมพัฒน์ฯ
  27 ต.ค. 2566, เกณฑ์ราคาทุนเดิม).
- Note 3.x intro paragraphs for: เงินสดและรายการเทียบเท่าเงินสด, สินค้าคงเหลือ (cost basis),
  ที่ดิน อาคารและอุปกรณ์ (straight-line intro), ภาษีเงินได้นิติบุคคล — all fixed.

**Client-specific (from CONTEXT / human):**
- Note 1 ข้อมูลทั่วไป — legal name, registration date, business description, address, reg no.
- Asset-class useful-life table under PP&E note (rates differ per client).
- Revenue-recognition first sentence — differs by business type (trading vs service).
- Which sub-notes exist at all (e.g. intangibles note only if client has them).
- Entity word: use "บริษัท" vs "ห้าง" per entity type (ground truth mis-copies this).

**Numbering:** always renumber notes programmatically; never copy ground-truth numbering.

---

## 8. Data-quality guards (both skills must enforce)

Every one of these appears in the real 16 files — the pipeline must tolerate/flag, not
propagate:
- Stale leftover sheets belonging to a different client ([202] `ปก`, [278] `งบทดลอง`,
  [งาน4] 3 sheets) → validate embedded company name/year, don't trust sheet names.
- `#REF!` / broken external `[n]workbook` links ([364], [448], [312]) → flag, refuse to
  freeze a งบ that still contains error cells.
- Stale year labels (tax/AJE sheets headed 2561/2564 in 2568 jobs) → derive year from
  CONTEXT `period_end`, never from a hardcoded sheet label.

---

## 9. Open items / resolved decisions
- Complete the §4 caption vocabulary against the full NPAE รายการย่อ schedule. *(still open)*
- **RESOLVED — file identity:** the งบ stays a live sheet inside the WP workbook. Skill 5.5
  writes a finalized COPY (`<wp> (final).xlsx`) with note detail expanded, never modifying the
  auditor's working file, and never freezing to values.
- **RESOLVED — defined-name set (§3.3):** sufficient for all skill-5.5 gate checks (BS balance,
  profit tie, tie-out). Confirmed against the built scaffold.
- **RESOLVED — formula evaluation for skill 5.5:** require the human's Excel-saved copy (cached
  values present) rather than a calc engine — LibreOffice is not installed and the `formulas`
  lib is too slow. This is not a workaround: a WP never opened & saved in a spreadsheet app is
  by definition not finished, so skill 5.5's QA gate refuses it (reads all formula cells as
  `None` → `unsaved` error with the "open in Excel and save" instruction). Skill 5 (scaffold)
  is unaffected — it computes live in Excel while the human works.

---

## 10. Print contract — the งบ is a printed deliverable (skill 5 sets it, 5.5 preserves it)

The `4 งบการเงิน` workbook is not an internal grid: its `งบการเงิน` sheet gets **printed and
bound behind `ใบปะหน้างบการเงิน.docx` (skill 2's cover) and the auditor's report**, then given
to the client and filed with DBD. All 10 ground-truth `4 *` files agree on the settings below,
so these are a **locked format**, not a preference. Verified 2026-07-17 against folder 5.

The single strongest tell: **every ground-truth file sets `firstPageNumber=4`** — the งบ is
page 4 onward because the cover + report occupy pages 1–3. A งบ with no page setup at all
(what skill 5 produced before this section existed) is a *format defect*, not a cosmetic one.

### 10.1 Sheet-level settings (`งบการเงิน`)
| setting | value | why |
|---|---|---|
| paper / orientation | A4 (`paperSize=9`), portrait | 10/10 ground truth |
| `firstPageNumber` / `useFirstPageNumber` | **4** / true | bound after cover + report (pages 1–3) |
| footer | right-aligned page number, Angsana New 14 (`&R&"Angsana New,Regular"&14&P`) | 10/10 |
| fit | `fitToPage`, `fitToWidth=1`, `fitToHeight=0` | 1 page wide, flow tall |
| print area | `$A$1:$G$<last used row>` | keeps col I tie-out off the page |
| margins (in) | L 0.79 · R 0.2 · T 0.59 · B 0.39 (header 0.51 · footer 0.28) | modal ground truth |
| default (Normal) font | Cordia New 14 | ground truth workbook default |
| cell font on `งบการเงิน` | **Browallia New 14** | 10/10 — every statement cell |
| row height | **21.0 pt, uniform** | this is what makes the page grid predictable |
| number format | accounting `_(* #,##0.00_);_(* \(#,##0.00\);_(* \-??_);_(@_)` | ground truth |

### 10.2 Column widths — D and F are print gutters, not data columns
```
A 31.7   B 16.7   C 16.7   D 0.9   E 16.7   F 0.9   G 16.7      (Σ ≈ 100.3)
caption  spill    note#    gutter  CY       gutter  PY
```
This is tuned, not arbitrary: Σ ≈ 100.3 char-units ≈ 7.31in ≈ the 7.28in printable width of
A4 portrait at the §10.1 margins. **Do not widen D/F to data width** — they carry no values,
and doing so pushes the sheet past one page wide (the original defect: `D=F=16` → Σ 121).

### 10.3 The page grid
`ROWS_PER_PAGE = 36` (36 × 21pt = 756pt ≤ the 771pt printable height of A4 at these margins).
Each statement is padded out to the grid and closed with an **explicit row break**, so a
statement never splits across a page and page N always starts with its own §3.1 header.
Skill 5 must warn if any statement's content exceeds the grid rather than silently overflow.

Titles (company / statement / period) are **merged `A:G` and centre-aligned**; the period row
carries a thin bottom border across `A:G` as the header rule.

### 10.4 Director sign-off block (bottom of the BS-assets page)
Fixed boilerplate, verbatim from ground truth (9/10 บจ. files):
```
งบการเงินนี้ได้รับอนุมัติจากที่ประชุมใหญ่สามัญผู้ถือหุ้น ครั้งที่ 1/<CY+1> เมื่อวันที่.......................
(blank)
ขอรับรองว่าถูกต้อง
(blank)
ลงชื่อ....................................................................กรรมการ
       ( <director name> )
```
**Variants:** หจก. has **no** approval line (a partnership has no shareholder meeting) and signs
**หุ้นส่วนผู้จัดการ**, not กรรมการ — confirmed in `[485] แจใจ`.
**Never invent the director name.** Take it from CONTEXT `director_1`; if that field is ⚠/⟨FILL⟩,
or its confidence column is ⚠, or the value is annotated *unconfirmed*, emit the dotted `( ...... )`
line and warn — do not guess (§ project rule: never-invent).

### 10.5 Skill 5.5's obligation
`expand_notes` rebuilds everything from the note marker down, which is **below** every statement
page break — so the §10.3 breaks survive. But 5.5 **must reset `print_area` to `$A$1:$G$<new last
row>`** after expanding, or the added note rows print outside the page. 5.5 must not touch the
statement pages' geometry.

---

## Implementation status
- **§10 print contract — built (2026-07-17).** Added after review found the `งบการเงิน` sheet
  shipped with *no page setup at all* across all 13 generated clients: no paper size, no print
  area, no footer, Calibri 11, default row heights, and D/F at data width (Σ 121 units = 8.82in
  → 21% past A4's 7.28in printable width, i.e. a second page sideways). Root cause: this contract
  specified the formula chain but never the print format, so the scaffold built a data grid.
  Skill 5 now sets §10.1–10.4 (verified against all 10 ground-truth files: paper, first page 4,
  footer, print area, fonts, 21pt grid, column widths, sign-off block incl. the หจก. variant);
  skill 5.5 does §10.5; `audit-orchestrate` gates it (`qa_fs_print`, verified to flag the old
  format on 13 dimensions and pass the new one). Note the ground-truth files disagree with each
  other on col-A width (28.1–45.3) and margins (0.24–0.79) — we take the modal value.
- **Skill 5 (`audit-workpaper`) — built.** Scaffolds บจ. going-concern: the full §1 sheet set,
  the §3 four-layer formula chain, §3.3 defined names, the §4 caption dropdown, §7
  boilerplate notes, and the §10 print contract. Best-effort client-TB import. Verified: a balanced TB flows through the
  chain to a balancing BS (tie-out row = 0). See `.claude/skills/audit-workpaper/`.
- **Skill 5.5 (`audit-financials`) — built.** QA gate (unsaved guard, error-cell scan, BS
  balance CY/PY, tie-out row, Mapping-H completeness + valid vocabulary, profit↔tax-sheet tie,
  stale company/year warnings; CIT50 tie is a manual warning) + note-detail expansion (statement
  lines re-pointed to self-contained SUMIF-by-caption, itemized per-account note rows linked live
  to Mapping, subtotals). Keeps the workbook live/editable — no freeze, no PDF. Writes a `(final)`
  copy; never touches the working file. See `.claude/skills/audit-financials/`.
- Variants หจก. / งบเลิก / ปีแรก (§6): not yet built.

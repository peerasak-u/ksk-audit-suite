# Planning.xlsx — structure and field map

Derived by opening all 16 ground-truth `1 Planning*.xlsx` files under
`5_ตัวอย่างไฟล์ผลลัพธ์/` (read-only) and diffing every cell. Do not re-derive —
update this file if a re-check finds new variance.

## The 14 sheets

`ข้อมูลลูกค้า, ประเภทบัญชี, 000, 001, 101, 102, 103, 202, 203 TB (2), 301, 302, 401, 601, 608`
— identical set/order in all 16 real cases.

**9 of the other 13 sheets pull from `ข้อมูลลูกค้า` via existing formulas** (202 has 45
refs, 601 has 18, etc.) — confirmed by counting `=ข้อมูลลูกค้า!...` formula cells across
every ground-truth file; the counts are stable case to case. **Only 3 sheets ever need
a script write: `ข้อมูลลูกค้า`, `301` (cell range A28:E28), `203 TB (2)` (rows 5-10).**
Every other sheet/cell must be left completely untouched — touching them risks breaking
a cascade formula or a print-area/format that the firm's template already has correct.

## `ข้อมูลลูกค้า` — row map

Column A = fixed label (never touch). Column C = the sheet's own permanent
instructional example text (never touch, never filled per client). Column B is the
only thing a script writes, and only some rows — see classification.

| row | label | classification | source |
|---|---|---|---|
| 1 | รอบบัญชี | CONTEXT | Thai BE year parsed from `period_end` |
| 2 | ชื่อ | CONTEXT | `company_legal_name` (fallback `company_name`) |
| 3 | ประเภทกิจการ | CONTEXT | `entity_type` → "บริษัท" (บจ.) / "ห้างหุ้นส่วนจำกัด" (หจก.); a liquidating (`is_liq`) บจ. renders the full "บริษัทจำกัด" instead of bare "บริษัท" — confirmed against the one งบเลิก ground-truth case ([404]) and the sheet's own D3 example cell, which only ever shows "บริษัทจำกัด" |
| 4 | ที่อยู่ | CONTEXT (light format) | `address`, prefixed with `"เลขที่ "` (plain space, never a line break) unless it already starts with "เลขที่". Sampled all 16 ground-truth files: 11/16 add the prefix (8 with a space, 3 with a stray newline — space wins), 5/16 have no prefix at all — CONTEXT's raw `address` field never contains "เลขที่" itself for either group, so there is no reliable per-client signal; space-prefixed-by-default is the closest single convention, not a 16/16 guarantee |
| 5 | เลขประจำตัวผู้เสียภาษี | CONTEXT | `tax_id` |
| 6 | ประเภทธุรกิจ | CONTEXT (light format) | บริษัท: `"บริษัทฯ ประกอบกิจการ\n" + business_type`; หจก.: bare `business_type` (no prefix — confirmed from the one หจก. ground-truth case) |
| 7 | วันอนุมัติ | rare, leave blank | almost always blank in ground truth (2/16 had a value) — leave `None` unless `--approval-date` given |
| 8 | วันจัดตั้ง | CONTEXT | `incorporation_date` |
| 9 | วันต้นงวด | CONTEXT (derived) | 1 Jan of `period_end`'s BE year, UNLESS: `--period-start` is given (explicit override, e.g. งบเลิก) or the client is first-year (`incorporation_date` falls after that notional 1-Jan, i.e. inside the audited period) — then `incorporation_date` itself is used, matching the ground-truth first-year case ("งาน 4_", incorporation 03/04/2025) |
| 10 | วันสิ้นงวด | CONTEXT | `period_end` |
| 11 | วันต้นงวด(ของปีเก่า) | CONTEXT (derived) | 1 Jan of prior BE year — left blank for liquidation clients (matches ground-truth liquidation case) and for first-year clients (no prior year existed — matches ground-truth "งาน 4_" case; do NOT fill with CONTEXT's `prior_period_end`, which is a mechanically-derived placeholder, not a real prior year, for a first-year client) |
| 12 | วันสิ้นงวด(ของปีเก่า) | CONTEXT | `prior_period_end` — same first-year/liquidation blank rule as row 11 |
| 13-15 | กรรมการ1-3 | CONTEXT / CLI | `director_1/2/3`; ⚠ if still unresolved in CONTEXT — pass `--director NAME` (repeatable, up to 3) |
| 16 | อำนาจกรรมการ | CLI, optional | `--director-authority "..."`; blank if not given (13/16 ground-truth cases leave it blank) |
| 18 | ชื่อผู้สอบบัญชี | CONTEXT | `auditor_name` |
| 19 | ผู้สอบบัญชีรับอนุญาตเลขทะเบียน | CONTEXT | `"ผู้สอบบัญชีรับอนุญาต เลขทะเบียน {auditor_license}"` (CPA) — TA phrase not yet seen in ground truth, flag ⚠ if `auditor_type` is TA |
| 20 | จำนวนผู้สอบบัญชีในทีม | CLI | count of `--audit-team NAME` values passed (excludes the lead auditor — ground truth confirms the count matches the assistant list only) |
| 21-22+ | ชื่อผู้สอบบัญชีในทีม | CLI | one `--audit-team NAME` per row, starting row 21 |
| 26 | ตรวจสอบตั้งแต่วันที่ | CLI, required for full plan | `--audit-start-date`; ⚠ if omitted (script still writes the file, just flags it) |
| 27 | ถึงวันที่ | **formula, do not touch** | `=+B26` already in template |
| 28 | ค่าตรวจสอบบัญชี | CONTEXT | `audit_fee` |
| 29 | จำนวนชั่วโมง | CLI, optional | `--audit-hours N`; ⚠ if omitted — no DB source, genuinely engagement-specific (ground truth ranges 4-24 with no formula) |
| 30 | ประมาณการค่าใช้จ่าย | auto-estimate ⚠ | default `round(audit_fee * 0.3, -2)` — one ground-truth file literally hardcodes the formula `=+B28*0.3`, confirming 30% is the firm's own rule of thumb; flagged as an estimate for human review, override with `--audit-expense N` |
| 31 | วิธีที่ใช้ในการตรวจสอบ | **constant, baked into template** | "ตรวจสอบเนื้อหาสาระ" (16/16) |
| 33 | เนื่องจาก | **constant, baked into template** | "กิจการว่าจ้างผู้ทำบัญชีภายนอก เพื่อบันทึกบัญชีและจัดทำงบการเงิน" (16/16) |
| 37 | คาดการณ์งบประมาณในการตรวจสอบ | **formula, do not touch** | `=+B30` already in template |
| 39 | วันที่ในหนังสือรับรองข้อมูล | CLI, optional | `--rep-letter-date`; ⚠ if omitted |
| 41-43 | ผู้ใช้งบการเงิน | **constant, baked into template** | 3-line list: กรรมการ เจ้าของกิจการ ผู้ถือหุ้น / ธนาคาร / หน่วยงานราชการ เช่น กรมสรรพากร กรมพัฒนาธุรกิจการค้า (16/16) |
| 45 | ระดับความมีสาระสำคัญ | leave blank | blank in 15/16 ground-truth files (the one exception looks like a data-entry slip into the wrong row) |
| 46 | กำหนดจาก | CLI, no reliable default, ⚠ | no auto-fill — pass `--materiality-basis "สินทรัพย์รวม"` or `"รายได้รวม"`. Sampled all 16 ground-truth files: 4/16 สินทรัพย์รวม, ~8/16 รายได้รวม (incl. one recovered from an adjacent-row data-entry slip), 3/16 left blank outright, 1/16 corrupted (numeric value in the wrong row) — รายได้รวม edges out as the largest single bucket but not by enough to safely auto-fill (and the 3 legitimate blanks show the firm sometimes deliberately leaves this for human judgment); left blank with a warning if omitted, mirroring row 49's pattern |
| 47 | เนื่องจาก | default, overridable | short reason tied to `--business-type` (see below); override with `--materiality-reason "..."` |
| 48 | ช่วงเวลาของจำนวนที่ใช้เป็นฐาน | CONTEXT (derived) | `"ณ วันที่ " + period_end` |
| 49 | จำนวนที่ใช้เป็นฐาน | CONTEXT, ⚠ | `materiality_base` — not yet populated by audit-ingest (later phase, needs งบการเงิน); leave ⚠, override with `--materiality-base N` |
| 50 | อัตราที่ใช้หา OM | **constant, baked into template** | 0.02 (16/16) |
| 51 | อัตราที่ใช้หา PM | **constant, baked into template** | 0.75 (16/16) |
| 52 | เนื่องจาก | **constant, baked into template** | "มีความเสี่ยงในระดับปานกลาง" (16/16) |
| 53 | อัตราที่ใช้หา AM | **constant, baked into template** | 0.1 (16/16) |
| 54 | เนื่องจาก | default, overridable | tied to `--business-type` (see below); override with `--prior-adjustment-note "..."` |
| 56 | สภาพแวดล้อมของกิจการ | default, overridable | tied to `--business-type` — **cascades into sheet 301!B20 via formula**, so this one cell drives that sheet's header too |
| 59, 62, 65 | การเปลี่ยนแปลง... (3 rows) | **constant, baked into template** | "ไม่มี" (16/16) |

## Sheets `301` and `203 TB (2)` — boilerplate, not bespoke narrative

**Key finding (2026-07-14):** despite the original design assumption that these sheets
hold client-specific analysis "grounded in real documents," the ground truth shows the
firm reuses **the same ~3 boilerplate blocks verbatim across unrelated industries** —
e.g. the "อย./inventory spoilage" risk text appears identically in a steel trader, a
tour operator, and a real-estate rental company. Confirmed by the user: **replicate the
firm's actual behavior** (deterministic template selection), don't attempt bespoke
document-grounded prose.

Sheet `301` rows 1-27 are **100% fixed across all 16 files** (header formulas + the
first two risk items, word-for-word identical) — baked permanently into the template.
Only row 28 (the third, optional risk item) and `203 TB (2)` rows 5-10 vary, and only
across 3 known variants, selected by `--business-type`:

### `default` (11/16 ground-truth cases — retail/consumer goods; also the fallback)

`301!A28:E28`:
- A28: `มีผลกระทบเนื่องจากมีปัญหาเรื่อง อย.`
- B28: `สินค้าคงเหลือเสื่อมสภาพ `
- C28: `V`
- D28: `มีโอกาสเกิดขึ้น แต่ไม่มาก เนื่องจากสินค้าส่วนใหญ่เป็นสารตั้งต้น และเหลือจำนวนไม่มากนัก`
- E28: `ไม่`

`203 TB (2)` rows 5-10:
- B5: `รายได้จากการขาย ลดลง` / C5: `เนื่องจากสินค้าต้องมีการแก้ไข อย.`
- B6: `ค่าใช้จ่ายส่วนที่เกี่ยวข้องกับการขาย ก็ลดลงด้วย  เช่น ค่าธรรมเนียม Lazada / ค่าน้ำมัน `
- B7: `ทั้งนี้มีค่าโฆษณาที่เพิ่มขึ้น เมื่อตรวจทานเอกสาร และสอบถามผู้บริหาร พบว่าเป็นค่า`
- B8: `จัดทำ วิดิโอ โปรโมท ซึ่งเตรียมไว้ใช้โฆษณาหลังจากที่ดำเนินการเรื่อง อย.ให้ถูกต้องเรียบร้อย `
- B10: `มีเงินกู้ยืมระยะสั้นเพิ่มขึ้น เนื่องจากต้องนำมาใช้หมุนเวียนในการดำเนินการกับกิจการ`

`ข้อมูลลูกค้า` row 47/54/56 defaults: `กิจการมีการรายได้จากการขาย` / `มีรายการปรับปรุงไม่มาก` / `กิจการเริ่มจำหน่ายสินค้าได้ `

### `construction` (3/16 — real estate / construction / % of completion revenue)

`301!A28:E28`:
- A28: `การรับรู้รายได้ตามอัตราร้อยละของงานที่ทำเสร็จ \nอาจทำไม่ได้จริง เนื่องจาก มีการประเมินค่อนข้างยาก อีกทั้งเมื่อพิจารณษจากการบันทึกบัญชี พบว่ากิจการน่าจะบันทึกรายได้ เมื่อได้รับชำระ`
- B28: `บันทึกรายได้ไม่ครบถ้วน \nบันทึกค่าใช้จ่ายสูงกว่าความเป็นจริง`
- C28: `C\nA\nCO`
- D28: `มีโอกาสเกิดขึ้นสูง แต่ไม่ทราบจำนวน ว่ากระทบกับงบการเงินมากน้อยเพียงใด`
- E28: `ใช่`

`203 TB (2)` rows 5-6:
- B5: `มีรายได้จากการบริการเพิ่มขึ้น` / D5: `เนื่องจากมีการดำเนินงานเป็นปีแรก`
- B6: `ค่าใช้จ่ายส่วนที่เกี่ยวข้อง ทั้งต้นทุน และค่าใช้จ่ายบริหารก็มีจำนวนเพิ่มขึ้นเช่นกัน`

`ข้อมูลลูกค้า` row 47/54/56 defaults: `กิจการมีการดำเนินงานปกติ` / `ปีก่อนมีรายการปรับปรุงไม่มาก` / `ไม่มี`

### `dormant` (1-2/16 — no operations / งบเลิก liquidation)

`301!A28:E28`: leave the whole row blank (no third risk item — matches the one ground-truth
dormant case exactly).

`203 TB (2)` row 5: `B5: กิจการไม่มีการดำเนินงาน` (rows 6-10 blank).

`ข้อมูลลูกค้า` row 47/54/56 defaults: `กิจการไม่มีรายได้` / `ปีก่อนมีรายการปรับปรุงไม่มาก` / `ห้างเลิกกิจการ` (or `ไม่มี` for a merely-dormant, non-liquidating case — pass `--dormant-reason` to distinguish)

**Auto-detection:** `--business-type` defaults to `dormant` when CONTEXT's
`juristic_status` mentions เลิก/ชำระบัญชี/ร้าง or the client folder name contains
"งบเลิก" (same signal `audit-cover-report` uses for the liquidation report variant) —
this is a real, legally-grounded signal and is trusted. For everything else, the
script now **refuses to render** without an explicit `--business-type` (2026-07-15
change — see below); the calling agent must ask the user, not guess.

**Confirmed empirically (2026-07-15): no field in CONTEXT predicts `default` vs
`construction`.** Cross-checked business_type/revenue/juristic_status against the
actual variant used in all 16 ground-truth files:
- งาน 4 (business_type literally `"การก่อสร้างอาคารที่ไม่ใช่ที่พักอาศัย"` — construction
  of buildings) → ground truth used `default`.
- งาน33 (business_type `"การบริการดูแลและบำรุงรักษาภูมิทัศน์"` — landscaping, nothing
  construction-related) → ground truth used `construction`.
- [404] (business_type `"การขนส่งและขนถ่ายสินค้า"` — transport) → ground truth used
  `construction`.
- งาน364 → business_type completely unresolved (⚠, no DB match at all) → ground truth
  still used `construction`.
- [485] (revenue ฿23M, `juristic_status` = "ยังดำเนินกิจการอยู่", i.e. actively
  operating, not dormant by any measure in CONTEXT) → ground truth used the
  "no row-28 risk item" shape anyway (the outlier noted below) — revenue-based
  dormant detection would also have gotten this one wrong.

Conclusion: the firm's real selection looks like case-by-case auditor judgment, not
something rule-derivable from DBD/DB งาน data. An auto-detect heuristic here would
guess wrong about as often as the current silent default did — worse, it would look
authoritative while doing so. Hard-blocking and asking the user is the only safe
option with the data this pipeline has access to.

**One ground-truth outlier does not fit cleanly:** the หจก. dormant-shaped case
(no 301 row 28, one-line 203 TB) actually has `กิจการมีการดำเนินงาน` ("business is
operating") in rows 47/54/56 — the opposite of the KT2Power liquidation case's text,
even though both share the same "no third risk item" *shape*. `--business-type
dormant` gets the shape right (empty row 28, short TB block) but the row 47/54/56
literal text and the TB row 5 text need `--materiality-reason` /
`--prior-adjustment-note` / `--biz-environment` / `--tb-note` overrides — don't trust
the `dormant` defaults blindly for a merely-quiet (not liquidating) client.

## Known gap: liquidation (งบเลิก) needs its own full template, not just business_type=dormant

Verified against the one งบเลิก ground-truth case ([404]): sheets `103`,
`601`, `608` use **liquidation-specific wording** ("ผู้ชำระบัญชี" instead of "กรรมการ",
different clause ordering/wrapping in the engagement-letter and management-letter
sheets) that the standard template (built from a non-liquidation source file) does not
have — this is not something `--business-type` or any `ข้อมูลลูกค้า` cell can fix, the
same way `audit-cover-report` needed a dedicated `report_cpa_liquidation.docx` asset
rather than a flag on `report_cpa.docx`. **`render_planning.py` currently only builds
the standard (non-liquidation) template correctly.** For a liquidation client, the
`ข้อมูลลูกค้า`/`301`/`203 TB (2)` cells this script writes are still correct (verified
scoped-diff clean against the one liquidation ground-truth case with `--period-end`
override for the actual dissolution cutoff date), but sheets 103/601/608 will read as
a normal engagement, not a liquidation one, until a `planning_template_liquidation.xlsx`
variant is built the same way `build_template.py` builds the standard one. Flag this to
the user before treating a liquidation Planning.xlsx as final.

## Confirmed cross-file drift in the firm's own ground truth (not a bug to chase)

Ground-truth files disagree with **each other**, not just with our output, on cells
outside `ข้อมูลลูกค้า`/`301` row 28/`203 TB (2)` — e.g. sheet `202` row 27 is a broken
`=ข้อมูลลูกค้า!#REF!` formula in the [103] ground-truth file but a working
`=ข้อมูลลูกค้า!B21` in the [งาน33] one; `103` row 8 col 7 is `=+ข้อมูลลูกค้า!B9` in one
file and `=ข้อมูลลูกค้า!$B$9` in another; some files even overwrite the `ถึงวันที่`/
`คาดการณ์งบประมาณ` formula cells (rows 27/37) with a hardcoded literal value. This
confirms the firm's own master file has drifted across client copies over time —
100% cell parity between our single template and every individual ground-truth file's
untouched sheets is not achievable and not the bar to hold this skill to. **Verify with
the scoped diff** (only `ข้อมูลลูกค้า` col B + `301!A28:E28` + `203 TB (2)!`rows 5-10 —
the cells this script actually writes), not a full-sheet diff.

# /// script
# requires-python = ">=3.12"
# dependencies = ["openpyxl"]
# ///
"""Scaffold the audit working-paper workbook (`4 งบการเงิน <client>.xlsx`) for a client.

This is skill 5 (audit-workpaper). It writes the *structure* of the WP before the numbers
and judgment exist — a formula-linked งบการเงิน template, a Mapping (account → FS caption)
sheet, an adjusted-TB layout, an empty adjusting-entries grid, and a tax shell — so that
once the auditor classifies each account in `Mapping!H` and adjusts the TB, the financial
statements recompute themselves ("approach A"). It NEVER invents numbers or classifications.

The shared structure this writes (and skill 5.5 `audit-financials` later reads) is specified
in `docs/financials-contract.md`. v1 covers บจ. going-concern only.

Usage (from repository root):
    uv run .claude/skills/audit-workpaper/scripts/scaffold_workpaper.py \
        "PATH/TO/6_ผลจากสกิล/<client>/CONTEXT.md" [--client-tb PATH [--tb-sheet NAME]]

Prints JSON: output path + warnings (every judgment cell left blank, every assumption made).
"""
import argparse
import json
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

FS = "งบการเงิน"          # financial-statements sheet name
MAP = "Mapping"
TB = "TB"
AJE = "ปรับปรุง"
TAX = "ภาษีเงินได้"
CAPREF = "_captions"      # hidden helper sheet holding the controlled vocabulary
NUMFMT = "#,##0.00;(#,##0.00)"

# ── Controlled caption vocabulary (docs/financials-contract.md §4) ────────────────
# side = which Mapping column carries the natural balance: "C" = ยอด เดบิต, "D" = ยอด เครดิต.
CUR_ASSETS = [
    ("เงินสดและรายการเทียบเท่าเงินสด", "C"),
    ("ลูกหนี้การค้าและลูกหนี้หมุนเวียนอื่น", "C"),
    ("เงินให้กู้ยืมระยะสั้น", "C"),
    ("สินค้าคงเหลือ", "C"),
    ("สินทรัพย์หมุนเวียนอื่น", "C"),
]
NONCUR_ASSETS = [
    ("ที่ดิน อาคารและอุปกรณ์", "C"),
    ("สินทรัพย์ไม่มีตัวตน", "C"),
    ("สินทรัพย์ไม่หมุนเวียนอื่น", "C"),
]
CUR_LIAB = [
    ("เจ้าหนี้การค้าและเจ้าหนี้หมุนเวียนอื่น", "D"),
    ("เงินกู้ยืมระยะสั้น", "D"),
    ("ภาษีเงินได้ค้างจ่าย", "D"),
]
NONCUR_LIAB = [
    ("เงินกู้ยืมระยะยาว", "D"),
    ("หนี้สินไม่หมุนเวียนอื่น", "D"),
]
SHARE_CAP = ("ทุนที่ออกและชำระแล้ว", "D")
RETAINED = ("กำไร(ขาดทุน)สะสม", "D")
REVENUE = [
    ("รายได้จากการขายหรือบริการ", "D"),
    ("รายได้อื่น", "D"),
]
EXPENSES = [
    ("ต้นทุนขายหรือบริการ", "C"),
    ("ค่าใช้จ่ายในการขายและบริหาร", "C"),
    ("ต้นทุนทางการเงิน", "C"),
]

# All captions that get a live SUMIF note-detail row (order defines note numbering).
ALL_CAPTIONS = (
    [c for c, _ in CUR_ASSETS + NONCUR_ASSETS]
    + [c for c, _ in CUR_LIAB + NONCUR_LIAB]
    + [SHARE_CAP[0], RETAINED[0]]
    + [c for c, _ in REVENUE + EXPENSES]
)
SIDE = dict(CUR_ASSETS + NONCUR_ASSETS + CUR_LIAB + NONCUR_LIAB + [SHARE_CAP, RETAINED] + REVENUE + EXPENSES)

# Fixed accounting-policy boilerplate (docs/financials-contract.md §7).
POLICY_BASIS = (
    "งบการเงินนี้จัดขึ้นตามมาตรฐานการรายงานทางการเงินสำหรับกิจการที่ไม่มีส่วนได้เสียสาธารณะ "
    "ที่ออกโดยสภาวิชาชีพบัญชี และวิธีปฏิบัติทางการบัญชีที่รับรองทั่วไปในประเทศไทย "
    "แสดงรายการตามประกาศของกรมพัฒนาธุรกิจการค้า ลงวันที่ 27 ตุลาคม พ.ศ.2566 "
    "เรื่องกำหนดรายการย่อที่ต้องมีในงบการเงิน พ.ศ.2566 "
    "งบการเงินนี้จัดทำบัญชีตามกฎหมายเป็นภาษาไทยและเป็นเงินบาท ใช้เกณฑ์ราคาทุนเดิมในการวัดมูลค่า"
)
POLICY_ITEMS = [
    ("เงินสดและรายการเทียบเท่าเงินสด",
     "เงินสดและรายการเทียบเท่าเงินสด ประกอบด้วย เงินสดในมือและเงินฝากธนาคารกระแสรายวันและออมทรัพย์ "
     "ที่ครบกำหนดไม่เกิน 3 เดือน ไม่รวมรายการเงินฝากธนาคารที่ภาระค้ำประกัน"),
    ("สินค้าคงเหลือ",
     "สินค้าคงเหลือแสดงด้วยราคาทุนหรือมูลค่าสุทธิที่จะได้รับแล้วแต่ราคาใดจะต่ำกว่า"),
    ("ที่ดิน อาคารและอุปกรณ์",
     "ที่ดินบันทึกด้วยราคาทุนเริ่มแรก ที่ดินไม่มีการคิดค่าเสื่อมราคาเนื่องจากมีอายุการใช้งานไม่จำกัด "
     "อาคารและอุปกรณ์แสดงด้วยราคาทุนหลังหักค่าเสื่อมราคาสะสม ค่าเสื่อมราคาคำนวณโดยวิธีเส้นตรง"),
    ("การรับรู้รายได้และค่าใช้จ่าย",
     "⚠ ⟨FILL: revenue-recognition first sentence differs by business type — trading vs service⟩"),
    ("ภาษีเงินได้นิติบุคคล",
     "การคำนวณภาษีเงินได้นิติบุคคลตามหลักเกณฑ์ประมวลรัษฎากรในอัตราที่กฎหมายกำหนด "
     "ของกำไรสุทธิก่อนค่าใช้จ่ายภาษีเงินได้"),
]


def fail(msg: str) -> None:
    print(json.dumps({"ok": False, "error": msg}, ensure_ascii=False))
    sys.exit(1)


def parse_context(path: Path) -> dict:
    profile = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        m = re.match(r"\|\s*([a-z_0-9]+)\s*\|\s*(.*?)\s*\|", line)
        if m:
            profile[m.group(1)] = m.group(2)
    return profile


def context_verdict(path: Path) -> str | None:
    """Return the §2 validation Verdict line's text, or None if absent."""
    m = re.search(r"Verdict:\s*(.+)", path.read_text(encoding="utf-8"))
    return m.group(1).strip(" *") if m else None


def clean(v):
    if not v or str(v).strip() in {"—", "-"} or "⚠" in str(v) or "⟨FILL" in str(v):
        return None
    return re.sub(r"\s{2,}", " ", str(v).strip())


def period_str(profile: dict):
    """Return (label, filename_tag) from period_end dd/mm/bbbb."""
    pe = profile.get("period_end", "")
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", pe)
    if not m:
        return "⚠ ⟨FILL: period_end⟩", "period"
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    months = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
              "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
    label = f"{d} {months[mo]} พ.ศ. {y}"
    return label, f"{d:02d}{mo:02d}{str(y)[-2:]}"


def be_year(period: str):
    m = re.search(r"/(\d{4})\s*$", period or "")
    return int(m.group(1)) if m else None


# ── client TB import (best-effort) ────────────────────────────────────────────────
# Header text that identifies the real *closing*-balance column pair, as opposed to
# opening (ยกมา/ต้นงวด) or movement (เคลื่อนไหว/ยอด) pairs that some client exports also
# carry (e.g. a ยกมา/เคลื่อนไหว/ยกไป 8-column layout, or a ต้นงวด/ยอด/ปลายงวด 3-pair layout).
_CLOSE_HDR_RE = re.compile(r"(ปลายงวด|ยกไป|คงเหลือ)", re.I)
_DR_HDR_RE = re.compile(r"(เดบิต|เดบิท|debit|\bdr\b)", re.I)
_CR_HDR_RE = re.compile(r"(เครดิต|credit|\bcr\b)", re.I)


def find_closing_pair(rows: list, scan_rows: int = 20):
    """Locate the real closing-balance Dr/Cr column pair by header text.

    Returns (dr_col, cr_col, dr_header_text, cr_header_text) using 0-based column
    indices, or None if no such header pair is found (caller falls back to the
    last-two-numeric-cells heuristic). Handles two real layouts:
      - single-row header where the pair is named together, e.g. "ยกไป Dr" / "ยกไป Cr"
        sitting in adjacent cells;
      - a merged two-row header where a group label ("ยอดคงเหลือ" / "ปลายงวด") sits above
        a เดบิต/เครดิต sub-header row, in the SAME two columns as the group label.
    """
    limit = min(scan_rows, len(rows))
    for i in range(limit):
        row = rows[i]
        if row is None:
            continue
        for j, cell in enumerate(row):
            if not isinstance(cell, str) or not _CLOSE_HDR_RE.search(cell):
                continue
            # case A: this cell itself names Dr, and the very next cell names Cr
            if _DR_HDR_RE.search(cell) and j + 1 < len(row):
                nxt = row[j + 1]
                if isinstance(nxt, str) and _CR_HDR_RE.search(nxt):
                    return j, j + 1, cell.strip(), nxt.strip()
            # case B: this is a merged group-label cell; the เดบิต/เครดิต sub-header sits
            # one or two rows below, in the same two columns (merged-cell layout)
            for k in range(i + 1, min(i + 3, len(rows))):
                subrow = rows[k]
                if subrow is None or j + 1 >= len(subrow):
                    continue
                c_dr, c_cr = subrow[j], subrow[j + 1]
                if (isinstance(c_dr, str) and _DR_HDR_RE.search(c_dr)
                        and isinstance(c_cr, str) and _CR_HDR_RE.search(c_cr)):
                    return j, j + 1, c_dr.strip(), c_cr.strip()
    return None


def import_client_tb(path: Path, sheet_name: str | None, warnings: list) -> list[tuple]:
    """Return list of (account_code, account_name, ending_dr, ending_cr).

    Client GL exports vary wildly (docs/financials-contract.md §2/§8). The closing Dr/Cr
    column pair is located by header text first (ปลายงวด/ยกไป/คงเหลือ, paired with
    เดบิต/เครดิต or Dr/Cr); only when no such header is found does it fall back to
    assuming the last two numeric cells in a row are the closing pair. On any doubt it
    warns and returns what it could parse so the human can correct the TB/Mapping by
    hand — it never fabricates balances.
    """
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as e:  # noqa: BLE001
        warnings.append(f"could not open --client-tb ({e}); produced an empty TB layout to fill by hand")
        return []
    ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb[wb.sheetnames[0]]
    if sheet_name and sheet_name not in wb.sheetnames:
        warnings.append(f"--tb-sheet '{sheet_name}' not found; used first sheet '{ws.title}'")

    rows = list(ws.iter_rows(values_only=True))
    # find a header row mentioning an account-code-ish column
    header_idx = None
    for i, r in enumerate(rows[:15]):
        joined = " ".join(str(c) for c in r if c is not None)
        if re.search(r"(เลขที่บัญชี|รหัส|ผังบัญชี|บัญชี|account)", joined, re.I):
            header_idx = i
            break
    if header_idx is None:
        warnings.append("could not locate a header row in --client-tb; produced an empty TB to fill by hand")
        return []

    closing = find_closing_pair(rows)
    dr_idx = cr_idx = dr_hdr = cr_hdr = None
    if closing:
        dr_idx, cr_idx, dr_hdr, cr_hdr = closing

    out = []
    seen = 0
    for r in rows[header_idx + 1:]:
        if r is None or all(c is None for c in r):
            continue
        seen += 1
        # account code: extract the LEADING numeric run from cell 0, not a whole-cell
        # match — client cells combine code+suffix ("111101 - CSH001") or code+name
        # ("1001000: เงินสด") in a single cell, and a whole-string match drops those rows.
        raw0 = str(r[0]).strip() if r[0] is not None else ""
        code_m = re.match(r"^(\d[\d\-\.]*)", raw0)
        if not code_m:
            continue
        code = code_m.group(1)
        remainder0 = raw0[code_m.end():].strip(" \t-:.").strip()
        name = ""
        for c in r[1:4]:
            if isinstance(c, str) and c.strip() and not re.match(r"^[\d,\.\-]+$", c.strip()):
                name = c.strip()
                break
        if not name and remainder0 and not re.match(r"^[\d,\.\-]+$", remainder0):
            # code cell combined code+name ("1001000: เงินสด") and no other column had a name
            name = remainder0
        if dr_idx is not None and dr_idx < len(r) and cr_idx < len(r):
            d_val, c_val = r[dr_idx], r[cr_idx]
            dr = float(d_val) if isinstance(d_val, (int, float)) else 0.0
            cr = float(c_val) if isinstance(c_val, (int, float)) else 0.0
        else:
            # fallback heuristic: the last two numeric cells are the ending Dr / Cr pair
            nums = [c for c in r if isinstance(c, (int, float))]
            dr = cr = 0.0
            if len(nums) >= 2:
                dr, cr = float(nums[-2]), float(nums[-1])
            elif len(nums) == 1:
                dr = float(nums[0])
        out.append((code, name, dr, cr))

    imported = len(out)
    if imported == 0 and seen > 0:
        warnings.append(f"parsed 0 accounts from {seen} non-blank source row(s) below the header — this is "
                        "almost certainly an account-code detection failure, NOT a genuinely-empty TB; "
                        "inspect --client-tb's layout by hand before trusting an empty TB")
    elif imported < seen:
        warnings.append(f"saw {seen} non-blank source row(s) below the header but only {imported} yielded a "
                        f"leading numeric account code ({seen - imported} skipped) — verify none of those "
                        "skipped rows are real accounts")

    if not out:
        warnings.append("no account rows parsed from --client-tb; produced an empty TB to fill by hand")
        return out

    total_dr = sum(d for _, _, d, _ in out)
    total_cr = sum(c for _, _, _, c in out)
    imbalance = total_dr - total_cr
    balanced = abs(imbalance) < 0.01
    if closing:
        msg = (f"imported {imported} accounts from --client-tb using header-matched closing-balance columns "
              f"('{dr_hdr}' / '{cr_hdr}') — Σdebit={total_dr:,.2f} Σcredit={total_cr:,.2f}")
    else:
        msg = (f"imported {imported} accounts from --client-tb — no closing-balance header "
              "(ปลายงวด/ยกไป/คงเหลือ) found; fell back to the 'last two numeric cells' heuristic "
              f"(client exports vary; VERIFY the Dr/Cr columns landed correctly) — "
              f"Σdebit={total_dr:,.2f} Σcredit={total_cr:,.2f}")
    if balanced:
        msg += " — balances to zero"
    else:
        msg += (f" — DOES NOT BALANCE (Σdebit-Σcredit = {imbalance:,.2f}); the Dr/Cr columns are "
               "very likely misaligned or rows are missing — fix the TB by hand before relying on it")
    warnings.append(msg)
    return out


# ── styling helpers ────────────────────────────────────────────────────────────────
BOLD = Font(bold=True)
HEAD = Font(bold=True, size=13)
THIN = Side(style="thin", color="BBBBBB")
TOPBORDER = Border(top=THIN)
TOPBOT = Border(top=THIN, bottom=Side(style="double", color="888888"))


def money(ws, cell):
    ws[cell].number_format = NUMFMT


def sumif(cap: str, col: str) -> str:
    return f'=SUMIF({MAP}!$H:$H,"{cap}",{MAP}!${col}:${col})'


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("context", help="path to the client's CONTEXT.md")
    ap.add_argument("--client-tb", help="path to the client's TB/GL export .xlsx to pre-load accounts")
    ap.add_argument("--tb-sheet", help="sheet name inside --client-tb (default: first sheet)")
    args = ap.parse_args()

    ctx_path = Path(args.context)
    if not ctx_path.exists():
        fail(f"CONTEXT.md not found: {ctx_path}")
    profile = parse_context(ctx_path)
    warnings: list[str] = []

    # Refuse to render off a broken/unverified CONTEXT.md — same gate as
    # audit-cover-report's render_cover_report.py and audit-planning's render_planning.py.
    verdict = context_verdict(ctx_path)
    if verdict and re.search(r"REVIEW|AMBIGUOUS|JOB NOT IN DB", verdict, re.I):
        fail(f"CONTEXT §2 validation verdict is unresolved ({verdict!r}) — do NOT auto-generate; "
             "resolve the job-number cross-check before running this skill.")

    company_raw = clean(profile.get("company_legal_name")) or clean(profile.get("company_name"))
    tax_id = clean(profile.get("tax_id"))
    period_end_raw = clean(profile.get("period_end"))
    required = {"company_name": company_raw, "tax_id": tax_id, "period_end": period_end_raw}
    missing = [k for k, v in required.items() if not v]
    if missing:
        fail(f"unresolved required field(s) in CONTEXT (⚠ or missing): {missing}. Resolve before rendering.")

    entity = profile.get("entity_type", "")
    if "หจก" in entity or "ห้างหุ้นส่วน" in entity:
        warnings.append("entity_type looks like หจก. (partnership) — v1 scaffolds บจ. equity only; "
                        "the equity section needs the partnership variant (see contract §6)")

    company = company_raw
    p_label, p_tag = period_str(profile)
    cy = be_year(profile.get("period_end", "")) or "⚠"
    py = be_year(profile.get("prior_period_end", "")) or (cy - 1 if isinstance(cy, int) else "⚠")

    tb_rows = import_client_tb(Path(args.client_tb), args.tb_sheet, warnings) if args.client_tb else []

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── hidden placeholder + caption vocabulary ────────────────────────────────────
    wb.create_sheet("0000").sheet_state = "veryHidden"
    cap_ws = wb.create_sheet(CAPREF)
    cap_ws.sheet_state = "hidden"
    cap_ws["A1"] = "controlled FS caption vocabulary — do not edit (docs/financials-contract.md §4)"
    for i, cap in enumerate(ALL_CAPTIONS, start=2):
        cap_ws[f"A{i}"] = cap
    cap_range = f"{CAPREF}!$A$2:$A${len(ALL_CAPTIONS) + 1}"

    # ── TB sheet ───────────────────────────────────────────────────────────────────
    tb = wb.create_sheet(TB)
    tb["A1"] = company
    tb["A2"] = f"งบทดลอง (ปรับปรุงแล้ว) ณ วันที่ {p_label}"
    hdr = ["เลขที่บัญชี", "ชื่อบัญชี", "ยอดยกมา เดบิต", "ยอดยกมา เครดิต",
           "ปรับปรุง เดบิต", "ปรับปรุง เครดิต", "คงเหลือ เดบิต", "คงเหลือ เครดิต"]
    for j, h in enumerate(hdr, start=1):
        c = tb.cell(row=3, column=j, value=h)
        c.font = BOLD
    r = 4
    for code, name, dr, cr in tb_rows:
        tb.cell(row=r, column=1, value=code)
        tb.cell(row=r, column=2, value=name)
        # ending balances go into opening (C/D); adjustments (E/F) start empty; G/H = C+E-D-F
        tb.cell(row=r, column=3, value=dr).number_format = NUMFMT
        tb.cell(row=r, column=4, value=cr).number_format = NUMFMT
        tb.cell(row=r, column=7, value=f"=C{r}+E{r}-D{r}-F{r}").number_format = NUMFMT
        tb.cell(row=r, column=8, value=f"=D{r}+F{r}-C{r}-E{r}").number_format = NUMFMT
        r += 1
    tb.column_dimensions["A"].width = 14
    tb.column_dimensions["B"].width = 34
    for col in "CDEFGH":
        tb.column_dimensions[col].width = 15

    # ── Mapping sheet ──────────────────────────────────────────────────────────────
    mp = wb.create_sheet(MAP)
    mp["A1"] = company
    mp["A2"] = f"กระดาษทำการจัดประเภทผังบัญชี (Mapping) ปี {cy} เทียบ {py}"
    mp["A3"] = "⚠ กรอกคอลัมน์ H (รายการในงบการเงิน) ให้ทุกบัญชี — เลือกจากรายการมาตรฐาน (dropdown)"
    mp["A3"].font = Font(color="9C6210")
    mhdr = ["ผังบัญชี", "ชื่อบัญชี", "ยอด เดบิต", "ยอด เครดิต",
            "ยอดปีก่อน เดบิต", "ยอดปีก่อน เครดิต", "หมวดตามหมายเหตุ", "รายการในงบการเงิน"]
    for j, h in enumerate(mhdr, start=1):
        mp.cell(row=5, column=j, value=h).font = BOLD
    r = 6
    for code, name, _dr, _cr in tb_rows:
        mp.cell(row=r, column=1, value=code)
        mp.cell(row=r, column=2, value=name)
        mp.cell(row=r, column=3, value=f'=SUMIF({TB}!$A:$A,A{r},{TB}!$G:$G)').number_format = NUMFMT
        mp.cell(row=r, column=4, value=f'=SUMIF({TB}!$A:$A,A{r},{TB}!$H:$H)').number_format = NUMFMT
        r += 1
    map_last = max(r - 1, 6)
    # dropdown on H restricting to the controlled vocabulary
    dv = DataValidation(type="list", formula1=f"={cap_range}", allow_blank=True)
    dv.error = "เลือกจากรายการมาตรฐาน (NPAE) เท่านั้น"
    dv.errorTitle = "caption ไม่ถูกต้อง"
    mp.add_data_validation(dv)
    dv.add(f"H6:H{map_last + 50}")
    mp.column_dimensions["A"].width = 14
    mp.column_dimensions["B"].width = 32
    for col in "CDEF":
        mp.column_dimensions[col].width = 14
    mp.column_dimensions["G"].width = 30
    mp.column_dimensions["H"].width = 32

    # ── งบการเงิน sheet: layout plan first, then render ─────────────────────────────
    # Build an ordered list of row specs; the row number is its position. Statement lines
    # reference note-detail rows, so we resolve positions in a plan pass then render.
    plan: list[dict] = []

    def add(kind, **kw):
        plan.append({"kind": kind, **kw})

    # --- Balance sheet ---
    add("title", text=company)
    add("title", text="งบแสดงฐานะการเงิน")
    add("title", text=f"ณ วันที่ {p_label}")
    add("colhdr")
    add("section", text="สินทรัพย์")
    add("subhdr", text="สินทรัพย์หมุนเวียน")
    for cap, _ in CUR_ASSETS:
        add("cap", cap=cap)
    add("subtotal", text="รวมสินทรัพย์หมุนเวียน", group="cur_assets")
    add("subhdr", text="สินทรัพย์ไม่หมุนเวียน")
    for cap, _ in NONCUR_ASSETS:
        add("cap", cap=cap)
    add("subtotal", text="รวมสินทรัพย์ไม่หมุนเวียน", group="noncur_assets")
    add("total", text="รวมสินทรัพย์", group="assets", name="FS_TOTAL_ASSETS")
    add("blank")
    add("section", text="หนี้สินและส่วนของผู้ถือหุ้น")
    add("subhdr", text="หนี้สินหมุนเวียน")
    for cap, _ in CUR_LIAB:
        add("cap", cap=cap)
    add("subtotal", text="รวมหนี้สินหมุนเวียน", group="cur_liab")
    add("subhdr", text="หนี้สินไม่หมุนเวียน")
    for cap, _ in NONCUR_LIAB:
        add("cap", cap=cap)
    add("subtotal", text="รวมหนี้สินไม่หมุนเวียน", group="noncur_liab")
    add("total", text="รวมหนี้สิน", group="liab")
    add("subhdr", text="ส่วนของผู้ถือหุ้น")
    add("cap", cap=SHARE_CAP[0])
    add("re_bs")  # retained earnings (from equity statement closing)
    add("subtotal", text="รวมส่วนของผู้ถือหุ้น", group="equity", name="FS_TOTAL_EQUITY")
    add("total", text="รวมหนี้สินและส่วนของผู้ถือหุ้น", group="liabeq", name="FS_TOTAL_LIAB_EQUITY")
    add("checkrow")  # BS-balances tie-out flag
    add("blank")

    # --- Income statement ---
    add("title", text=company)
    add("title", text="งบกำไรขาดทุน")
    add("title", text=f"สำหรับปีสิ้นสุดวันที่ {p_label}")
    add("colhdr")
    add("section", text="รายได้")
    for cap, _ in REVENUE:
        add("cap", cap=cap)
    add("subtotal", text="รวมรายได้", group="revenue")
    add("section", text="ค่าใช้จ่าย")
    for cap, _ in EXPENSES:
        add("cap", cap=cap)
    add("subtotal", text="รวมค่าใช้จ่าย", group="expenses")
    add("pbt", text="กำไร(ขาดทุน)ก่อนภาษีเงินได้", name="TAX_NET_PROFIT")
    add("taxline", text="ค่าใช้จ่ายภาษีเงินได้")
    add("netprofit", text="กำไร(ขาดทุน)สุทธิ", name="FS_NET_PROFIT")
    add("blank")

    # --- Statement of changes in equity (minimal) ---
    add("title", text="งบแสดงการเปลี่ยนแปลงส่วนของผู้ถือหุ้น")
    add("eq_open", text="ยอดคงเหลือต้นปี — กำไร(ขาดทุน)สะสม")
    add("eq_profit", text="กำไร(ขาดทุน)สุทธิสำหรับปี")
    add("eq_close", text="ยอดคงเหลือปลายปี — กำไร(ขาดทุน)สะสม")
    add("blank")

    # --- Notes ---
    add("title", text="หมายเหตุประกอบงบการเงิน")
    add("note_general")
    add("note_basis")
    add("note_policy")
    add("section", text="รายละเอียดประกอบ (Note detail — รายบรรทัดเพิ่มโดย audit-financials)")
    add("colhdr")
    for cap in ALL_CAPTIONS:
        add("note_detail", cap=cap)

    # resolve row numbers
    for i, spec in enumerate(plan, start=1):
        spec["row"] = i
    cap_note_row = {s["cap"]: s["row"] for s in plan if s["kind"] == "note_detail"}
    group_rows: dict[str, list[int]] = {}
    for s in plan:
        if s["kind"] == "cap":
            grp = _group_of(s["cap"])
            group_rows.setdefault(grp, []).append(s["row"])
    # named single rows resolved during render
    named: dict[str, int] = {}
    row_of = {"pbt": None, "netprofit": None, "eq_close": None, "re_bs": None,
              "eq_open": None, "eq_profit": None, "taxline": None}
    for s in plan:
        if s["kind"] in row_of:
            row_of[s["kind"]] = s["row"]

    ws = wb.create_sheet(FS)  # final sheet order is fixed at the end

    def rng(group):
        rows = sorted(group_rows.get(group, []))
        if not rows:
            return None
        return f"E{rows[0]}:E{rows[-1]}", f"G{rows[0]}:G{rows[-1]}"

    # composite subtotal groups (assets = cur+noncur, etc.)
    composite = {
        "assets": ["cur_assets", "noncur_assets"],
        "liab": ["cur_liab", "noncur_liab"],
        "liabeq": ["cur_liab", "noncur_liab", "equity_all"],
    }

    for s in plan:
        row = s["row"]
        k = s["kind"]
        a = ws.cell(row=row, column=1)
        if k == "title":
            a.value = s["text"]
            a.font = HEAD if s["text"] in (company,) else BOLD
            a.alignment = Alignment(horizontal="left")
        elif k == "colhdr":
            ws.cell(row=row, column=3, value="หมายเหตุ").font = BOLD
            ws.cell(row=row, column=5, value=cy).font = BOLD
            ws.cell(row=row, column=7, value=py).font = BOLD
        elif k in ("section", "subhdr"):
            a.value = s["text"]
            a.font = BOLD
        elif k == "cap":
            cap = s["cap"]
            a.value = cap
            nr = cap_note_row[cap]
            ws.cell(row=row, column=3, value=cap_note_index(cap)).alignment = Alignment(horizontal="center")
            _put(ws, f"E{row}", f"=+E{nr}")
            _put(ws, f"G{row}", f"=+G{nr}")
        elif k == "re_bs":
            a.value = RETAINED[0]
            _put(ws, f"E{row}", f"=+E{row_of['eq_close']}")
            _put(ws, f"G{row}", 0)
        elif k in ("subtotal", "total"):
            a.value = s["text"]
            a.font = BOLD
            groups = composite.get(s["group"], [s["group"]])
            # equity subtotal must include share cap row + re_bs row
            if s["group"] == "equity":
                erows = sorted(group_rows.get("equity", []) + [row_of["re_bs"]])
                _put(ws, f"E{row}", f"=SUM(E{erows[0]}:E{erows[-1]})")
                _put(ws, f"G{row}", f"=SUM(G{erows[0]}:G{erows[-1]})")
                group_rows["equity_all"] = erows
            else:
                parts_e, parts_g = [], []
                for g in groups:
                    rr = rng(g)
                    if rr:
                        parts_e.append(f"SUM({rr[0]})")
                        parts_g.append(f"SUM({rr[1]})")
                _put(ws, f"E{row}", "=" + "+".join(parts_e) if parts_e else "=0")
                _put(ws, f"G{row}", "=" + "+".join(parts_g) if parts_g else "=0")
            ws[f"E{row}"].border = TOPBORDER
            ws[f"G{row}"].border = TOPBORDER
            if s.get("name"):
                named[s["name"] + "_CY"] = row
                named[s["name"] + "_PY"] = row
        elif k == "pbt":
            a.value = s["text"]
            a.font = BOLD
            rev = rng("revenue")
            exp = rng("expenses")
            _put(ws, f"E{row}", f"=SUM({rev[0]})-SUM({exp[0]})")
            _put(ws, f"G{row}", f"=SUM({rev[1]})-SUM({exp[1]})")
            named["TAX_NET_PROFIT"] = row
        elif k == "taxline":
            a.value = s["text"]
            _put(ws, f"E{row}", "=TAX_EXPENSE")  # defined name → ภาษีเงินได้ computed cell
            _put(ws, f"G{row}", 0)
        elif k == "netprofit":
            a.value = s["text"]
            a.font = BOLD
            _put(ws, f"E{row}", f"=E{row_of['pbt']}-E{row_of['taxline']}")
            _put(ws, f"G{row}", f"=G{row_of['pbt']}-G{row_of['taxline']}")
            ws[f"E{row}"].border = TOPBOT
            ws[f"G{row}"].border = TOPBOT
            named["FS_NET_PROFIT_CY"] = row
            named["FS_NET_PROFIT_PY"] = row
        elif k == "eq_open":
            a.value = s["text"]
            _put(ws, f"E{row}", sumif(RETAINED[0], "D"))  # opening RE = mapped RE-account balance
        elif k == "eq_profit":
            a.value = s["text"]
            _put(ws, f"E{row}", f"=+E{row_of['netprofit']}")
        elif k == "eq_close":
            a.value = s["text"]
            a.font = BOLD
            _put(ws, f"E{row}", f"=E{row_of['eq_open']}+E{row_of['eq_profit']}")
            ws[f"E{row}"].border = TOPBORDER
        elif k == "checkrow":
            a.value = "ตรวจสอบ: สินทรัพย์ = หนี้สิน+ทุน (ต้องเป็น 0)"
            a.font = Font(italic=True, color="9C6210")
            ta = named.get("FS_TOTAL_ASSETS_CY")
            tle = named.get("FS_TOTAL_LIAB_EQUITY_CY")
            if ta and tle:
                _put(ws, f"E{row}", f"=E{ta}-E{tle}")
        elif k == "note_general":
            a.value = "1. ข้อมูลทั่วไป"
            a.font = BOLD
            ws.cell(row=row, column=1)  # detail appended on next rows via note_basis/policy blocks
        elif k == "note_basis":
            a.value = "2. เกณฑ์การจัดทำงบการเงิน"
            a.font = BOLD
        elif k == "note_policy":
            a.value = "3. นโยบายการบัญชีที่สำคัญ"
            a.font = BOLD
        elif k == "note_detail":
            cap = s["cap"]
            a.value = cap
            a.alignment = Alignment(indent=1)
            ws.cell(row=row, column=3, value=cap_note_index(cap)).alignment = Alignment(horizontal="center")
            side = SIDE[cap]
            _put(ws, f"E{row}", sumif(cap, side))
            _put(ws, f"G{row}", sumif(cap, "E" if side == "C" else "F"))
        elif k == "blank":
            pass
        # apply money format to E/G on value rows
        if k in ("cap", "subtotal", "total", "pbt", "taxline", "netprofit", "re_bs",
                 "eq_open", "eq_profit", "eq_close", "note_detail", "checkrow"):
            money(ws, f"E{row}")
            money(ws, f"G{row}")

    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 3
    ws.column_dimensions["C"].width = 8
    for col in ("D", "E", "F", "G"):
        ws.column_dimensions[col].width = 16

    # ── ภาษีเงินได้ (tax shell) ─────────────────────────────────────────────────────
    tx = wb.create_sheet(TAX)
    tx["A1"] = company
    tx["A2"] = f"การคำนวณภาษีเงินได้นิติบุคคล ปี {cy}"
    tx["A2"].font = BOLD
    trows = [
        ("กำไร(ขาดทุน)สุทธิทางบัญชีก่อนภาษี", f"=+'{FS}'!E{row_of['pbt']}", None),
        ("บวก รายจ่ายต้องห้าม", "⚠ ⟨human⟩", "warn"),
        ("หัก รายได้ที่ได้รับยกเว้น", "⚠ ⟨human⟩", "warn"),
        ("กำไร(ขาดทุน)สุทธิทางภาษี", None, "nettax"),
        ("หัก ยกเว้น SME 300,000 บาทแรก", "⚠ ⟨human: ถ้าเข้าเงื่อนไข SME⟩", "warn"),
        ("ฐานภาษี", None, "base"),
        ("อัตราภาษี", 0.15, "rate"),
        ("ภาษีเงินได้คำนวณ", None, "computed"),
        ("หัก ภาษีหัก ณ ที่จ่าย / ภงด.51", "⚠ ⟨human⟩", "warn"),
        ("ภาษีเงินได้ที่ต้องชำระ(ชำระเกิน)", None, "payable"),
    ]
    tr = {}
    r = 4
    for label, val, key in trows:
        tx.cell(row=r, column=1, value=label)
        cell = tx.cell(row=r, column=3)
        if isinstance(val, (int, float)):
            cell.value = val
        elif isinstance(val, str):
            cell.value = val
        if key:
            tr[key] = r
        r += 1
    # computed formulas (best-effort; add-backs are human, so N() treats blanks as 0)
    tx.cell(row=tr["nettax"], column=3,
            value=f"=C4+N(C{tr['nettax']-2})-N(C{tr['nettax']-1})").number_format = NUMFMT
    tx.cell(row=tr["base"], column=3,
            value=f"=MAX(0,C{tr['nettax']}-N(C{tr['base']-1}))").number_format = NUMFMT
    tx.cell(row=tr["computed"], column=3,
            value=f"=ROUND(C{tr['base']}*C{tr['rate']},2)").number_format = NUMFMT
    tx.cell(row=tr["payable"], column=3,
            value=f"=C{tr['computed']}-N(C{tr['payable']-1})").number_format = NUMFMT
    tx["C3"].number_format = NUMFMT
    tx.cell(row=tr["rate"], column=3).number_format = "0%"
    tx.column_dimensions["A"].width = 40
    tx.column_dimensions["C"].width = 18

    # tax expense on IS references the computed tax
    wb.defined_names["TAX_EXPENSE"] = DefinedName("TAX_EXPENSE", attr_text=f"'{TAX}'!$C${tr['computed']}")
    wb.defined_names["TAX_NET_PROFIT"] = DefinedName("TAX_NET_PROFIT", attr_text=f"'{FS}'!$E${row_of['pbt']}")
    wb.defined_names["TAX_PAYABLE"] = DefinedName("TAX_PAYABLE", attr_text=f"'{TAX}'!$C${tr['payable']}")

    # ── ปรับปรุง (AJE grid) ─────────────────────────────────────────────────────────
    aj = wb.create_sheet(AJE)
    aj["A1"] = company
    aj["A2"] = f"รายการปรับปรุง (Adjusting entries) ปี {cy}"
    aj["A2"].font = BOLD
    ajhdr = ["เลขที่", "คำอธิบายรายการ", "Ref", "เดบิต (งบดุล)", "เครดิต (งบดุล)",
             "เดบิต (กำไรขาดทุน)", "เครดิต (กำไรขาดทุน)"]
    for j, h in enumerate(ajhdr, start=1):
        aj.cell(row=3, column=j, value=h).font = BOLD
    aj["A4"] = "⚠ ⟨human: ลงรายการปรับปรุงตามดุลพินิจ — สคริปต์ไม่สร้างรายการเอง⟩"
    aj["A4"].font = Font(color="9C6210")
    aj.column_dimensions["A"].width = 8
    aj.column_dimensions["B"].width = 40
    for col in "DEFG":
        aj.column_dimensions[col].width = 15

    # ── defined names for FS totals ────────────────────────────────────────────────
    for nm, rownum in named.items():
        col = "E" if nm.endswith("_CY") else ("G" if nm.endswith("_PY") else "E")
        wb.defined_names[nm] = DefinedName(nm, attr_text=f"'{FS}'!${col}${rownum}")

    # order sheets: 0000, งบการเงิน, TB, Mapping, ปรับปรุง, ภาษีเงินได้, _captions
    order = ["0000", FS, TB, MAP, AJE, TAX, CAPREF]
    wb._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 99)

    # ── write ──────────────────────────────────────────────────────────────────────
    out_dir = ctx_path.parent / "WP"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"4 งบการเงิน {company} {p_tag}.xlsx"
    wb.save(out_path)

    if not tb_rows:
        warnings.append("no client TB imported — TB and Mapping are empty layouts; paste the client's "
                        "trial balance into TB and the accounts into Mapping (or re-run with --client-tb)")
    warnings.append("Mapping!H (การจัดประเภทแต่ละบัญชีเข้า caption งบ) is a human judgment call — the งบ "
                    "stays zero until it is filled from the dropdown")
    warnings.append("v1 covers บจ. going-concern only; หจก./งบเลิก/ปีแรก need their variants (contract §6)")

    print(json.dumps({"ok": True, "output": str(out_path), "accounts_imported": len(tb_rows),
                      "warnings": warnings}, ensure_ascii=False))


# ── caption grouping / note numbering ───────────────────────────────────────────────
_GROUP = {}
for _c, _ in CUR_ASSETS:
    _GROUP[_c] = "cur_assets"
for _c, _ in NONCUR_ASSETS:
    _GROUP[_c] = "noncur_assets"
for _c, _ in CUR_LIAB:
    _GROUP[_c] = "cur_liab"
for _c, _ in NONCUR_LIAB:
    _GROUP[_c] = "noncur_liab"
_GROUP[SHARE_CAP[0]] = "equity"
_GROUP[RETAINED[0]] = "equity"
for _c, _ in REVENUE:
    _GROUP[_c] = "revenue"
for _c, _ in EXPENSES:
    _GROUP[_c] = "expenses"
_NOTE_INDEX = {cap: i + 4 for i, cap in enumerate(ALL_CAPTIONS)}  # notes start at 4


def _group_of(cap):
    return _GROUP[cap]


def cap_note_index(cap):
    return _NOTE_INDEX.get(cap, "")


def _put(ws, cell, value):
    ws[cell] = value


if __name__ == "__main__":
    main()

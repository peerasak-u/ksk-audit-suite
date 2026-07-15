# /// script
# requires-python = ">=3.12"
# dependencies = ["pypdf", "cryptography"]
# ///
"""Dev-only: regenerate the client-neutral CIT50 (ภ.ง.ด.50) AcroForm template.

Reads ONE real file under `5_ตัวอย่างไฟล์ผลลัพธ์/` (read-only — never writes there),
blanks every text field's value while leaving the AcroForm structure, checkbox/radio
group definitions, and page appearance completely intact, and writes the result to
`.claude/skills/audit-cit50/assets/cit50_template.pdf`.

The source form ([103]) was chosen because it is the simplest real case — no tax
adjustments, no liquidation quirks — so its baked-in checkbox/radio defaults (filing
type, currency, SME reduced-rate election, related-party disclosure) represent the
common case documented in references/cit50-field-map.md. Checkbox/radio fields are
intentionally NOT cleared: they already hold the firm's standard selections, and
render_cit50.py only overrides them for known exceptions (e.g. liquidation).

Not part of a normal run — only re-run this if the government's own CIT50 form
changes (new tax year AcroForm), and re-verify the field map afterward.

Usage (from repository root):
    uv run .claude/skills/audit-cit50/scripts/build_template.py

NOTE (public repo): the source path below points at a real client's folder name
under the firm's private `5_ตัวอย่างไฟล์ผลลัพธ์/` (gitignored, never committed).
Substitute your own local client folder/file names to actually re-run this.
"""
from pathlib import Path

from pypdf import PdfReader, PdfWriter

ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent
SOURCE = ROOT / "5_ตัวอย่างไฟล์ผลลัพธ์" / "S [103] <local client folder>" / "WP" / "3 CIT50 <local client name>.pdf"
OUT = Path(__file__).resolve().parent.parent / "assets" / "cit50_template.pdf"


def main() -> None:
    if not SOURCE.exists():
        raise SystemExit(f"source ground-truth file not found: {SOURCE}")
    reader = PdfReader(SOURCE)
    writer = PdfWriter()
    writer.append(reader)

    fields = writer.get_fields()
    clear = {name: "" for name, f in fields.items() if str(f.get("/FT")) == "/Tx" and f.get("/V")}
    for page in writer.pages:
        writer.update_page_form_field_values(page, clear, auto_regenerate=False)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "wb") as fh:
        writer.write(fh)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    main()

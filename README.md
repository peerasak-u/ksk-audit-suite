# KSK Audit Suite

Six [Claude Code Agent Skills](https://docs.claude.com/en/docs/claude-code/skills) that implement
a Thai accounting firm's audit-document pipeline: from a raw client folder to a filled cover
page, auditor's report, planning workbook, financial-statement working paper, and corporate
income tax return (CIT50) — with a status dashboard tying them all together.

Every skill is **deterministic, template-driven, and refuses to invent data**. An LLM only reads
source documents and decides *which* value goes where; a plain Python script (run via `uv run`)
is what actually writes the `.docx`/`.xlsx`/`.pdf` file. If a required fact isn't in the firm's
own databases or the client's documents, the skill leaves it flagged (`⚠`) for a human instead
of guessing.

> **This repo ships no client data.** The sample client folders, ground-truth output folders,
> and the firm's database exports are all gitignored and were never committed — see
> [Required inputs](#required-inputs) for what you need to supply yourself.

> **Built for one specific firm's real templates and database schema.** The `.docx`/`.xlsx`
> template assets under each skill's `assets/` folder are tokenized copies of one firm's actual
> working papers. To use this with a different firm, you will need to rebuild those templates
> (each skill has a `scripts/build_template*.py` dev script for that) and adjust the three
> `Database *.xlsx/.csv` schemas each skill reads.

## The pipeline

```
Phase A — pre-judgment, from CONTEXT (auto)
  1  audit-ingest        client folder → CONTEXT.md + WP/ skeleton      [HUMAN: confirm CONTEXT]
  2  audit-cover-report  CONTEXT → 2 หน้ารายงาน.docx + ใบปะหน้างบการเงิน.docx
  3  audit-planning      CONTEXT → 1 Planning.xlsx
  5  audit-workpaper     CONTEXT → 4 งบการเงิน.xlsx (formula-linked scaffold)

[HUMAN GATE] auditor opens '4 งบการเงิน' in Excel: adjusts the TB, classifies every account,
             posts adjusting entries, writes the risk narrative, and saves.

Phase B — post-judgment (after the WP is adjusted & saved)
  5.5 audit-financials   QA gate + expand note detail → '4 งบการเงิน (final).xlsx'
  4   audit-cit50        pulls tax numbers from the adjusted งบ → 3 CIT50.pdf
```

`audit-orchestrate` (skill 6) is read-only: point it at a client and it reports which stage the
job is at, which of the 6 targets exist, format-QAs anything already produced, and tells you the
single next action.

| # | skill | produces | trust model |
|---|---|---|---|
| 1 | `audit-ingest` | `CONTEXT.md` + `WP/` skeleton | auto-extract + human confirms |
| 2 | `audit-cover-report` | ใบปะหน้า + หน้ารายงาน (`.docx`) | fully deterministic |
| 3 | `audit-planning` | `1 Planning.xlsx` | deterministic fields + one required judgment call (business type) |
| 4 | `audit-cit50` | `3 CIT50.pdf` (925-field AcroForm) | fully deterministic, numbers sourced from the adjusted งบ |
| 5 | `audit-workpaper` | `4 งบการเงิน.xlsx` scaffold | deterministic scaffold; TB import best-effort |
| 5.5 | `audit-financials` | finalized งบ + note detail | QA gate, always human-reviewed |
| 6 | `audit-orchestrate` | status report | read-only dashboard, no writes |

Design docs: [`docs/OVERVIEW.md`](docs/OVERVIEW.md) (how the pipeline was derived from the real
ground-truth files) and [`docs/financials-contract.md`](docs/financials-contract.md) (the shared
contract between skills 5 and 5.5).

## Installation

### Option A — clone the whole repo as your project

```bash
git clone https://github.com/peerasak-u/ksk-audit-suite.git
cd ksk-audit-suite
```

Claude Code auto-discovers every skill under `.claude/skills/` in the current project. Open the
folder with Claude Code and the 6 skills above are immediately available.

### Option B — copy the skills into an existing project

```bash
cp -r ksk-audit-suite/.claude/skills/audit-* your-project/.claude/skills/
```

Copy only the skills you need — each is self-contained (its own `SKILL.md`, `scripts/`,
`references/`, and `assets/`).

### Option C — install individual `.skill` packages

Each skill is also available as a standalone `<skill-name>.skill` file (a zip archive, the
format Claude Code / claude.ai accept for one-off skill installs). Build them locally:

```bash
./scripts/package_skills.sh
```

This writes `dist/skills/audit-ingest.skill`, `dist/skills/audit-cover-report.skill`, etc. — one
per skill, ready to hand to someone or upload via Claude Code's skill installer.

## Required inputs

None of these are in this repo. You need your own copies, laid out like this at the repo root:

```
Database งาน.xlsx                    # job registry — company, tax id, dates, fees (PK = job number)
Database ข้อมูลผู้สอบ.csv.xlsx        # auditor registry — name, CPA/TA license, office
Database เรทคิดเงิน Audit.csv        # fee rate card
4_ตัวอย่างไฟล์ลูกค้า/<job>/           # one folder per client — raw source documents
5_ตัวอย่างไฟล์ผลลัพธ์/<job>/          # (optional) ground-truth prior output, only needed to
                                       # regenerate template assets via build_template*.py
```

`audit-ingest` (skill 1) is the only skill that reads the `Database *` files and `4_.../`
directly; every other skill downstream only ever reads the `CONTEXT.md` that ingest produces.

## Quick start

```
"ingest client folder 4_.../S [123] Example Co 311268"
   → runs audit-ingest, writes 6_.../S [123].../CONTEXT.md

"generate the cover page and report for this client"
   → runs audit-cover-report

"make the planning file"
   → runs audit-planning (will ask you to confirm business-type if it can't be auto-detected)

"scaffold the workpaper"
   → runs audit-workpaper

"what's the status of this job?"
   → runs audit-orchestrate, tells you the next step
```

## Requirements

- [Claude Code](https://docs.claude.com/en/docs/claude-code) or another agent harness that
  supports Agent Skills
- [`uv`](https://docs.astral.sh/uv/) — every script declares its own dependencies inline
  (PEP 723) and `uv run` installs them automatically on first use; no separate `pip install`
  step
- Python ≥ 3.12 (uv will fetch this if you don't have it)

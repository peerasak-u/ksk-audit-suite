---
name: audit-orchestrate
description: Drive a KSK audit client through the whole pipeline (skills 1→5.5) and show where it stands. Use when asked to "run the full pipeline", "orchestrate a client", "what's the status of this job", "what's next for this client", "QA the whole job", or "do the whole audit for a client". Reports the pipeline stage, which of the 5 deliverables exist, the next skill to run, the human gates that block progress, and a format-QA check of each produced file against its locked format contract (plus a cross-check vs the ground-truth folder 5). This skill coordinates the other skills; it does not re-implement their work. This is skill 6 of the pipeline.
allowed-tools: Bash, Read
---

Coordinate the five per-client deliverables that the other skills produce (they chain through
a shared `CONTEXT.md`; see `docs/OVERVIEW.md`). This skill is a **dashboard + runbook**: it
reports where a client stands and what to do next, and format-QAs what has been produced. It
does **not** re-do any skill's work — you drive each step by invoking that skill.

## The pipeline (order + who does what)

```
Phase A — pre-judgment, from CONTEXT (auto)
  1  audit-ingest       client folder → CONTEXT.md + WP/ skeleton      [HUMAN: confirm CONTEXT]
  2  audit-cover-report CONTEXT → 2 หน้ารายงาน.docx + ใบปะหน้างบการเงิน.docx
  3  audit-planning     CONTEXT → 1 Planning.xlsx  (301/203 narrative = human, grounded in docs)
  5  audit-workpaper    CONTEXT → 4 งบการเงิน.xlsx  (formula-linked scaffold)

[HUMAN GATE] auditor opens '4 งบการเงิน' in Excel: adjust TB, classify every Mapping!H,
             post AJEs, write the 301/203 narratives, and SAVE.

Phase B — post-judgment (after the WP is adjusted & saved)
  5.5 audit-financials  QA gate + expand note detail → '4 งบการเงิน (final).xlsx'
  4   audit-cit50       pull tax numbers from the adjusted งบ → 3 CIT50.pdf
```

> **Ordering note:** `audit-cit50` runs in **phase B**, not before the งบ. It reads the tax
> computation out of the adjusted `4 งบการเงิน` workbook, so the WP must be finished first —
> this differs from the file-number order (3 before 4).

Each skill has its own SKILL.md with the exact command and flags. Invoke them there; this skill
only tells you *which* to run next.

## Step 1: Check status

From the repository root:

```bash
uv run ${CLAUDE_SKILL_ROOT}/scripts/orchestrate.py "PATH/TO/6_ผลจากสกิล/<client>/CONTEXT.md"
# or a client folder, or --all for every client under 6_ผลจากสกิล/
```

Read the JSON:
- `stage` — one of `0-not-ingested`, `1-phase-A-scaffolding`, `2-awaiting-human-WP-adjust`,
  `3-phase-B-finalizing`, `4-complete`.
- `files` — each of the 6 targets: `present`, `path`, `format_ok`, `format_notes`. The
  workpaper's notes include whether it has been opened & saved in Excel (the phase-B gate).
- `format_failures` — any produced file that violates its locked format contract (see below).
- `ground_truth` — the matching `5_ตัวอย่างไฟล์ผลลัพธ์/` folder and which deliverables it has,
  so you can see if something is missing that the firm's real job had.
- `next_actions` — the concrete next step(s): which skill to run, or the human gate to wait on.

## Step 2: Act on `next_actions`

- If a skill is named, invoke that skill (read its SKILL.md for flags — e.g. planning needs
  `--director`, cit50 needs the tax figures). Never fabricate the judgment inputs.
- If the action is a **HUMAN GATE**, stop and tell the user exactly what to do (confirm CONTEXT,
  or adjust+save the WP in Excel). Do not fake past it — phase B genuinely needs the human's
  adjusted numbers.
- Re-run `orchestrate.py` after each step to confirm the stage advanced and the new file passes
  format-QA.

## Step 3: Report

Tell the user, in a few lines:
- the current stage and which deliverables exist (n/6),
- the single next action (or the human gate blocking),
- any `format_failures` to fix,
- anything the ground-truth folder has that we are missing.

## Format-QA — what "format_ok" checks (locked contracts, not content)

The goal is **same format every time, content differs per client** (OVERVIEW §เป้าหมาย). So QA
checks each file against its invariant, not against a ground-truth sample:

| file | format check |
|---|---|
| `1 Planning.xlsx` | the 14 fixed sheets are all present |
| `4 งบการเงิน.xlsx` | the core FS sheets (งบการเงิน, TB, Mapping, ปรับปรุง, ภาษีเงินได้) present; reports whether it has been opened & saved in Excel |
| `2 หน้ารายงาน` / `ใบปะหน้า` (docx) | opens and has non-empty paragraphs |
| `3 CIT50.pdf` | is a รด. AcroForm with ~925 fields |

Deeper validation of the งบ (balances, ties, classification) is `audit-financials`' QA gate,
not this one. This skill is read-only — it never writes a deliverable.

## Done when

- `orchestrate.py` reports `stage: 4-complete` with all 6 targets present and no
  `format_failures`, and the user has been pointed at the final human review vs folder 5.

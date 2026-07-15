#!/usr/bin/env bash
# Package every skill under .claude/skills/ into a standalone <name>.skill zip
# (the format Claude Code / claude.ai accept for one-off skill installs).
#
# Usage (from repository root):
#   ./scripts/package_skills.sh
#
# Writes dist/skills/<name>.skill — one per skill. dist/ is gitignored (build output,
# regenerate any time from .claude/skills/).
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SKILLS_DIR="$ROOT/.claude/skills"
OUT_DIR="$ROOT/dist/skills"

rm -rf "$OUT_DIR"
mkdir -p "$OUT_DIR"

for skill_path in "$SKILLS_DIR"/*/; do
    name="$(basename "$skill_path")"
    [ -f "$skill_path/SKILL.md" ] || { echo "skip $name — no SKILL.md"; continue; }

    out_file="$OUT_DIR/$name.skill"
    # zip root must be the skill's own folder name (matches SKILL.md's own convention),
    # so cd into .claude/skills/ and reference the folder by relative name.
    (cd "$SKILLS_DIR" && zip -rq -X "$out_file" "$name" -x "*__pycache__*" -x "*.pyc")
    echo "built $name.skill ($(du -h "$out_file" | cut -f1))"
done

echo
echo "done: $OUT_DIR"

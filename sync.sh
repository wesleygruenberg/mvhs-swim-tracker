#!/usr/bin/env bash
set -euo pipefail

# Always use the real Google clasp via npx (aliases don't expand in scripts)
CLASP="npx -y @google/clasp@latest"

MODE="${1:-pull}"   # pull | push
MSG="${2:-auto-sync: $(date -Iseconds)}"

echo "Using clasp: $CLASP"

# Sanity checks
command -v git >/dev/null || { echo "git not found"; exit 1; }
[ -f .clasp.json ] || { echo ".clasp.json not found (run in repo root)"; exit 1; }
$CLASP --help >/dev/null

case "$MODE" in
  pull) echo "🟢 clasp pull (Google → local)"; $CLASP pull ;;
  push) echo "🔵 clasp push (local → Google)"; $CLASP push ;;
  *)    echo "Unknown mode: $MODE (use pull|push)"; exit 1 ;;
esac

git add -A
git diff --cached --quiet || git commit -m "$MSG"
git push origin "$(git rev-parse --abbrev-ref HEAD)"
echo "✅ Done ($MODE → commit → push)."

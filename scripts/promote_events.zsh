#!/bin/zsh
# Append unpromoted events from data/events.jsonl into Registro_Real.

set -euo pipefail

SCRIPT_DIR="${0:A:h}"
PROJECT="${SCRIPT_DIR:h}"
PYTHON="$PROJECT/.venv/bin/python"

if [[ ! -x "$PYTHON" ]]; then
  print -u2 -- "ERROR: Missing venv. Run: $PROJECT/scripts/setup.zsh"
  exit 1
fi

cd "$PROJECT"
exec "$PYTHON" "$PROJECT/src/promote_main.py"

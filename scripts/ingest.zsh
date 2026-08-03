#!/bin/zsh
# Ingest IBKR Activity/Flex CSV files from data/statements into events.jsonl
# and refresh IBKR_Eventos_Staging on the tax workbook.

set -euo pipefail

SCRIPT_DIR="${0:A:h}"
PROJECT="${SCRIPT_DIR:h}"
PYTHON="$PROJECT/.venv/bin/python"

if [[ ! -x "$PYTHON" ]]; then
  print -u2 -- "ERROR: Missing venv. Run: $PROJECT/scripts/setup.zsh"
  exit 1
fi

cd "$PROJECT"
exec "$PYTHON" "$PROJECT/src/ingest_main.py"

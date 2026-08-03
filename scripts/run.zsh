#!/bin/zsh
# Run the IBKR-Excel-Bridge collector once, with pre-flight checks.
#
# Uses ${0:A:h} to derive the project directory so the script is portable
# across clones and does not embed an absolute user path.

set -euo pipefail

SCRIPT_DIR="${0:A:h}"
PROJECT="${SCRIPT_DIR:h}"
PYTHON="$PROJECT/.venv/bin/python"
MAIN="$PROJECT/src/main.py"
CONFIG="$PROJECT/config/settings.json"

if [[ ! -d "$PROJECT" ]]; then
  print -u2 -- "ERROR: Project directory missing: $PROJECT"
  exit 1
fi

if [[ ! -x "$PYTHON" ]]; then
  print -u2 -- "ERROR: Virtual environment is missing at $PROJECT/.venv"
  print -u2 -- "Run: $PROJECT/scripts/setup.zsh"
  exit 1
fi

if ! "$PYTHON" -c "import ibapi, openpyxl" 2>/dev/null; then
  print -u2 -- "ERROR: Required Python packages are not installed."
  print -u2 -- "Run: $PROJECT/scripts/setup.zsh"
  print -u2 -- "Then follow the printed instructions to install ibapi from IBKR."
  exit 1
fi

if [[ ! -f "$CONFIG" ]]; then
  print -u2 -- "ERROR: Missing $CONFIG"
  exit 1
fi

# Derive host and port from settings.json (fallback to TWS paper defaults).
HOST="$("$PYTHON" -c 'import json,sys; print(json.load(open(sys.argv[1]))["ibkr"]["host"])' "$CONFIG" 2>/dev/null || print -- "127.0.0.1")"
PORT="$("$PYTHON" -c 'import json,sys; print(json.load(open(sys.argv[1]))["ibkr"]["port"])' "$CONFIG" 2>/dev/null || print -- "7497")"

if ! /usr/bin/nc -z "$HOST" "$PORT" 2>/dev/null; then
  print -u2 -- "ERROR: IBKR API is not listening on $HOST:$PORT."
  print -u2 -- "Open TWS or IB Gateway and confirm socket clients are enabled."
  print -u2 -- "Default ports: TWS paper 7497 / live 7496; Gateway paper 4002 / live 4001."
  exit 1
fi

cd "$PROJECT"
exec "$PYTHON" "$MAIN"

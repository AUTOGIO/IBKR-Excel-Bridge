#!/bin/zsh
# One-time environment bootstrap for IBKR-Excel-Bridge.
#
# Creates the virtual environment, installs runtime deps, and prints the exact
# command to install the IBKR TWS API from the IBKR-distributed source tree.

set -euo pipefail

SCRIPT_DIR="${0:A:h}"
PROJECT="${SCRIPT_DIR:h}"
VENV="$PROJECT/.venv"
PYTHON_BIN="$VENV/bin/python"

print -- "Project: $PROJECT"

if [[ ! -d "$VENV" ]]; then
  print -- "Creating virtual environment at $VENV"
  /usr/bin/env python3 -m venv "$VENV"
fi

"$PYTHON_BIN" -m pip install --upgrade pip
"$PYTHON_BIN" -m pip install -r "$PROJECT/requirements.txt"

print -- ""
print -- "Runtime dependencies installed."
print -- ""

if "$PYTHON_BIN" -c "import ibapi" 2>/dev/null; then
  print -- "ibapi is already available in the virtual environment."
else
  cat <<'INSTRUCTIONS'
NEXT STEP: install the IBKR TWS API (ibapi) from IBKR's Mac/Unix distribution.

  1. Download "API Latest (Mac / Unix)" from:
       https://interactivebrokers.github.io/
  2. Extract the archive; the Python source lives at:
       <TWS API extract>/source/pythonclient
  3. Install into this project's venv:
       "PROJECT/.venv/bin/python" -m pip install "<TWS API extract>/source/pythonclient"
  4. Verify:
       "PROJECT/.venv/bin/python" -c "import ibapi; print(ibapi.__file__)"

  (Replace PROJECT with this project's path and <TWS API extract> with wherever
  you unpacked the archive.)
INSTRUCTIONS
fi

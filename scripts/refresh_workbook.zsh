#!/bin/zsh
# Close the configured workbook in Excel (if open), run the IBKR snapshot,
# then reopen the refreshed file. Intended for the Overview "Refresh" control
# and for double-clicking refresh_workbook.command.
#
# Uses ${0:A:h} so the script is portable across clones.

set -euo pipefail

SCRIPT_DIR="${0:A:h}"
PROJECT="${SCRIPT_DIR:h}"
PYTHON="$PROJECT/.venv/bin/python"
RUN="$PROJECT/scripts/run.zsh"

if [[ ! -x "$PYTHON" ]]; then
  print -u2 -- "ERROR: Virtual environment is missing at $PROJECT/.venv"
  print -u2 -- "Run: $PROJECT/scripts/setup.zsh"
  exit 1
fi

if [[ ! -x "$RUN" ]]; then
  print -u2 -- "ERROR: Missing executable $RUN"
  exit 1
fi

cd "$PROJECT"

OUTPUT_PATH="$("$PYTHON" - <<'PY'
from pathlib import Path
import sys

sys.path.insert(0, str(Path("src").resolve()))
from config_loader import load_config
from main import resolve_output_path

root = Path(".").resolve()
cfg = load_config(root)
print(resolve_output_path(cfg["excel"], root))
PY
)"

WORKBOOK_NAME="${OUTPUT_PATH:t}"
# Excel lock files look like "~$IBKR_Portfolio.xlsx" next to the workbook.
LOCK_PATH=$(printf '%s/~$%s' "${OUTPUT_PATH:h}" "$WORKBOOK_NAME")

print -- "Refreshing: $OUTPUT_PATH"

# Close only this workbook so other Excel files stay open.
osascript <<EOF >/dev/null 2>&1 || true
tell application "Microsoft Excel"
  if it is running then
    try
      close workbook "$WORKBOOK_NAME" saving yes
    end try
  end if
end tell
EOF

# Drop a leftover Excel lock if Excel quit uncleanly.
rm -f "$LOCK_PATH"

# Brief pause so Excel finishes releasing the file before we overwrite it.
sleep 1

"$RUN"

print -- "Reopening workbook…"
open "$OUTPUT_PATH"
print -- "Done."

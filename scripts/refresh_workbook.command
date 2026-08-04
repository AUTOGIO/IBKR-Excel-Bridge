#!/bin/zsh
# Double-clickable / Excel-hyperlink entry point for refresh_workbook.zsh.
# Opens in Terminal.app when launched from Finder or a spreadsheet link.

set -euo pipefail

SCRIPT_DIR="${0:A:h}"
exec "$SCRIPT_DIR/refresh_workbook.zsh"

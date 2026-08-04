#!/usr/bin/env python3
"""Promote unpromoted events from events.jsonl into Registro_Real."""

from __future__ import annotations

import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))

from config_loader import load_config  # noqa: E402
from events_store import load_events  # noqa: E402
from promote_events import promote_events_to_workbook, write_staging_sheets  # noqa: E402

PROJECT_ROOT = Path(__file__).resolve().parent.parent


def main() -> int:
    config = load_config(PROJECT_ROOT)
    ingest_cfg = config.get("ingest", {})
    excel_cfg = config.get("excel", {})

    events_path = PROJECT_ROOT / ingest_cfg.get("events_file", "data/events.jsonl")
    tax_workbook = PROJECT_ROOT / excel_cfg.get(
        "tax_workbook", "data/output/TRIBUTACAO_WORKING.xlsx"
    )

    summary = promote_events_to_workbook(tax_workbook, events_path)
    events = load_events(events_path)
    write_staging_sheets(tax_workbook, events)

    print(
        f"\nSUCCESS: appended={summary['appended']} "
        f"skipped_existing={summary['skipped']} "
        f"workbook={tax_workbook}\n"
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # noqa: BLE001
        print(f"\nERROR: {exc}\n", file=sys.stderr)
        raise SystemExit(1) from exc

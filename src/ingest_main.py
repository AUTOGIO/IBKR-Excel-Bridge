#!/usr/bin/env python3
"""Ingest IBKR statement CSVs → events.jsonl + staging sheets."""

from __future__ import annotations

import json
import logging
import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))

from events_store import load_events, merge_events, save_events  # noqa: E402
from flex_client import download_flex_statement  # noqa: E402
from ingest_statements import ingest_directory, load_aliases  # noqa: E402
from promote_events import write_staging_sheets  # noqa: E402

PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_FILE = PROJECT_ROOT / "config" / "settings.json"


def main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)-7s | %(message)s",
    )
    log = logging.getLogger("ingest")

    config = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    ingest_cfg = config.get("ingest", {})
    excel_cfg = config.get("excel", {})
    flex_cfg = config.get("flex", {})

    statements_dir = PROJECT_ROOT / ingest_cfg.get("statements_dir", "data/statements")
    events_path = PROJECT_ROOT / ingest_cfg.get("events_file", "data/events.jsonl")
    aliases_path = PROJECT_ROOT / ingest_cfg.get(
        "aliases_file", "config/symbol_aliases.json"
    )
    tax_workbook = PROJECT_ROOT / excel_cfg.get(
        "tax_workbook", "output/U6658119_TRIBUTACAO_WORKING.xlsx"
    )

    statements_dir.mkdir(parents=True, exist_ok=True)

    if flex_cfg.get("enabled"):
        try:
            download_flex_statement(
                token=str(flex_cfg.get("token", "")),
                query_id=str(flex_cfg.get("query_id", "")),
                out_dir=statements_dir,
                base_url=str(
                    flex_cfg.get(
                        "base_url",
                        "https://ndcdyn.interactivebrokers.com/AccountManagement/FlexWebService",
                    )
                ),
            )
        except NotImplementedError as exc:
            log.error("%s", exc)
            return 2

    aliases = load_aliases(aliases_path)
    incoming = ingest_directory(statements_dir, aliases=aliases)
    existing = load_events(events_path)
    merged, new_count = merge_events(existing, incoming)
    save_events(events_path, merged)

    staging_ok = False
    if tax_workbook.exists():
        try:
            write_staging_sheets(tax_workbook, merged)
            staging_ok = True
            log.info("Staging sheets updated in %s", tax_workbook)
        except PermissionError as exc:
            log.warning("%s", exc)
            log.warning("events.jsonl saved; close Excel and re-run ingest to refresh staging.")
    else:
        log.warning("Tax workbook missing (%s); events saved only.", tax_workbook)

    print(
        f"\nSUCCESS: ingested {len(incoming)} row(s) from CSV; "
        f"{new_count} new event(s); store={events_path} "
        f"total={len(merged)}; staging={'yes' if staging_ok else 'deferred'}\n"
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # noqa: BLE001
        print(f"\nERROR: {exc}\n", file=sys.stderr)
        raise SystemExit(1) from exc

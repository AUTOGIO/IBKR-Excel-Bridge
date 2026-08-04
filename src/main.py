#!/usr/bin/env python3
"""Entry point: connect to TWS/Gateway, collect a read-only snapshot, write Excel."""

from __future__ import annotations

import logging
import logging.handlers
import socket
import sys
from pathlib import Path
from typing import Any

# Support being run as ``python src/main.py`` and as ``python -m src.main``.
_HERE = Path(__file__).resolve().parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))

from config_loader import load_config as _load_settings  # noqa: E402
from excel_exporter import ExcelExporter  # noqa: E402
from ibkr_client import IBKRClient  # noqa: E402


PROJECT_ROOT = Path(__file__).resolve().parent.parent
LOG_FILE = PROJECT_ROOT / "logs" / "ibkr_excel_bridge.log"

# Cap per-file log size so a chatty ibapi run cannot fill the disk over time.
_LOG_MAX_BYTES = 5_000_000
_LOG_BACKUP_COUNT = 5


def configure_logging(level_name: str = "INFO") -> None:
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    level = getattr(logging, level_name.upper(), logging.INFO)
    # ``force=True`` lets a second call (e.g. from tests) replace handlers
    # instead of silently no-oping when the root logger is already set up.
    logging.basicConfig(
        level=level,
        format="%(asctime)s | %(levelname)-7s | %(name)s | %(message)s",
        handlers=[
            logging.handlers.RotatingFileHandler(
                LOG_FILE,
                maxBytes=_LOG_MAX_BYTES,
                backupCount=_LOG_BACKUP_COUNT,
                encoding="utf-8",
            ),
            logging.StreamHandler(sys.stdout),
        ],
        force=True,
    )
    # ibapi INFO logs every REQUEST/ANSWER protobuf frame — several MB per
    # run. Real errors still surface at WARNING/ERROR.
    logging.getLogger("ibapi").setLevel(logging.WARNING)


def load_config() -> dict:
    return _load_settings(PROJECT_ROOT)


def check_port_open(host: str, port: int, timeout: float = 2.0) -> bool:
    """Quick TCP probe so we fail fast with a friendly message if TWS is closed."""
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def resolve_output_path(excel_config: dict[str, Any], project_root: Path) -> Path:
    mode = excel_config.get("output_mode", "standalone")
    if mode == "tax_workbook":
        rel = excel_config["tax_workbook"]
    elif mode == "standalone":
        rel = excel_config["output_file"]
    else:
        raise ValueError(f"Unknown excel.output_mode: {mode!r}")
    return project_root / rel


def assert_expected_account(data: dict[str, Any], expected: str) -> None:
    expected = (expected or "").strip()
    if not expected:
        return
    accounts = {row.get("account") for row in data.get("accounts", [])}
    if expected not in accounts:
        found = sorted(a for a in accounts if a)
        raise ValueError(
            f"expected_account {expected!r} not in managed accounts {found}"
        )


def main() -> int:
    config = load_config()
    configure_logging(config.get("logging", {}).get("level", "INFO"))
    logger = logging.getLogger("ibkr_excel_bridge")

    client: IBKRClient | None = None

    try:
        ibkr_config = config["ibkr"]
        excel_config = config["excel"]

        if not ibkr_config.get("require_read_only_confirmation", True):
            logger.warning(
                "'require_read_only_confirmation' is disabled in settings.json. "
                "This flag is a client-side self-check only; enforce Read-Only "
                "API in TWS/Gateway Global Configuration -> API -> Settings."
            )

        host = ibkr_config["host"]
        port = int(ibkr_config["port"])
        if not check_port_open(host, port):
            raise ConnectionError(
                f"IBKR API socket at {host}:{port} is not reachable. "
                "Open TWS or IB Gateway, verify socket clients are enabled, "
                "and confirm the port matches (TWS paper 7497 / live 7496; "
                "Gateway paper 4002 / live 4001)."
            )

        mode = excel_config.get("output_mode", "standalone")
        output_path = resolve_output_path(excel_config, PROJECT_ROOT)
        qty_tolerance = float(excel_config.get("qty_tolerance", 0.0001))

        client = IBKRClient()
        client.connect_and_start(
            host=host,
            port=port,
            client_id=int(ibkr_config["client_id"]),
            timeout=int(ibkr_config["connection_timeout_seconds"]),
        )

        data = client.collect(
            timeout=int(ibkr_config["collection_timeout_seconds"])
        )

        assert_expected_account(
            data, str(ibkr_config.get("expected_account", "") or "")
        )

        exporter = ExcelExporter(
            output_path,
            mode=mode,
            qty_tolerance=qty_tolerance,
            events_file=PROJECT_ROOT
            / config.get("ingest", {}).get("events_file", "data/events.jsonl"),
        )
        created_file = exporter.export(data)

        logger.info("Excel workbook created: %s", created_file)
        print(f"\nSUCCESS: {created_file}\n")
        return 0

    except KeyboardInterrupt:
        logger.warning("Execution interrupted by user.")
        return 130
    except Exception as error:  # noqa: BLE001 - top-level guard
        logger.exception("Execution failed: %s", error)
        print(f"\nERROR: {error}\n", file=sys.stderr)
        return 1
    finally:
        if client is not None:
            client.stop()


if __name__ == "__main__":
    raise SystemExit(main())

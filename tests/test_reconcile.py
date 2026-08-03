"""Unit tests for quantity reconciliation helpers."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from reconcile import normalize_symbol, reconcile_quantities  # noqa: E402


def test_normalize_symbol_trims_and_uppercases() -> None:
    assert normalize_symbol("  amzn ") == "AMZN"


def test_reconcile_statuses_and_tolerance() -> None:
    live = [
        {"symbol": "AMZN", "quantity": 7.0, "instrument_kind": "Stock"},
        {"symbol": "EUR", "quantity": 1000.0, "instrument_kind": "FX Cash"},
        {"symbol": "NEW", "quantity": 1.0, "instrument_kind": "Stock"},
    ]
    fiscal = [
        {"symbol": "amzn", "quantity": 7.0},
        {"symbol": "BIL", "quantity": 100.0},
    ]
    rows = {r["symbol"]: r for r in reconcile_quantities(live=live, fiscal=fiscal)}
    assert "EUR" not in rows
    assert rows["AMZN"]["status"] == "OK"
    assert rows["NEW"]["status"] == "ONLY_IBKR"
    assert rows["BIL"]["status"] == "ONLY_FISCAL"


def test_diverge_above_tolerance() -> None:
    live = [{"symbol": "BIL", "quantity": 100.0, "instrument_kind": "Stock"}]
    fiscal = [{"symbol": "BIL", "quantity": 99.0}]
    rows = reconcile_quantities(live=live, fiscal=fiscal, tolerance=0.0001)
    assert rows[0]["status"] == "DIVERGE"
    assert rows[0]["delta"] == 1.0


def test_within_tolerance_is_ok() -> None:
    live = [{"symbol": "X", "quantity": 10.00005, "instrument_kind": "Stock"}]
    fiscal = [{"symbol": "X", "quantity": 10.0}]
    rows = reconcile_quantities(live=live, fiscal=fiscal, tolerance=0.0001)
    assert rows[0]["status"] == "OK"

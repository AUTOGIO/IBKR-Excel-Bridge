"""Quantity reconciliation between live IBKR positions and fiscal sheets."""

from __future__ import annotations

from typing import Any


def normalize_symbol(value: Any) -> str:
    return str(value or "").strip().upper()


def reconcile_quantities(
    *,
    live: list[dict[str, Any]],
    fiscal: list[dict[str, Any]],
    tolerance: float = 0.0001,
) -> list[dict[str, Any]]:
    """Compare live vs fiscal quantities by normalized symbol.

    Skips live rows whose ``instrument_kind`` is ``FX Cash``. Status values:
    ``OK``, ``DIVERGE``, ``ONLY_IBKR``, ``ONLY_FISCAL``.
    """
    live_map: dict[str, float] = {}
    for row in live:
        if (row.get("instrument_kind") or "") == "FX Cash":
            continue
        symbol = normalize_symbol(row.get("symbol"))
        if not symbol:
            continue
        qty = float(row.get("quantity") or 0)
        live_map[symbol] = live_map.get(symbol, 0.0) + qty

    fiscal_map: dict[str, float] = {}
    for row in fiscal:
        symbol = normalize_symbol(row.get("symbol"))
        if not symbol:
            continue
        raw = row.get("quantity")
        if raw is None or raw == "":
            continue
        try:
            qty = float(raw)
        except (TypeError, ValueError):
            continue
        fiscal_map[symbol] = fiscal_map.get(symbol, 0.0) + qty

    out: list[dict[str, Any]] = []
    for symbol in sorted(set(live_map) | set(fiscal_map)):
        q_live = live_map.get(symbol)
        q_fiscal = fiscal_map.get(symbol)
        if q_live is None:
            status = "ONLY_FISCAL"
            delta = None if q_fiscal is None else -float(q_fiscal)
        elif q_fiscal is None:
            status = "ONLY_IBKR"
            delta = float(q_live)
        else:
            delta = float(q_live) - float(q_fiscal)
            status = "OK" if abs(delta) <= tolerance else "DIVERGE"
        out.append(
            {
                "symbol": symbol,
                "qty_ibkr": q_live,
                "qty_fiscal": q_fiscal,
                "delta": delta,
                "status": status,
            }
        )
    return out


__all__ = ["normalize_symbol", "reconcile_quantities"]

"""Append-only JSONL event store for IBKR statement ingest."""

from __future__ import annotations

import hashlib
import json
from collections.abc import Iterable
from pathlib import Path
from typing import Any


def make_event_id(
    *,
    date: str,
    symbol: str,
    tipo_evento: str,
    quantity: float,
    price_usd: float | None,
    source_key: str = "",
) -> str:
    raw = "|".join(
        [
            date,
            symbol.upper().strip(),
            tipo_evento,
            f"{float(quantity):.8f}",
            "" if price_usd is None else f"{float(price_usd):.8f}",
            source_key,
        ]
    )
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def load_events(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    events: list[dict[str, Any]] = []
    with path.open("r", encoding="utf-8") as handle:
        for line in handle:
            line = line.strip()
            if not line:
                continue
            events.append(json.loads(line))
    return events


def save_events(path: Path, events: Iterable[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    ordered = sorted(
        events,
        key=lambda e: (
            str(e.get("date", "")),
            str(e.get("symbol", "")),
            str(e.get("event_id", "")),
        ),
    )
    with path.open("w", encoding="utf-8") as handle:
        for event in ordered:
            handle.write(json.dumps(event, ensure_ascii=False, sort_keys=True) + "\n")


def merge_events(
    existing: list[dict[str, Any]],
    incoming: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], int]:
    """Merge by event_id. Preserve promoted=True from existing. Returns (merged, new_count)."""
    by_id: dict[str, dict[str, Any]] = {
        str(e["event_id"]): dict(e) for e in existing if e.get("event_id")
    }
    new_count = 0
    for event in incoming:
        eid = str(event["event_id"])
        if eid not in by_id:
            by_id[eid] = dict(event)
            new_count += 1
            continue
        prior = by_id[eid]
        merged = dict(event)
        if prior.get("promoted"):
            merged["promoted"] = True
        by_id[eid] = merged
    return list(by_id.values()), new_count


__all__ = ["make_event_id", "load_events", "save_events", "merge_events"]

"""Parse IBKR Activity / Flex CSV exports into canonical events."""

from __future__ import annotations

import csv
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from events_store import make_event_id


_DATE_FORMATS = (
    "%Y-%m-%d",
    "%Y%m%d",
    "%d/%m/%Y",
    "%m/%d/%Y",
    "%Y-%m-%d %H:%M:%S",
    "%Y%m%d%H%M%S",
    "%Y-%m-%d, %H:%M:%S",
)


def load_aliases(path: Path | None) -> dict[str, str]:
    if path is None or not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    return {str(k).upper(): str(v).upper() for k, v in data.items()}


def normalize_symbol(symbol: str, aliases: dict[str, str]) -> str:
    sym = str(symbol or "").strip().upper()
    return aliases.get(sym, sym)


def _parse_date(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip().strip('"')
    if not text:
        return None
    # Flex sometimes uses 20260715 or 20260715;120000
    text = text.split(";")[0].split(",")[0].strip()
    if " " in text and re.match(r"^\d{4}-\d{2}-\d{2}", text):
        text = text.split(" ")[0]
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    # last resort: first 8 digits YYYYMMDD
    digits = re.sub(r"\D", "", text)
    if len(digits) >= 8:
        try:
            return datetime.strptime(digits[:8], "%Y%m%d").date().isoformat()
        except ValueError:
            return None
    return None


def _to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    text = str(value).strip().replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def _norm_header(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(name).strip().lower())


_HEADER_ALIASES: dict[str, tuple[str, ...]] = {
    "symbol": ("symbol", "symbollocal", "underlyingSymbol".lower()),
    "datetime": (
        "datetime",
        "datetime",
        "tradedatetime",
        "dateitime",
        "reportdate",
        "settledate",
        "date",
        "tradedate",
        "exdate",
        "paydate",
    ),
    "quantity": ("quantity", "qty", "shares"),
    "price": ("tradeprice", "price", "unitprice", "proceeds"),
    "buysell": ("buysell", "side", "buy/sell", "buySell".lower()),
    "tradetype": ("tradetype", "transactiontype", "type", "datatype"),
    "commission": ("ibcommission", "commission"),
    "tax": ("tax", "withholdingtax", "wht"),
    "tradeid": ("tradeid", "transactionid", "orderid"),
    "currency": ("currencyprimary", "currency"),
}


def _resolve_columns(headers: list[str]) -> dict[str, int]:
    normalized = [_norm_header(h) for h in headers]
    resolved: dict[str, int] = {}
    for logical, aliases in _HEADER_ALIASES.items():
        for alias in aliases:
            key = _norm_header(alias)
            if key in normalized:
                resolved[logical] = normalized.index(key)
                break
    return resolved


def _looks_like_header(row: list[str]) -> bool:
    joined = " ".join(_norm_header(c) for c in row)
    return "symbol" in joined and (
        "quantity" in joined or "qty" in joined or "tradeprice" in joined or "price" in joined
    )


def _map_tipo(buysell: str | None, tradetype: str | None, quantity: float) -> str | None:
    bs = (buysell or "").strip().upper()
    tt = (tradetype or "").strip().upper()
    if any(x in tt for x in ("DIV", "DIVIDEND", "PIK", "PAYMENT IN LIEU")):
        return "Rendimento"
    if bs in {"BUY", "BOT", "B"}:
        return "Compra"
    if bs in {"SELL", "SLD", "S"}:
        return "Venda"
    if "BUY" in tt:
        return "Compra"
    if "SELL" in tt:
        return "Venda"
    if quantity > 0 and not bs:
        return "Compra"
    if quantity < 0 and not bs:
        return "Venda"
    return None


def parse_statement_csv(
    path: Path,
    *,
    aliases: dict[str, str] | None = None,
) -> list[dict[str, Any]]:
    aliases = aliases or {}
    text = path.read_text(encoding="utf-8-sig", errors="replace")
    # IBKR Flex often has section lines before the real header
    lines = text.splitlines()
    reader_rows: list[list[str]] = []
    for line in lines:
        if not line.strip():
            continue
        reader_rows.append(next(csv.reader([line])))

    header_idx = None
    for idx, row in enumerate(reader_rows):
        if _looks_like_header(row):
            header_idx = idx
            break
    if header_idx is None:
        raise ValueError(f"No IBKR header row found in {path.name}")

    headers = reader_rows[header_idx]
    cols = _resolve_columns(headers)
    if "symbol" not in cols or "quantity" not in cols:
        raise ValueError(
            f"CSV {path.name} missing Symbol/Quantity columns (got {headers})"
        )

    events: list[dict[str, Any]] = []
    for row_num, row in enumerate(reader_rows[header_idx + 1 :], start=header_idx + 2):
        if not row or all(not str(c).strip() for c in row):
            continue
        # skip section banners
        if len(row) == 1 or _looks_like_header(row):
            continue

        def cell(logical: str) -> str:
            idx = cols.get(logical)
            if idx is None or idx >= len(row):
                return ""
            return str(row[idx]).strip()

        symbol_raw = cell("symbol")
        if not symbol_raw or symbol_raw.upper() in {"TOTAL", "SYMBOL"}:
            continue
        symbol = normalize_symbol(symbol_raw, aliases)
        date = _parse_date(cell("datetime"))
        if not date:
            continue
        qty = _to_float(cell("quantity"))
        if qty is None or qty == 0:
            continue
        price = _to_float(cell("price"))
        tax = _to_float(cell("tax"))
        tipo = _map_tipo(cell("buysell"), cell("tradetype"), qty)
        if tipo is None:
            continue
        quantity = abs(qty)
        # For Rendimento, price column may be dividend amount total — keep as unit if qty==1
        trade_id = cell("tradeid")
        source_key = trade_id or f"{path.name}:{row_num}"
        event_id = make_event_id(
            date=date,
            symbol=symbol,
            tipo_evento=tipo,
            quantity=quantity,
            price_usd=price,
            source_key=source_key,
        )
        if tipo == "Rendimento":
            obs = (
                f"IBKR dividendo {date} | event_id={event_id} | "
                f"source={path.name} | ano={date[:4]}"
            )
        else:
            side = "O" if tipo == "Compra" else "C"
            obs = (
                f"IBKR trade {date} | code={side} | event_id={event_id} | "
                f"source={path.name} | ano={date[:4]}"
            )
        events.append(
            {
                "event_id": event_id,
                "date": date,
                "symbol": symbol,
                "tipo_evento": tipo,
                "quantity": quantity,
                "price_usd": price,
                "ir_retido_usd": tax if tipo == "Rendimento" else None,
                "ptax": None,
                "observacoes": obs,
                "source_file": path.name,
                "promoted": False,
            }
        )
    return events


def ingest_directory(
    statements_dir: Path,
    *,
    aliases: dict[str, str] | None = None,
) -> list[dict[str, Any]]:
    if not statements_dir.exists():
        raise FileNotFoundError(f"Statements directory not found: {statements_dir}")
    files = sorted(statements_dir.glob("*.csv"))
    if not files:
        return []
    all_events: list[dict[str, Any]] = []
    errors: list[str] = []
    for path in files:
        try:
            all_events.extend(parse_statement_csv(path, aliases=aliases))
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{path.name}: {exc}")
    if errors and not all_events:
        raise ValueError("No events parsed:\n" + "\n".join(errors))
    if errors:
        # partial success — surface via observacoes on a synthetic log event? Prefer raise soft.
        pass
    return all_events


__all__ = [
    "load_aliases",
    "normalize_symbol",
    "parse_statement_csv",
    "ingest_directory",
]

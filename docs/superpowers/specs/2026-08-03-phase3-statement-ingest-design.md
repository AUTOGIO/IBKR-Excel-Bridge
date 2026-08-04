# Phase 3 — Activity Statement Ingest Design

**Date:** 2026-08-03  
**Status:** Approved  
**Depends on:** Phase 1 tax-workbook mode (`docs/superpowers/specs/2026-08-03-tax-workbook-api-integration-design.md`)

## Goal

Automate the path from IBKR Activity / Flex exports into `Registro_Real`, with a review gate. CSV drop ships first; Flex Web Service is stubbed for the next slice.

## Non-goals (this slice)

- Auto-overwrite `MyProfit_2026` (write derived `IBKR_Posicao_From_Events` only).
- Live BCB PTAX fetch (PTAX left blank; Excel validation flags missing FX).
- Futures/options tax mapping beyond passing symbol/qty/price through.
- Placing orders or disabling Read-Only API.

## Pipeline

```text
data/statements/*.csv  (+ optional Flex download into same folder)
        │
        ▼
  ingest_statements.parse + merge
        ▼
  data/events.jsonl          canonical event store (deduped by event_id)
        ▼
  IBKR_Eventos_Staging       review sheet in tax workbook
        ▼
  promote_events             append NEW rows only → Registro_Real
        │                    (copy formulas H–O from template)
        ▼
  IBKR_Posicao_From_Events   derived qty / cost from ledger (not MyProfit)
```

## Canonical event schema (`data/events.jsonl`)

One JSON object per line:

| Field | Type | Notes |
| --- | --- | --- |
| `event_id` | string | Stable hash: first 16 hex chars of `sha1(date\|symbol\|tipo\|qty\|price\|source_row)` (64 bits — collision-free for personal statement volumes); or the IBKR trade id when present in the CSV |
| `date` | string | ISO date `YYYY-MM-DD` (trade/report date from file; settlement adjustment later) |
| `symbol` | string | Normalized upper-case; alias map applied |
| `tipo_evento` | string | `Compra` \| `Venda` \| `Rendimento` |
| `quantity` | number | Always positive; side encoded in `tipo_evento` |
| `price_usd` | number \| null | Unit price or dividend per share gross |
| `ir_retido_usd` | number \| null | Withholding for dividends |
| `ptax` | number \| null | Blank in this slice |
| `observacoes` | string | Human + machine tag including `event_id=` |
| `source_file` | string | Basename of CSV |
| `promoted` | bool | Set true after successful append to `Registro_Real` |

## CSV parsing

- Input dir: `data/statements/` (configurable).
- Skip non-`.csv`.
- Detect header row by known column-name tokens (`Symbol`, `Date/Time`, `Quantity`, `TradePrice`, `Buy/Sell`, `Proceeds`, or Flex dividend headers).
- Map Buy → `Compra`, Sell → `Venda`.
- Dividends / Payment In Lieu → `Rendimento` when detectable.
- Alias file: `config/symbol_aliases.json` (e.g. `{"LQDEz":"LQDE"}`).
- Unknown layouts: fail the file with a clear error; do not partially promote.

## Staging sheet `IBKR_Eventos_Staging`

Machine-owned (deleted/rewritten like other `IBKR_*` sheets). Columns mirror canonical fields + `promoted` status. Written by ingest into the tax workbook path from settings.

## Promote rules

- Only events with `promoted=false`.
- Append after last non-empty `Registro_Real` data row (col A).
- Write values: A–G, L (`observacoes` must include `event_id=...`).
- Copy formulas from the previous data row for columns H–O, adjusting row index.
- Mark events `promoted=true` in `events.jsonl` after save.
- Never delete or rewrite existing `Registro_Real` rows.

## `IBKR_Posicao_From_Events`

Derived from **promoted + staging-eligible** events in `events.jsonl` plus existing `Registro_Real` rows already in the workbook (read values for A–G). Roll forward per symbol: Compra adds qty/cost, Venda reduces qty (cost basis simplified USD average; BRL average only when PTAX present). Output for human copy into MyProfit — not an automatic overwrite.

## Flex API (stub this slice)

- Module `src/flex_client.py` with `download_flex_statement(token, query_id, out_dir) -> Path`.
- Raises `NotImplementedError` with setup instructions unless `flex.enabled` is true **and** credentials present — for this slice, keep `flex.enabled: false` and document the interface.
- When enabled later: download into `data/statements/` then call the same ingest.

## Config additions (`config/settings.json`)

```json
"ingest": {
  "statements_dir": "data/statements",
  "events_file": "data/events.jsonl",
  "aliases_file": "config/symbol_aliases.json"
},
"flex": {
  "enabled": false,
  "token": "",
  "query_id": "",
  "base_url": "https://ndcdyn.interactivebrokers.com/AccountManagement/FlexWebService"
}
```

## CLI

| Script | Action |
| --- | --- |
| `./scripts/ingest.zsh` | Parse CSVs → merge `events.jsonl` → refresh staging (+ posicao sheet) on tax workbook |
| `./scripts/promote_events.zsh` | Append new events to `Registro_Real`, mark promoted |

Tax workbook path = `excel.tax_workbook` from settings. Workbook must exist and not be Excel-locked.

## Testing

- Fixture CSVs for trades + dividends.
- Unit: parse → event_id stability, buy/sell mapping, alias, dedupe.
- Unit: promote appends formulas with correct row numbers; second promote is no-op.
- Do not require live TWS for ingest/promote tests.

## Success criteria

1. Dropping a valid trade CSV and running ingest creates/updates `events.jsonl` and `IBKR_Eventos_Staging`.
2. Promote appends only new rows; re-run is idempotent.
3. `MyProfit_2026` unchanged by automation.
4. Flex module exists but is disabled by default.
5. OPERATING_GUIDE documents the CSV → review → promote → snapshot loop.

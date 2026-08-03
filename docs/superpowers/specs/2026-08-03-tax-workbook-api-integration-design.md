# Tax Workbook ↔ IBKR API Integration Design

**Date:** 2026-08-03  
**Status:** Draft for review  
**Scope:** Phase 1 (live TWS/Gateway snapshot into Lei 14.754 workbook) first; Phase 3 (Flex / statement ingest into `Registro_Real`) designed at high level only.

## Problem

The Lei 14.754 workbook

`output/U6658119_TRIBUTACAO-LEI14754_v5-1-RECONCILIADO_2021-2026 copy.xlsx`

is a curated fiscal model (events, MyProfit costs, annual apuração). This repo’s MVP already snapshots IBKR via the read-only TWS API into `output/IBKR_Portfolio.xlsx`, but that snapshot is isolated from the tax file.

We need a safe way to:

1. **Phase 1** — Refresh live account/position data next to the tax model without corrupting fiscal inputs.
2. **Phase 3** — Later ingest historical trades/dividends into `Registro_Real` without treating Excel as the sole source of truth.

## Non-goals (this design)

- Placing orders or disabling Read-Only API.
- Silently overwriting `Posicoes_Atuais` / `MyProfit_2026` quantities or costs from the live snapshot.
- Replacing Excel formulas for apuração with a Python tax motor (see workbook `Spec_Motor_Python`; out of Phase 1/3).
- Auto-fetching PTAX from BCB (Phase 3 may add it; Phase 1 does not).
- Continuous streaming or scheduled daemon.

## Decisions

| Decision | Choice | Rationale |
| --- | --- | --- |
| Where live data lands | Machine-owned `IBKR_*` sheets **inside a working copy** of the tax workbook | One file to open; matches existing preserve-foreign-sheets exporter; no fragile cross-workbook links |
| Fiscal sheets | Never rewritten by the exporter | Tax integrity; MyProfit remains fiscal cost source of truth for now |
| Quantity trust | Explicit `IBKR_Reconciliacao` diff only | Avoids clean-looking wrong DIRPF |
| Runtime host | Support TWS **and** IB Gateway | Gateway uses ~40% fewer resources for API-only runs |
| Account check | Optional `expected_account` in config | Fail loud if paper `DUR…` is used against a U6658119 workbook by mistake |
| Phase 3 truth | Canonical `events.jsonl` then staged Excel write | Aligns with workbook `Spec_Motor_Python` `ingest.py` guidance |

## Architecture — Phase 1

```text
TWS or IB Gateway (Read-Only)
        │
        ▼
  ibkr_client.collect()
        │
        ▼
  excel_exporter.export()
        │
        ├─ rewrite only OWNED sheets (IBKR_* prefix)
        ├─ write IBKR_Reconciliacao (reads fiscal qty sheets)
        └─ leave all other sheets untouched
        │
        ▼
  tax working workbook under output/
```

### Workbook target

- Long-term path: `output/U6658119_TRIBUTACAO_WORKING.xlsx`
- Bootstrap: copy from the current reconciled file once (manual or setup helper). Do not use the dated `… copy.xlsx` name as the permanent target.
- Optional: keep writing a standalone `output/IBKR_Portfolio.xlsx` when `excel.output_mode` is `standalone` (paper smoke tests without touching the tax file).

### Owned sheets (exporter-owned; deleted and rewritten each run)

Rename from the current unprefixed names to avoid collisions with fiscal tabs:

| Sheet | Role |
| --- | --- |
| `IBKR_Overview` | Generation time, account, row counts, message summary |
| `IBKR_Account_Summary` | Account summary tags |
| `IBKR_Cash_By_Currency` | `$LEDGER-*` per-currency metrics |
| `IBKR_Positions` | Positions including FX cash |
| `IBKR_Portfolio` | Mark-to-market portfolio (no CASH secType) |
| `IBKR_API_Messages` | API info/warning/error rows |
| `IBKR_Reconciliacao` | Symbol-level qty comparison vs fiscal sheets |

On each tax-workbook run, also delete legacy unprefixed owned names if present (`Overview`, `Account Summary`, `Cash By Currency`, `Positions`, `Portfolio`, `API Messages`) so old MVP tabs cannot linger beside the new `IBKR_*` set.

All existing Lei 14.754 sheets (`Posicoes_Atuais`, `MyProfit_2026`, `Registro_Real`, `Apuracao_Anual`, …) are **foreign** and preserved in place.

### Reconciliation rules

`IBKR_Reconciliacao` joins on **normalized symbol** (trim, upper-case; strip known IBKR FX cash rows from the fiscal comparison unless explicitly present).

Sources:

- Live: all `IBKR_Positions` rows except `Kind == FX Cash`.
- Fiscal: `MyProfit_2026` columns `Ativo normalizado` + `Quantidade`. If that sheet is absent, use `Posicoes_Atuais` columns `Ativo` + `Quantidade_Atual` (cached values, not unresolved formula blanks).

Per-row status:

| Status | Meaning |
| --- | --- |
| `OK` | Absolute qty delta ≤ configured tolerance (default `0.0001`) |
| `DIVERGE` | Both sides present; delta above tolerance |
| `ONLY_IBKR` | Live position with no fiscal row |
| `ONLY_FISCAL` | Fiscal row with no live position |

Header metadata: snapshot timestamp, connected account id, expected account (if set), tolerance.

Phase 1 does **not** change formulas in `Posicoes_Atuais`.

### Config (`config/settings.json`)

```json
{
  "ibkr": {
    "host": "127.0.0.1",
    "port": 7497,
    "client_id": 21,
    "connection_timeout_seconds": 15,
    "collection_timeout_seconds": 30,
    "require_read_only_confirmation": true,
    "download_all_accounts": true,
    "expected_account": ""
  },
  "excel": {
    "output_mode": "standalone",
    "output_file": "output/IBKR_Portfolio.xlsx",
    "tax_workbook": "output/U6658119_TRIBUTACAO_WORKING.xlsx",
    "qty_tolerance": 0.0001
  },
  "logging": {
    "level": "INFO"
  }
}
```

Notes:

- **TWS paper** default port `7497`; **IB Gateway paper** default `4002` (live Gateway `4001`, live TWS `7496`). Operator sets `port` to match the running app.
- Default `output_mode` stays `standalone` so existing paper smoke tests keep working unchanged.
- When `output_mode` is `tax_workbook`, exporter writes to `tax_workbook` and requires the file to exist (refuse to create a blank tax model). Reconciliation sheet is written only in this mode.
- When `output_mode` is `standalone`, behavior matches today’s MVP (`output_file` only; no reconciliation sheet).
- If `expected_account` is non-empty and the collected account set does not contain it, exit non-zero after collect and **before** mutating owned sheets.

### Code touchpoints

| File | Change |
| --- | --- |
| `config/settings.json` | New excel/ibkr keys above |
| `src/excel_exporter.py` | `IBKR_` owned names; tax-workbook path; reconciliation writer; refuse missing tax file |
| `src/main.py` | Resolve output path from `output_mode`; enforce `expected_account` |
| `src/ibkr_client.py` | No behavioral change required for Phase 1 (already returns accounts/positions) |
| `README.md` | Document tax mode, Gateway ports, working-copy bootstrap |
| `tests/` | Unit tests for symbol normalize, recon statuses, owned-sheet rename, missing tax file error |

### Operations checklist (Phase 1)

1. Copy reconciled tax xlsx → `output/U6658119_TRIBUTACAO_WORKING.xlsx`.
2. Prefer **IB Gateway** for API-only runs (lower resource use); enable socket clients + Read-Only API; set port in config.
3. For real recon against the tax account, log Gateway/TWS into **U6658119** and set `expected_account` to that id.
4. Run `./scripts/run.zsh` with `output_mode: tax_workbook`.
5. Open `IBKR_Reconciliacao`; resolve `DIVERGE` / `ONLY_*` manually before trusting any downstream use.
6. Keep paper/`standalone` available for regression without touching the tax file.

### Error handling

- Tax workbook path missing → clear error; do not create empty workbook.
- Tax workbook locked by Excel (`~$…`) → fail with “close Excel and retry”.
- Account mismatch → fail before rewrite.
- Corrupt xlsx on load → do **not** fall back to a blank workbook when `output_mode` is `tax_workbook` (standalone mode may keep today’s fallback).

## Architecture — Phase 3 (outline only)

Phase 3 starts only after Phase 1 recon is trustworthy on the target account.

```text
IBKR Flex Query / Activity Statement CSV
        │
        ▼
  ingest (new module)
        │  alias map, settlement date T+1/T+2, event typing
        ▼
  data/events.jsonl   ← canonical, versionable
        │
        ├─ optional ptax cache (append-only) later
        ▼
  staged sheet IBKR_Eventos_Staging
        │  human review
        ▼
  append/merge into Registro_Real (never silent full replace)
```

Constraints carried from `Spec_Motor_Python`:

- Alias resolution before average-cost logic (R14).
- Sale without acquisition cost must block apuração (R07) — ingest must surface incomplete cost basis, not drop rows.
- PTAX by **settlement** date, not order date (R12); cache append-only if/when added.
- Excel remains a presentation / review surface; `events.jsonl` is the durable ingest product.

Phase 3 implementation plan is a separate document after Phase 1 ships.

## Testing strategy — Phase 1

- **Unit:** reconciliation matrix (OK / DIVERGE / ONLY_IBKR / ONLY_FISCAL); tolerance boundary; symbol normalization.
- **Unit:** exporter with a fixture workbook containing foreign sheets + fake fiscal qty table; assert foreign sheets survive and owned sheets refresh.
- **Unit:** `tax_workbook` mode refuses missing/corrupt file.
- **Smoke (manual):** Gateway paper → `standalone`; then tax mode against working copy with `expected_account` unset; then live U6658119 with expected account set.

## Success criteria — Phase 1

1. One command refreshes only `IBKR_*` sheets inside the working tax workbook.
2. All pre-existing Lei 14.754 sheet **cell values and formulas** are preserved; foreign sheet order stays ahead of newly appended `IBKR_*` tabs.
3. `IBKR_Reconciliacao` correctly flags known mismatches in a fixture.
4. Wrong-account runs fail when `expected_account` is set.
5. README documents TWS vs Gateway ports and the bootstrap copy step.
6. Read-only API posture unchanged.

## Open follow-ups (explicitly deferred)

- Rewiring `Posicoes_Atuais` to prefer IBKR qty after recon is clean.
- Flex Query token/storage layout for Phase 3.
- BCB PTAX client and holiday calendars.
- Python apuração motor from `Spec_Motor_Python`.

# IBKR → Excel Bridge

Read-only macOS pipeline that connects Python to Interactive Brokers via
Trader Workstation (TWS) or IB Gateway, snapshots account and portfolio data,
and writes it into structured `IBKR_*` Excel sheets.

- One command, one workbook, no order placement.
- Paper account first. Never widen scope until reconciliation succeeds.
- Optional **tax workbook** mode refreshes `IBKR_*` tabs inside a Lei 14.754
  working copy and writes `IBKR_Reconciliacao` without touching fiscal sheets.

## What this MVP does

- Connects to TWS on `127.0.0.1:7497` (paper).
- Retrieves:
  - Managed accounts.
  - Account summary (NetLiq, Cash, Buying Power, P&L, margin metrics).
  - Cash and P&L broken out **by currency** (via `updateAccountValue`).
  - Positions (symbol, quantity, average cost, conId).
  - Portfolio (market price, market value, unrealized/realized P&L).
- Writes everything to `output/IBKR_Portfolio.xlsx` with named tables,
  header styling, and currency/quantity number formats.

## What this MVP explicitly avoids

- Live account.
- Order placement (no `placeOrder` code exists).
- Continuous streaming or scheduled runs.
- SQLite, Power Query, dashboards.
- Historical trades / dividends / withholding taxes — those belong in a
  separate Flex Query ingestion module.

## Project layout

```
IBKR-Excel-Bridge/
├── README.md
├── requirements.txt
├── .gitignore
├── config/
│   └── settings.json
├── docs/
├── logs/
├── output/
├── scripts/
│   ├── setup.zsh
│   └── run.zsh
├── src/
│   ├── ibkr_client.py
│   ├── excel_exporter.py
│   ├── reconcile.py
│   └── main.py
└── tests/
```

## Prerequisites

- macOS with Python 3.11 or newer.
- Trader Workstation (TWS) installed and logged into a **paper** account.
- IBKR Mac/Unix TWS API archive downloaded from
  [interactivebrokers.github.io](https://interactivebrokers.github.io/).

## TWS configuration

In TWS or IB Gateway, open API settings and set:

- **Enable ActiveX and Socket Clients**: ON
- **Socket port**: `7497` (TWS paper) or `4002` (Gateway paper)
- **Read-Only API**: ON
- **Allow connections from localhost only**: ON
- **Master API client ID**: leave blank (or set and do not use `21`)

Do not disable Read-Only API for this MVP. The `require_read_only_confirmation`
flag in `config/settings.json` is a self-check only — the real enforcement lives
in TWS/Gateway.

## Setup

```bash
cd /Users/eduardofgiovannini/Documents/GitHub/IBKR-Excel-Bridge
./scripts/setup.zsh
```

`setup.zsh` creates `.venv`, installs `openpyxl`, and prints the exact command
to install `ibapi` from the IBKR-distributed source tree. After extracting the
TWS API archive, run something like:

```bash
./.venv/bin/python -m pip install "/path/to/TWS API/source/pythonclient"
./.venv/bin/python -c "import ibapi; print(ibapi.__file__)"
```

This project is tested against `ibapi` 10.19 and later (the release that
introduced the `errorTime` argument to `EWrapper.error`). The collector
auto-detects either signature, but 10.19+ is recommended.

## Running

```bash
./scripts/run.zsh
```

`run.zsh` performs pre-flight checks (venv exists, packages import, TWS socket
responds) and then runs `python src/main.py`. On success you should see:

```
SUCCESS: /path/to/IBKR-Excel-Bridge/output/IBKR_Portfolio.xlsx
```

Logs are appended to `logs/ibkr_excel_bridge.log`.

## Workbook contents

Machine-owned sheets (rewritten every run):

| Worksheet | Contents |
| --- | --- |
| `IBKR_Overview` | Generation time, mode, accounts, counts, API message breakdown |
| `IBKR_Account_Summary` | NetLiq, cash, buying power, P&L, margin (base currency). `Value` is the IBKR-native string; `Value (Numeric)` is a coerced `float` for sorting and formulas. |
| `IBKR_Cash_By_Currency` | Per-currency ledger: cash balance, net-liq contribution, FX rate, accrued cash, P&L. Sourced from IBKR's `$LEDGER-*` account values. |
| `IBKR_Positions` | Symbol, `Kind` (Stock / FX Cash / Option / ...), quantity, average cost, currency, conId |
| `IBKR_Portfolio` | Market price, market value, unrealized/realized P&L per position. Excludes `CASH` secType (see limitations). |
| `IBKR_API_Messages` | Info, warnings, and errors from the API session, with a human-readable `Error Time (Local)` column |
| `IBKR_Reconciliacao` | **Tax mode only.** Qty diff vs `MyProfit_2026` / `Posicoes_Atuais` (`OK` / `DIVERGE` / `ONLY_IBKR` / `ONLY_FISCAL`). |

## Tax workbook mode (Phase 1)

Use this to refresh live IBKR data **inside** a Lei 14.754 working copy without mutating fiscal sheets.

1. Bootstrap the working file once:

```bash
cp "output/U6658119_TRIBUTACAO-LEI14754_v5-1-RECONCILIADO_2021-2026 copy.xlsx" \
   output/U6658119_TRIBUTACAO_WORKING.xlsx
```

2. In `config/settings.json` set:

```json
"excel": {
  "output_mode": "tax_workbook",
  "tax_workbook": "output/U6658119_TRIBUTACAO_WORKING.xlsx",
  "qty_tolerance": 0.0001
}
```

Optionally set `"expected_account": "U6658119"` under `ibkr` so a paper login cannot silently write into the tax file.

3. Prefer **IB Gateway** for API-only runs (lower resource use). Align `ibkr.port`:

| App | Paper | Live |
| --- | --- | --- |
| TWS | 7497 | 7496 |
| IB Gateway | 4002 | 4001 |

4. Run `./scripts/run.zsh`, then open `IBKR_Reconciliacao` before trusting any qty differences.

Default `output_mode` remains `standalone` (writes `output/IBKR_Portfolio.xlsx` only).

Design notes: `docs/superpowers/specs/2026-08-03-tax-workbook-api-integration-design.md`.

## Validation checklist

Before treating a run as authoritative:

- [ ] TWS Paper Trading is open and logged in
- [ ] API socket access enabled, port `7497`, Read-Only API on
- [ ] `./scripts/run.zsh` exits 0 with a `SUCCESS:` line
- [ ] Excel account number matches TWS
- [ ] `NetLiquidation` matches TWS within 0.5%
- [ ] Position quantities and conIds match the TWS Positions window
- [ ] Every position appears exactly once (no duplicates)
- [ ] `API Messages` contains only `info` rows (no `error` kind)

## Stop condition for Version 1

Version 1 is complete when:

1. One command generates the workbook.
2. Account balances match TWS.
3. Every current position appears once.
4. Quantity, average cost, and market value reconcile.
5. API remains read-only.
6. Three consecutive runs finish without manual code changes.

Only after all six is it worth moving on to Version 2 (trades, dividends,
withholding taxes, historical ledger ingestion via Flex Query).

## Known limitations

- Snapshot only. No transaction history, dividend detail, or withholding tax
  reporting — those require IBKR Flex Query, which is deliberately out of scope.
- Market values reflect the moment `updatePortfolio` last fired; they can shift
  slightly between the workbook and the TWS Portfolio window.
- `AccruedCash`, `Cushion`, and `EquityWithLoanValue` are not populated for
  every account type; empty values are expected in that case.
- **`Positions` can have more rows than `Portfolio`.** This is IBKR-defined
  behavior:
  - `reqPositions()` includes FX cash-conversion positions with
    `secType=CASH`. When you convert USD into EUR (or any non-base currency)
    inside TWS, that FX balance appears as a `CASH` position in `Positions`.
  - `reqAccountUpdates()` (which feeds `Portfolio`) does **not** re-emit those
    cash positions — the balance is already reflected in `TotalCashValue` and
    `$LEDGER-CashBalance`.
  - Do not reconcile the two sheets row-for-row. Reconcile by `Contract ID`
    within `Kind = Stock/Option/Future/...`, and expect `Positions` to be a
    superset when `Kind = FX Cash` rows exist.
- **`Positions.quantity` for an FX cash conversion is the gross conversion
  amount, not the settled balance.** For example, if you converted USD into
  40,000 EUR today and IBKR runs T+2 settlement, `Positions` shows `40,000`
  while `Cash By Currency > CashBalance` shows the settled figure (often
  lower). The two agree once settlement completes.
- **Average cost between `Positions` and `Portfolio` sheets can differ by a
  small amount for the same instrument.** This is IBKR-defined behavior:
  - The `Positions` sheet's `Average Cost` comes from `reqPositions()`, which
    returns the **native** (price-only) average cost.
  - The `Portfolio` sheet's `Average Cost` comes from `updatePortfolio()`,
    which returns the **commission-inclusive** average cost:
    `(execution_price × qty + commissions) / qty`.
  - The delta multiplied by quantity equals your round-trip commission for
    that position. For example, `(87.7979933 - 87.7828476) × 70 ≈ $1.06`.

### `Cash By Currency` sheet — how it's populated

IBKR routes per-currency ledger data through `updateAccountValue` callbacks
with keys of the form `$LEDGER-<metric>` and a real ISO currency code
(never `BASE`). The collector filters on that prefix and forwards these
metrics into the sheet:

- `CashBalance`, `TotalCashBalance` — settled cash balance in that currency
- `NetLiquidationByCurrency` — total value contributed by that currency
- `ExchangeRate` — IBKR's mark-to-base rate at snapshot time
- `AccruedCash`, `RealizedPnL`, `UnrealizedPnL` — currency-scoped
- `StockMarketValue`, `FxCashBalance` — for context

Plain `CashBalance` (without the `$LEDGER-` prefix) is deliberately dropped
because IBKR only ever emits it with segment suffixes (`-C`/`-S`/`-P`) in
the base currency, which duplicates `Account Summary` data.

## Troubleshooting

- **`IBKR API socket at … is not reachable`** — TWS/Gateway is not running,
  or API socket access is disabled, or the port in `settings.json` does not
  match (TWS paper `7497` / Gateway paper `4002`).
- **`Managed account request timed out`** — TWS accepted the connection but
  never returned `managedAccounts`. Usually caused by a stale connection; quit
  TWS fully and reopen it.
- **`ibapi` import fails** — you skipped the IBKR-distributed install. Rerun
  `./scripts/setup.zsh` and follow the printed pip command.
- **All error codes look wrong** — you're on `ibapi` < 10.19 with a mismatched
  `error()` signature. Upgrade `ibapi`; the client auto-detects either
  signature, but 10.19+ is the supported baseline.

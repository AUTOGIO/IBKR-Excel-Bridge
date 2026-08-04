# How to work in this repo

This repo is a **read-only bridge**. It never buys, sells, or edits fiscal costs.
You trade in TWS (or IBKR Client Portal). You declare tax in the Lei 14.754 workbook.
This project only **snapshots** live positions and **reconciles** them against fiscal sheets.

> **Personal values:** replace `<YOUR_ACCT>` with your real IBKR account id
> (e.g. `U0000000`) in `config/settings.local.json` — that file is gitignored
> and deep-merged over `config/settings.json` at runtime. Copy
> `config/settings.local.json.example` to get started. Never put your real
> account id in tracked files.

```text
YOU TRADE IN TWS          THIS REPO                    YOU REVIEW IN EXCEL
─────────────────         ──────────────               ────────────────────
Buy / sell ETF     →      ./scripts/run.zsh     →      IBKR_* sheets
                          (snapshot + recon)           IBKR_Reconciliacao
                                                       fiscal sheets unchanged
```

**Phase 1 (built):** live snapshot + qty reconciliation.  
**Phase 3 (built — CSV path):** drop Activity/Flex CSVs → `events.jsonl` → staging → promote into `Registro_Real`. Flex Web Service download is stubbed (`flex.enabled: false`).

---

## Two modes (pick one per run)

Change shared defaults in [`config/settings.json`](../config/settings.json); put
personal values (account id, custom tax-workbook filename, Flex token) in
[`config/settings.local.json`](../config/settings.local.json.example) — that
file is gitignored and deep-merged on top of the base settings at runtime.

| Mode | When | Output |
| --- | --- | --- |
| `standalone` (default) | Paper smoke tests, quick portfolio peek | `data/output/IBKR_Portfolio.xlsx` |
| `tax_workbook` | Real work against the Lei 14.754 file | `data/output/TRIBUTACAO_WORKING.xlsx` |

Rules of thumb:

1. Close the target workbook in Excel before running (lock files block writes).
2. Prefer **IB Gateway** for API-only runs (less CPU/RAM). Ports: Gateway paper `4002`, TWS paper `7497`.
3. Keep **Read-Only API** ON. This repo has no order code.
4. Set `"expected_account": "<YOUR_ACCT>"` only when TWS/Gateway is logged into that account.
5. Never hand-edit `IBKR_*` tabs — they are rewritten every run.
6. Fiscal truth stays in `MyProfit_2026`, `Posicoes_Atuais`, `Registro_Real`.

---

## Daily loop (after Phase 1 + Phase 3)

```bash
cd /Users/eduardofgiovannini/Documents/GitHub/IBKR-Excel-Bridge

# A. After new trades/dividends — export Flex/Activity CSV from IBKR Client Portal
#    and drop the file(s) into data/statements/

./scripts/ingest.zsh
#    → updates data/events.jsonl
#    → refreshes IBKR_Eventos_Staging + IBKR_Posicao_From_Events

# B. Review IBKR_Eventos_Staging in Excel (fill PTAX when ready)

./scripts/promote_events.zsh
#    → appends NEW rows only into Registro_Real (formulas copied)

# C. Snapshot live positions + qty recon
#    Prefer the in-sheet control, or close the workbook first:

./scripts/refresh_workbook.command
#    → closes the target workbook in Excel, runs snapshot, reopens it
#    Equivalent terminal path: ./scripts/run.zsh (workbook must be closed)

# D. Open workbook → IBKR_Overview (Refresh from TWS) / IBKR_Reconciliacao / Registro_Real
```

On `IBKR_Overview`, click the green **Refresh from TWS** cell to update all `IBKR_*` tabs. That launches `scripts/refresh_workbook.command`, which closes this file, snapshots TWS, and reopens it.

Optional later: set `"flex.enabled": true` plus token/query_id once Flex download is implemented; until then keep exporting CSVs manually.

Logs: `logs/ibkr_excel_bridge.log`.  
Tests after code changes: `./.venv/bin/python -m pytest tests -q`.

---

## Worked example: buy an ETF → last task this repo can do today

Scenario: you buy **100 shares of BIL** (US T-bill ETF) in account **<YOUR_ACCT>**, then use this repo through the last automated step (qty reconciliation). Fiscal event registration remains manual until Phase 3.

### Step 0 — One-time bootstrap (already done if the working file exists)

```bash
cp "archive/<YOUR_ACCT>_TRIBUTACAO-LEI14754_v5-1-RECONCILIADO_2021-2026 copy.xlsx" \
   data/output/TRIBUTACAO_WORKING.xlsx
```

Always work on `TRIBUTACAO_WORKING.xlsx`, not the dated `copy` file.

### Step 1 — Buy the ETF in TWS (outside this repo)

In TWS (live or paper — match the account you intend to reconcile):

1. Search **BIL**.
2. Buy **100** shares (market or limit).
3. Confirm fill in the Trades / Portfolio window.
4. Note: settlement is T+1 (US equities/ETFs after May 2024). Tax PTAX uses **settlement** date, not order date (`Spec_Motor_Python` R12).

This repo does **not** place that order.

### Step 2 — Point config at the tax workbook

In `config/settings.json`:

```json
{
  "ibkr": {
    "host": "127.0.0.1",
    "port": 7497,
    "client_id": 21,
    "expected_account": "<YOUR_ACCT>"
  },
  "excel": {
    "output_mode": "tax_workbook",
    "tax_workbook": "data/output/TRIBUTACAO_WORKING.xlsx",
    "qty_tolerance": 0.0001
  }
}
```

If you are still on **paper** (`DUR…`), leave `expected_account` empty or you will get a deliberate failure — that is correct.

For Gateway paper, use `"port": 4002`.

### Step 3 — Snapshot + reconcile (last automated task in the repo today)

```bash
# Close Excel first if the working file is open
./scripts/run.zsh
```

Success looks like:

```text
SUCCESS: .../data/output/TRIBUTACAO_WORKING.xlsx
```

What the run did:

| Sheet | Effect |
| --- | --- |
| `IBKR_Positions` / `IBKR_Portfolio` | Live BIL qty/price from TWS |
| `IBKR_Reconciliacao` | Compares live qty vs `MyProfit_2026` |
| `MyProfit_2026`, `Registro_Real`, `Apuracao_Anual`, … | **Untouched** |

### Step 4 — Read `IBKR_Reconciliacao` (your job)

Open the working workbook → tab **IBKR_Reconciliacao**.

Typical statuses after buying BIL:

| Status | Meaning | What you do |
| --- | --- | --- |
| `ONLY_IBKR` for `BIL` | Live has BIL; fiscal sheet does not (or qty missing) | Expected until you update fiscal books |
| `DIVERGE` for `BIL` | Both sides exist but qty differs | Check fill qty vs MyProfit / Posicoes |
| `OK` | Qtys match within tolerance | Nothing |
| `ONLY_FISCAL` | Fiscal has a ticker TWS does not | Sold elsewhere, alias, or stale MyProfit |

Paper vs live tax account will show many `ONLY_FISCAL` / `ONLY_IBKR` rows — that is a **wrong-account** signal, not a BIL bug.

### Step 5 — Ingest the Activity / Flex CSV (automated)

1. In IBKR Client Portal, export an Activity Statement or Flex Query as **CSV**.
2. Drop the file into `data/statements/` (any `*.csv` name).
3. Close the tax workbook in Excel, then:

```bash
./scripts/ingest.zsh
```

4. Open `IBKR_Eventos_Staging` — confirm the Compra for BIL. Fill **PTAX** in staging or later in `Registro_Real` (blank PTAX is flagged by Excel validation).
5. Promote into the fiscal ledger:

```bash
./scripts/promote_events.zsh
```

6. Optionally copy quantities/costs from `IBKR_Posicao_From_Events` into `MyProfit_2026` when you trust them (never auto-overwritten).

### Step 6 — Re-snapshot and confirm recon

```bash
./scripts/run.zsh
```

`BIL` should move toward `OK` on `IBKR_Reconciliacao` once MyProfit qty matches live.

### Step 7 — Excel apuração (product tasks)

With the event and costs updated:

1. `Posicoes_Atuais` — confirm BIL qty/cost.
2. `Simulador_Vendas` / `Cenarios_Venda` — optional what-if sells.
3. `Apuracao_Anual` — annual Lei 14.754 base / DARF estimate.
4. `Auditoria_Dados` / `Relatorio_Final` — clear blockers before DIRPF.

End-to-end automated path now: **trade → CSV drop → ingest → review → promote → snapshot → recon**.

---

## What remains for Flex API

When `src/flex_client.py` is fully implemented, Step 5 download becomes automatic (`flex.enabled` + token + query id). The CSV parser and promote path stay the same.

---

## Quick troubleshooting

| Symptom | Fix |
| --- | --- |
| Socket not reachable | Open TWS/Gateway; match `port` |
| Workbook appears open | Close Excel; remove stale `data/output/~$*` only if Excel is quit |
| `expected_account` error | Log into the right account or clear the setting |
| Everything `ONLY_FISCAL` | You snapshotted paper against the live tax book |
| Want portfolio only | `"output_mode": "standalone"` |

---

## Mental checklist before any “real” tax run

- [ ] Logged into **<YOUR_ACCT>** (not paper)
- [ ] Read-Only API on
- [ ] `output_mode` = `tax_workbook`
- [ ] `expected_account` = `<YOUR_ACCT>`
- [ ] Working xlsx closed in Excel
- [ ] After run: `IBKR_Reconciliacao` reviewed
- [ ] New trades reflected in `Registro_Real` + MyProfit before trusting DARF

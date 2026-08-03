# How to work in this repo

This repo is a **read-only bridge**. It never buys, sells, or edits fiscal costs.
You trade in TWS (or IBKR Client Portal). You declare tax in the Lei 14.754 workbook.
This project only **snapshots** live positions and **reconciles** them against fiscal sheets.

```text
YOU TRADE IN TWS          THIS REPO                    YOU REVIEW IN EXCEL
─────────────────         ──────────────               ────────────────────
Buy / sell ETF     →      ./scripts/run.zsh     →      IBKR_* sheets
                          (snapshot + recon)           IBKR_Reconciliacao
                                                       fiscal sheets unchanged
```

**Phase 1 (built):** live snapshot + qty reconciliation.  
**Phase 3 (not built yet):** Flex/statement → `Registro_Real`. Until then, new trades are entered in Excel (or left pending) using Activity Statements.

---

## Two modes (pick one per run)

Edit only [`config/settings.json`](../config/settings.json).

| Mode | When | Output |
| --- | --- | --- |
| `standalone` (default) | Paper smoke tests, quick portfolio peek | `output/IBKR_Portfolio.xlsx` |
| `tax_workbook` | Real work against the Lei 14.754 file | `output/U6658119_TRIBUTACAO_WORKING.xlsx` |

Rules of thumb:

1. Close the target workbook in Excel before running (lock files block writes).
2. Prefer **IB Gateway** for API-only runs (less CPU/RAM). Ports: Gateway paper `4002`, TWS paper `7497`.
3. Keep **Read-Only API** ON. This repo has no order code.
4. Set `"expected_account": "U6658119"` only when TWS/Gateway is logged into that account.
5. Never hand-edit `IBKR_*` tabs — they are rewritten every run.
6. Fiscal truth stays in `MyProfit_2026`, `Posicoes_Atuais`, `Registro_Real`.

---

## Daily loop (after Phase 1)

```bash
cd /Users/eduardofgiovannini/Documents/GitHub/IBKR-Excel-Bridge

# 1. TWS or Gateway open, API on, correct account logged in
# 2. Close the Excel target file
# 3. Config: output_mode + expected_account as needed

./scripts/run.zsh

# 4. Open the workbook → IBKR_Reconciliacao
# 5. Resolve ONLY_IBKR / DIVERGE / ONLY_FISCAL manually
# 6. Use Excel simulators / Apuracao_Anual as before
```

Logs: `logs/ibkr_excel_bridge.log`.  
Tests after code changes: `./.venv/bin/python -m pytest tests -q`.

---

## Worked example: buy an ETF → last task this repo can do today

Scenario: you buy **100 shares of BIL** (US T-bill ETF) in account **U6658119**, then use this repo through the last automated step (qty reconciliation). Fiscal event registration remains manual until Phase 3.

### Step 0 — One-time bootstrap (already done if the working file exists)

```bash
cp "output/U6658119_TRIBUTACAO-LEI14754_v5-1-RECONCILIADO_2021-2026 copy.xlsx" \
   output/U6658119_TRIBUTACAO_WORKING.xlsx
```

Always work on `U6658119_TRIBUTACAO_WORKING.xlsx`, not the dated `copy` file.

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
    "expected_account": "U6658119"
  },
  "excel": {
    "output_mode": "tax_workbook",
    "tax_workbook": "output/U6658119_TRIBUTACAO_WORKING.xlsx",
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
SUCCESS: .../output/U6658119_TRIBUTACAO_WORKING.xlsx
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

### Step 5 — Manual fiscal update (still outside automation)

Until Phase 3 exists, record the purchase yourself:

1. Download IBKR **Activity Statement** (or Flex) for the trade day.
2. In `Registro_Real`, add a **Compra** row: date, `BIL`, qty, USD price, PTAX on **settlement** date, etc. (follow existing row patterns / TOC).
3. Refresh / update `MyProfit_2026` (or your cost source) so fiscal qty/cost for BIL matches reality.
4. Re-run `./scripts/run.zsh` → `BIL` should move to `OK` on `IBKR_Reconciliacao`.

### Step 6 — Last *product* tasks in the Excel model (not Python yet)

With the event and costs updated:

1. `Posicoes_Atuais` — confirm BIL qty/cost (formulas pull from MyProfit today).
2. `Simulador_Vendas` / `Cenarios_Venda` — optional what-if sells.
3. `Apuracao_Anual` — annual Lei 14.754 base / DARF estimate.
4. `Auditoria_Dados` / `Relatorio_Final` — clear blockers before DIRPF.

That is the end of the current operating path: **trade → snapshot → reconcile → manual ledger → Excel apuração**.

---

## What this repo will add next (Phase 3)

When built, Step 5 becomes:

```text
Flex CSV / Activity Statement
  → ingest → data/events.jsonl
  → IBKR_Eventos_Staging (review)
  → append to Registro_Real
```

Do not start Phase 3 until live `IBKR_Reconciliacao` is trustworthy on **U6658119**.

---

## Quick troubleshooting

| Symptom | Fix |
| --- | --- |
| Socket not reachable | Open TWS/Gateway; match `port` |
| Workbook appears open | Close Excel; remove stale `output/~$*` only if Excel is quit |
| `expected_account` error | Log into the right account or clear the setting |
| Everything `ONLY_FISCAL` | You snapshotted paper against the live tax book |
| Want portfolio only | `"output_mode": "standalone"` |

---

## Mental checklist before any “real” tax run

- [ ] Logged into **U6658119** (not paper)
- [ ] Read-Only API on
- [ ] `output_mode` = `tax_workbook`
- [ ] `expected_account` = `U6658119`
- [ ] Working xlsx closed in Excel
- [ ] After run: `IBKR_Reconciliacao` reviewed
- [ ] New trades reflected in `Registro_Real` + MyProfit before trusting DARF

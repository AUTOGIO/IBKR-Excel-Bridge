# Phase 1 Tax Workbook IBKR Snapshot Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Refresh machine-owned `IBKR_*` sheets inside a Lei 14.754 working workbook and write `IBKR_Reconciliacao` without touching fiscal sheets.

**Architecture:** Keep `ibkr_client` unchanged. Add pure `reconcile.py` for qty diffs. Extend `ExcelExporter` with `IBKR_` owned sheets, strict tax-workbook load, and recon sheet. Wire `main.py` to `output_mode` / `expected_account` / Gateway-friendly ports via config only.

**Tech Stack:** Python 3.11+, openpyxl, pytest, existing ibapi client.

**Spec:** `docs/superpowers/specs/2026-08-03-tax-workbook-api-integration-design.md`

## Global Constraints

- Read-only API posture unchanged; no order placement.
- Never rewrite fiscal sheets (`MyProfit_2026`, `Posicoes_Atuais`, `Registro_Real`, …).
- Default `excel.output_mode` remains `standalone`.
- Tax mode refuses missing/corrupt workbook (no blank fallback).
- Commits only when the user explicitly asks (do not auto-commit).

## File structure

| File | Responsibility |
| --- | --- |
| `src/reconcile.py` | Symbol normalize + qty reconciliation rows |
| `src/excel_exporter.py` | Owned `IBKR_*` sheets, tax load rules, write recon |
| `src/main.py` | Resolve path, enforce expected account, pass tolerance |
| `config/settings.json` | New keys |
| `tests/test_reconcile.py` | Pure recon unit tests |
| `tests/test_excel_exporter_tax.py` | Exporter tax/standalone behavior |
| `README.md` | Tax mode + Gateway ports + bootstrap |
| `scripts/run.zsh` | Friendlier socket error mentioning Gateway |

---

### Task 1: Reconciliation helpers

**Files:**
- Create: `src/reconcile.py`
- Test: `tests/test_reconcile.py`

**Interfaces:**
- Produces:
  - `normalize_symbol(value: Any) -> str`
  - `reconcile_quantities(*, live: list[dict[str, Any]], fiscal: list[dict[str, Any]], tolerance: float = 0.0001) -> list[dict[str, Any]]`
  - Live dict keys used: `symbol`, `quantity`, `instrument_kind` (skip when kind is `FX Cash`)
  - Fiscal dict keys used: `symbol`, `quantity`
  - Output row keys: `symbol`, `qty_ibkr`, `qty_fiscal`, `delta`, `status` where status ∈ `OK|DIVERGE|ONLY_IBKR|ONLY_FISCAL`

- [ ] **Step 1: Write failing tests**

```python
# tests/test_reconcile.py
from reconcile import normalize_symbol, reconcile_quantities

def test_normalize_symbol_trims_and_uppercases():
    assert normalize_symbol("  amzn ") == "AMZN"

def test_reconcile_statuses_and_tolerance():
    live = [
        {"symbol": "AMZN", "quantity": 7.0, "instrument_kind": "Stock"},
        {"symbol": "EUR", "quantity": 1000.0, "instrument_kind": "FX Cash"},
        {"symbol": "NEW", "quantity": 1.0, "instrument_kind": "Stock"},
    ]
    fiscal = [
        {"symbol": "amzn", "quantity": 7.0},
        {"symbol": "BIL", "quantity": 100.0},
        {"symbol": "CLSK", "quantity": 305.00005},  # within default tol vs absent live → ONLY_FISCAL
    ]
    rows = {r["symbol"]: r for r in reconcile_quantities(live=live, fiscal=fiscal)}
    assert "EUR" not in rows
    assert rows["AMZN"]["status"] == "OK"
    assert rows["NEW"]["status"] == "ONLY_IBKR"
    assert rows["BIL"]["status"] == "ONLY_FISCAL"

def test_diverge_above_tolerance():
    live = [{"symbol": "BIL", "quantity": 100.0, "instrument_kind": "Stock"}]
    fiscal = [{"symbol": "BIL", "quantity": 99.0}]
    rows = reconcile_quantities(live=live, fiscal=fiscal, tolerance=0.0001)
    assert rows[0]["status"] == "DIVERGE"
    assert rows[0]["delta"] == 1.0
```

- [ ] **Step 2: Run tests — expect FAIL** (module missing)

Run: `./.venv/bin/python -m pytest tests/test_reconcile.py -v`

- [ ] **Step 3: Implement `src/reconcile.py`**

```python
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

    symbols = sorted(set(live_map) | set(fiscal_map))
    out: list[dict[str, Any]] = []
    for symbol in symbols:
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
```

- [ ] **Step 4: Run tests — expect PASS**

Run: `./.venv/bin/python -m pytest tests/test_reconcile.py -v`

---

### Task 2: ExcelExporter tax mode + IBKR_ sheets

**Files:**
- Modify: `src/excel_exporter.py`
- Test: `tests/test_excel_exporter_tax.py`

**Interfaces:**
- Consumes: `reconcile_quantities`, `normalize_symbol`
- Produces:
  - `ExcelExporter(output_path, *, mode: str = "standalone", qty_tolerance: float = 0.0001)`
  - `OWNED_SHEETS` = IBKR_ prefixed tuple including `IBKR_Reconciliacao`
  - `LEGACY_OWNED_SHEETS` = old unprefixed names (deleted when present)
  - `export(data)` behavior:
    - `standalone`: create/fallback OK; write owned sheets; **no** recon sheet content required (still include empty recon or omit recon from standalone — **omit `IBKR_Reconciliacao` from standalone writes**; do not add it to standalone OWNED rewrite set… Spec says recon only in tax mode. Cleanest: always list recon in OWNED for tax; for standalone, OWNED excludes recon.)
  - Spec clarity lock: **standalone OWNED** = six IBKR_ data sheets (no Reconciliacao). **tax OWNED** = those six + `IBKR_Reconciliacao`. Both delete `LEGACY_OWNED_SHEETS`.
  - Tax mode: file must exist; corrupt load raises; never blank fallback
  - Tax mode reads fiscal qty from `MyProfit_2026` (`Ativo normalizado`, `Quantidade`) else `Posicoes_Atuais` (`Ativo`, `Quantidade_Atual`) using `data_only=True` pass for values when needed — implement by loading once with formulas for rewrite and a second `data_only=True` load only to extract fiscal maps (read-only), then apply writes on the formula workbook.

- [ ] **Step 1: Write failing tests** for preserve foreign sheets, IBKR_ names, missing tax file error, recon statuses written

```python
# tests/test_excel_exporter_tax.py
from pathlib import Path
from openpyxl import Workbook, load_workbook
from excel_exporter import ExcelExporter

def _sample_data():
    return {
        "accounts": [{"account": "U6658119"}],
        "account_summary": [],
        "account_values": [],
        "positions": [
            {"account": "U6658119", "symbol": "AMZN", "instrument_kind": "Stock",
             "security_type": "STK", "currency": "USD", "exchange": "NASDAQ",
             "primary_exchange": "NASDAQ", "contract_id": 1, "quantity": 7,
             "average_cost": 1.0},
            {"account": "U6658119", "symbol": "EUR", "instrument_kind": "FX Cash",
             "security_type": "CASH", "currency": "USD", "exchange": "IDEALPRO",
             "primary_exchange": "", "contract_id": 2, "quantity": 1000,
             "average_cost": 0},
        ],
        "portfolio": [],
        "errors": [],
    }

def test_standalone_uses_ibkr_prefix(tmp_path: Path):
    out = tmp_path / "IBKR_Portfolio.xlsx"
    ExcelExporter(out, mode="standalone").export(_sample_data())
    wb = load_workbook(out)
    assert "IBKR_Positions" in wb.sheetnames
    assert "Positions" not in wb.sheetnames
    assert "IBKR_Reconciliacao" not in wb.sheetnames

def test_tax_mode_preserves_foreign_and_writes_recon(tmp_path: Path):
    path = tmp_path / "tax.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "MyProfit_2026"
    ws["A4"] = "Nome"
    ws["B4"] = "Ativo normalizado"
    ws["C4"] = "Quantidade"
    ws["B5"] = "AMZN"
    ws["C5"] = 7
    ws["B6"] = "BIL"
    ws["C6"] = 100
    wb.create_sheet("Registro_Real")["A1"] = "keep-me"
    wb.save(path)
    ExcelExporter(path, mode="tax_workbook", qty_tolerance=0.0001).export(_sample_data())
    wb2 = load_workbook(path)
    assert wb2["Registro_Real"]["A1"].value == "keep-me"
    assert "IBKR_Positions" in wb2.sheetnames
    recon = {row[0].value: row[4].value for row in wb2["IBKR_Reconciliacao"].iter_rows(min_row=2, max_col=5)}
    # header humanized — assert via table values by scanning
    statuses = {}
    sheet = wb2["IBKR_Reconciliacao"]
    headers = [c.value for c in next(sheet.iter_rows(min_row=1, max_row=1))]
    # After humanize: Symbol, Qty Ibkr, Qty Fiscal, Delta, Status
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0]:
            statuses[row[0]] = row[4]
    assert statuses["AMZN"] == "OK"
    assert statuses["BIL"] == "ONLY_FISCAL"

def test_tax_mode_missing_file_raises(tmp_path: Path):
    missing = tmp_path / "nope.xlsx"
    try:
        ExcelExporter(missing, mode="tax_workbook").export(_sample_data())
        assert False, "expected FileNotFoundError"
    except FileNotFoundError:
        pass
```

- [ ] **Step 2: Run — expect FAIL**

Run: `./.venv/bin/python -m pytest tests/test_excel_exporter_tax.py -v`

- [ ] **Step 3: Implement exporter changes**

Key constants and constructor:

```python
OWNED_SHEETS_STANDALONE = (
    "IBKR_Overview",
    "IBKR_Account_Summary",
    "IBKR_Cash_By_Currency",
    "IBKR_Positions",
    "IBKR_Portfolio",
    "IBKR_API_Messages",
)
OWNED_SHEETS_TAX = OWNED_SHEETS_STANDALONE + ("IBKR_Reconciliacao",)
LEGACY_OWNED_SHEETS = (
    "Overview", "Account Summary", "Cash By Currency",
    "Positions", "Portfolio", "API Messages",
)
```

Update `_write_overview` to use sheet `IBKR_Overview`. Update all `_write_records` sheet names. In tax mode after writing positions, extract fiscal rows and write recon via `_write_records(..., "IBKR_Reconciliacao", ...)`.

Fiscal extract helper (same file or reconcile):

```python
def extract_fiscal_quantities(workbook_path: Path) -> list[dict[str, Any]]:
    wb = load_workbook(workbook_path, data_only=True, read_only=True)
    try:
        if "MyProfit_2026" in wb.sheetnames:
            ws = wb["MyProfit_2026"]
            rows = []
            for row in ws.iter_rows(min_row=5, values_only=True):
                symbol, qty = row[1], row[2]  # B, C
                if symbol:
                    rows.append({"symbol": symbol, "quantity": qty})
            return rows
        if "Posicoes_Atuais" in wb.sheetnames:
            ws = wb["Posicoes_Atuais"]
            rows = []
            for row in ws.iter_rows(min_row=5, values_only=True):
                symbol, qty = row[0], row[3]  # A, D
                if symbol:
                    rows.append({"symbol": symbol, "quantity": qty})
            return rows
        return []
    finally:
        wb.close()
```

- [ ] **Step 4: Run — expect PASS**

Run: `./.venv/bin/python -m pytest tests/test_excel_exporter_tax.py tests/test_reconcile.py -v`

---

### Task 3: Wire `main.py` + settings

**Files:**
- Modify: `src/main.py`
- Modify: `config/settings.json`
- Test: `tests/test_main_config.py` (pure helpers — extract small functions to keep testable without TWS)

**Interfaces:**
- Produces:
  - `resolve_output_path(excel_config: dict, project_root: Path) -> Path`
  - `assert_expected_account(data: dict, expected: str) -> None` raises `ValueError` on mismatch

```python
def resolve_output_path(excel_config: dict, project_root: Path) -> Path:
    mode = excel_config.get("output_mode", "standalone")
    if mode == "tax_workbook":
        rel = excel_config["tax_workbook"]
    elif mode == "standalone":
        rel = excel_config["output_file"]
    else:
        raise ValueError(f"Unknown excel.output_mode: {mode!r}")
    return project_root / rel

def assert_expected_account(data: dict, expected: str) -> None:
    expected = (expected or "").strip()
    if not expected:
        return
    accounts = {row.get("account") for row in data.get("accounts", [])}
    if expected not in accounts:
        raise ValueError(
            f"expected_account {expected!r} not in managed accounts {sorted(a for a in accounts if a)}"
        )
```

Call order in `main`: collect → assert_expected_account → ExcelExporter(...).export

`settings.json` defaults:

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
  "logging": { "level": "INFO" }
}
```

- [ ] **Step 1–4:** TDD the two helpers; wire `main.py`; update settings; run pytest suite.

---

### Task 4: README + run.zsh messaging

**Files:**
- Modify: `README.md`
- Modify: `scripts/run.zsh`

- [ ] Document tax workbook mode, bootstrap copy command, TWS vs Gateway ports, `expected_account`.
- [ ] `run.zsh` error text: mention TWS **or** IB Gateway and ports 7497/4002.

Bootstrap command to document:

```bash
cp "output/U6658119_TRIBUTACAO-LEI14754_v5-1-RECONCILIADO_2021-2026 copy.xlsx" \
   output/U6658119_TRIBUTACAO_WORKING.xlsx
```

Then set `"output_mode": "tax_workbook"` and optionally `"expected_account": "U6658119"`.

---

### Task 5: Full regression

- [ ] Run: `./.venv/bin/python -m pytest tests -q`
- [ ] Expected: all pass (including existing ibkr_client smoke tests)

---

## Spec coverage check

| Spec requirement | Task |
| --- | --- |
| IBKR_ owned sheets | 2 |
| Preserve fiscal sheets | 2 |
| IBKR_Reconciliacao | 1+2 |
| Skip FX Cash in recon | 1 |
| MyProfit then Posicoes fallback | 2 |
| output_mode standalone default | 3 |
| tax file must exist / no blank fallback | 2 |
| expected_account | 3 |
| Gateway ports documented | 4 |
| Legacy unprefixed delete | 2 |
| Phase 3 | deferred (not in this plan) |

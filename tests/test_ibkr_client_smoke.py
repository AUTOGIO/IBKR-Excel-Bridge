"""Smoke tests for ``ibkr_client`` using a stub ``ibapi`` package.

These tests validate:

- The module imports without a real IBKR install.
- Portfolio updates upsert on ``(account, conId)`` (no duplicate rows).
- ``updateAccountValue`` filters to the currency-bearing keys we care about.
- ``error()`` correctly parses arguments for both the pre-10.19 and
  10.19+ ibapi signatures.
- The informational error-code set is applied.

Run with:
  ./.venv/bin/python -m pytest tests -q
or directly:
  ./.venv/bin/python tests/test_ibkr_client_smoke.py
"""

from __future__ import annotations

import sys
import types
from decimal import Decimal
from pathlib import Path


def _install_ibapi_stub(with_error_time: bool) -> None:
    """Inject a fake ``ibapi`` into sys.modules before importing ibkr_client."""
    for name in [
        "ibapi",
        "ibapi.client",
        "ibapi.contract",
        "ibapi.wrapper",
        "ibkr_client",
    ]:
        sys.modules.pop(name, None)

    ibapi = types.ModuleType("ibapi")
    ibapi_client = types.ModuleType("ibapi.client")
    ibapi_contract = types.ModuleType("ibapi.contract")
    ibapi_wrapper = types.ModuleType("ibapi.wrapper")

    class EClient:
        def __init__(self, wrapper):  # noqa: D401
            self._wrapper = wrapper

        def connect(self, *args, **kwargs) -> None:  # noqa: D401
            return None

        def disconnect(self) -> None:
            return None

        def isConnected(self) -> bool:  # noqa: N802
            return False

        def run(self) -> None:
            return None

        # Request methods used by IBKRClient — all no-ops in the stub.
        def reqManagedAccts(self) -> None:  # noqa: N802
            return None

        def reqAccountSummary(self, *args, **kwargs) -> None:  # noqa: N802
            return None

        def cancelAccountSummary(self, *args, **kwargs) -> None:  # noqa: N802
            return None

        def reqPositions(self) -> None:  # noqa: N802
            return None

        def cancelPositions(self) -> None:  # noqa: N802
            return None

        def reqAccountUpdates(self, *args, **kwargs) -> None:  # noqa: N802
            return None

    class Contract:
        def __init__(self):
            self.symbol = ""
            self.secType = ""
            self.currency = ""
            self.exchange = ""
            self.primaryExchange = ""
            self.conId = 0

    if with_error_time:

        class EWrapper:
            def error(
                self,
                reqId: int,
                errorTime: int,
                errorCode: int,
                errorString: str,
                advancedOrderRejectJson: str = "",
            ) -> None:
                return None
    else:

        class EWrapper:  # type: ignore[no-redef]
            def error(
                self,
                reqId: int,
                errorCode: int,
                errorString: str,
                advancedOrderRejectJson: str = "",
            ) -> None:
                return None

    ibapi_client.EClient = EClient
    ibapi_contract.Contract = Contract
    ibapi_wrapper.EWrapper = EWrapper

    sys.modules["ibapi"] = ibapi
    sys.modules["ibapi.client"] = ibapi_client
    sys.modules["ibapi.contract"] = ibapi_contract
    sys.modules["ibapi.wrapper"] = ibapi_wrapper


def _load_client_module(with_error_time: bool):
    _install_ibapi_stub(with_error_time=with_error_time)
    src = Path(__file__).resolve().parent.parent / "src"
    if str(src) not in sys.path:
        sys.path.insert(0, str(src))
    import ibkr_client  # noqa: WPS433 - dynamic import intentional

    return ibkr_client


def _fake_contract(conId: int, symbol: str, currency: str = "USD"):  # noqa: N803
    contract = sys.modules["ibapi.contract"].Contract()
    contract.symbol = symbol
    contract.secType = "STK"
    contract.currency = currency
    contract.exchange = "SMART"
    contract.primaryExchange = "NASDAQ"
    contract.conId = conId
    return contract


def test_portfolio_upserts_on_repeated_updates() -> None:
    mod = _load_client_module(with_error_time=True)
    client = mod.IBKRClient()

    contract = _fake_contract(265598, "AAPL")

    client.updatePortfolio(contract, Decimal("10"), 180.0, 1800.0, 175.0, 50.0, 0.0, "DU1")
    client.updatePortfolio(contract, Decimal("10"), 181.0, 1810.0, 175.0, 60.0, 0.0, "DU1")
    client.updatePortfolio(contract, Decimal("10"), 182.0, 1820.0, 175.0, 70.0, 0.0, "DU1")

    snapshot = client._snapshot()
    assert len(snapshot["portfolio"]) == 1, snapshot["portfolio"]
    assert snapshot["portfolio"][0]["market_price"] == 182.0
    assert snapshot["portfolio"][0]["unrealized_pnl"] == 70.0


def test_portfolio_distinct_accounts_and_contracts_kept_separate() -> None:
    mod = _load_client_module(with_error_time=True)
    client = mod.IBKRClient()

    a = _fake_contract(1, "AAPL")
    b = _fake_contract(2, "MSFT")

    client.updatePortfolio(a, Decimal("1"), 1.0, 1.0, 1.0, 0.0, 0.0, "DU1")
    client.updatePortfolio(b, Decimal("1"), 2.0, 2.0, 2.0, 0.0, 0.0, "DU1")
    client.updatePortfolio(a, Decimal("1"), 1.0, 1.0, 1.0, 0.0, 0.0, "DU2")

    snapshot = client._snapshot()
    assert len(snapshot["portfolio"]) == 3


def test_update_account_value_uses_ledger_prefix() -> None:
    """Real IBKR feed uses ``$LEDGER-<metric>`` keys for per-currency data.
    Plain ``CashBalance`` never carries a non-BASE currency, so the filter
    must key on the ``$LEDGER-`` prefix."""
    mod = _load_client_module(with_error_time=True)
    client = mod.IBKRClient()

    # $LEDGER-* with real currency: keep.
    client.updateAccountValue("$LEDGER-CashBalance", "38000.00", "EUR", "DU1")
    client.updateAccountValue("$LEDGER-TotalCashBalance", "38000.00", "EUR", "DU1")
    client.updateAccountValue("$LEDGER-CashBalance", "159347.20", "USD", "DU1")
    client.updateAccountValue("$LEDGER-ExchangeRate", "1.1508", "EUR", "DU1")
    # $LEDGER-* with BASE: drop.
    client.updateAccountValue("$LEDGER-CashBalance", "203081.02", "BASE", "DU1")
    # Non-$LEDGER key: drop (base-currency aggregates duplicate Account Summary).
    client.updateAccountValue("CashBalance", "12345.67", "USD", "DU1")
    # $LEDGER-* metric not in our whitelist: drop.
    client.updateAccountValue("$LEDGER-Cryptocurrency", "0.00", "EUR", "DU1")

    snap = client._snapshot()
    triples = {(r["currency"], r["metric"], r["value"]) for r in snap["account_values"]}
    assert ("EUR", "CashBalance", "38000.00") in triples
    assert ("EUR", "TotalCashBalance", "38000.00") in triples
    assert ("USD", "CashBalance", "159347.20") in triples
    assert ("EUR", "ExchangeRate", "1.1508") in triples
    assert not any(currency == "BASE" for currency, *_ in triples)
    assert not any(metric == "Cryptocurrency" for _, metric, _ in triples)
    # Every kept row must have a parseable numeric value.
    for r in snap["account_values"]:
        assert r["value_numeric"] is not None


def test_error_new_signature_parses_code_and_message() -> None:
    mod = _load_client_module(with_error_time=True)
    client = mod.IBKRClient()

    client.error(-1, 1700000000, 2104, "Market data farm connection is OK:usfarm")

    snapshot = client._snapshot()
    assert len(snapshot["errors"]) == 1
    assert snapshot["errors"][0]["code"] == 2104
    assert snapshot["errors"][0]["kind"] == "info"
    assert "usfarm" in snapshot["errors"][0]["message"]


def test_error_legacy_signature_parses_code_and_message() -> None:
    mod = _load_client_module(with_error_time=False)
    client = mod.IBKRClient()

    client.error(-1, 2158, "Sec-def data farm connection is OK:secdefil")

    snapshot = client._snapshot()
    assert len(snapshot["errors"]) == 1
    assert snapshot["errors"][0]["code"] == 2158
    assert snapshot["errors"][0]["kind"] == "info"


def test_error_classifies_client_errors_as_error_kind() -> None:
    mod = _load_client_module(with_error_time=True)
    client = mod.IBKRClient()

    client.error(1, 1700000000, 200, "No security definition has been found")

    snapshot = client._snapshot()
    assert snapshot["errors"][0]["code"] == 200
    assert snapshot["errors"][0]["kind"] == "error"


def test_snapshot_shape_is_stable_empty() -> None:
    mod = _load_client_module(with_error_time=True)
    client = mod.IBKRClient()

    snapshot = client._snapshot()
    assert set(snapshot.keys()) == {
        "accounts",
        "account_summary",
        "account_values",
        "positions",
        "portfolio",
        "errors",
    }
    assert all(isinstance(v, list) for v in snapshot.values())


def test_account_summary_dedupes_dual_dispatch() -> None:
    """ibapi 10.19+ fires the legacy and *ProtoBuf callback for the same
    server message. Handler must upsert, not append."""
    mod = _load_client_module(with_error_time=True)
    client = mod.IBKRClient()

    for _ in range(2):  # simulate dual-dispatch
        client.accountSummary(9001, "DU1", "NetLiquidation", "202588.97", "USD")
        client.accountSummary(9001, "DU1", "TotalCashValue", "196948.30", "USD")
        client.accountSummary(9001, "DU1", "AccountType", "INDIVIDUAL", "")

    snap = client._snapshot()
    assert len(snap["account_summary"]) == 3, snap["account_summary"]
    values = {(r["tag"], r["value"]) for r in snap["account_summary"]}
    assert ("NetLiquidation", "202588.97") in values
    assert ("AccountType", "INDIVIDUAL") in values


def test_positions_dedupe_dual_dispatch() -> None:
    mod = _load_client_module(with_error_time=True)
    client = mod.IBKRClient()

    contract = _fake_contract(265598, "AAPL")

    for _ in range(2):
        client.position("DU1", contract, Decimal("10"), 175.42)

    snap = client._snapshot()
    assert len(snap["positions"]) == 1
    assert snap["positions"][0]["symbol"] == "AAPL"
    assert snap["positions"][0]["quantity"] == 10.0


def test_account_values_dedupe_dual_dispatch() -> None:
    mod = _load_client_module(with_error_time=True)
    client = mod.IBKRClient()

    for _ in range(2):
        client.updateAccountValue("$LEDGER-CashBalance", "12345.67", "EUR", "DU1")
        client.updateAccountValue("$LEDGER-CashBalance", "9876.54", "USD", "DU1")

    snap = client._snapshot()
    assert len(snap["account_values"]) == 2
    pairs = {(r["currency"], r["value"]) for r in snap["account_values"]}
    assert pairs == {("EUR", "12345.67"), ("USD", "9876.54")}


def test_account_summary_carries_numeric_and_text_value() -> None:
    """Both the raw string and a coerced float should be stored so Excel can
    sort and sum the numeric column without lexicographic surprises."""
    mod = _load_client_module(with_error_time=True)
    client = mod.IBKRClient()

    client.accountSummary(9001, "DU1", "NetLiquidation", "202588.97", "USD")
    client.accountSummary(9001, "DU1", "AccountType", "INDIVIDUAL", "")
    client.accountSummary(9001, "DU1", "Cushion", "0.973595", "")

    snap = {r["tag"]: r for r in client._snapshot()["account_summary"]}
    assert snap["NetLiquidation"]["value"] == "202588.97"
    assert snap["NetLiquidation"]["value_numeric"] == 202588.97
    assert snap["Cushion"]["value_numeric"] == 0.973595
    # Non-numeric tags: keep the string, but value_numeric is None.
    assert snap["AccountType"]["value"] == "INDIVIDUAL"
    assert snap["AccountType"]["value_numeric"] is None


def test_position_and_portfolio_carry_instrument_kind() -> None:
    mod = _load_client_module(with_error_time=True)
    client = mod.IBKRClient()

    stk = _fake_contract(265598, "AAPL")
    fx = _fake_contract(12087792, "EUR")
    fx.secType = "CASH"

    client.position("DU1", stk, Decimal("10"), 305.17)
    client.position("DU1", fx, Decimal("40000"), 1.15)
    client.updatePortfolio(stk, Decimal("10"), 305.0, 3050.0, 305.17, -1.7, 0.0, "DU1")

    snap = client._snapshot()
    kinds = {r["symbol"]: r["instrument_kind"] for r in snap["positions"]}
    assert kinds == {"AAPL": "Stock", "EUR": "FX Cash"}
    portfolio_kinds = {r["symbol"]: r["instrument_kind"] for r in snap["portfolio"]}
    assert portfolio_kinds == {"AAPL": "Stock"}


def test_error_row_carries_local_timestamp() -> None:
    mod = _load_client_module(with_error_time=True)
    client = mod.IBKRClient()

    client.error(-1, 1785785588057, 2104, "Market data farm connection is OK")

    row = client._snapshot()["errors"][0]
    assert row["error_time"] == 1785785588057
    # Local timestamp should be an ISO-8601 string, non-empty, and start with a year.
    assert row["error_time_local"]
    assert row["error_time_local"][:4].isdigit()


def test_snapshot_is_stably_sorted() -> None:
    """Two identical insertion orders reversed should still emit the same
    output ordering after the sort in _snapshot()."""
    mod = _load_client_module(with_error_time=True)
    a = mod.IBKRClient()
    b = mod.IBKRClient()

    for client, order in ((a, ("MSFT", "AAPL", "GOOG")), (b, ("GOOG", "AAPL", "MSFT"))):
        for i, sym in enumerate(order):
            c = _fake_contract(i + 1, sym)
            client.position("DU1", c, Decimal("1"), 1.0)

    syms_a = [r["symbol"] for r in a._snapshot()["positions"]]
    syms_b = [r["symbol"] for r in b._snapshot()["positions"]]
    assert syms_a == syms_b == sorted(syms_a)


def _run_all() -> int:
    tests = [
        test_portfolio_upserts_on_repeated_updates,
        test_portfolio_distinct_accounts_and_contracts_kept_separate,
        test_update_account_value_uses_ledger_prefix,
        test_error_new_signature_parses_code_and_message,
        test_error_legacy_signature_parses_code_and_message,
        test_error_classifies_client_errors_as_error_kind,
        test_snapshot_shape_is_stable_empty,
        test_account_summary_dedupes_dual_dispatch,
        test_positions_dedupe_dual_dispatch,
        test_account_values_dedupe_dual_dispatch,
        test_account_summary_carries_numeric_and_text_value,
        test_position_and_portfolio_carry_instrument_kind,
        test_error_row_carries_local_timestamp,
        test_snapshot_is_stably_sorted,
    ]
    failures = 0
    for test in tests:
        try:
            test()
        except AssertionError as failure:
            failures += 1
            print(f"FAIL  {test.__name__}: {failure}")
        except Exception as failure:  # noqa: BLE001
            failures += 1
            print(f"ERROR {test.__name__}: {failure!r}")
        else:
            print(f"PASS  {test.__name__}")
    print(f"\n{len(tests) - failures}/{len(tests)} tests passed")
    return 0 if failures == 0 else 1


if __name__ == "__main__":
    raise SystemExit(_run_all())

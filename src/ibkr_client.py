"""Read-only IBKR TWS API collector.

Retrieves managed accounts, account summary, per-currency cash/values,
positions, and portfolio snapshots. Does not implement any order-placing
method. Compatible with ibapi >= 10.19 (new ``error()`` signature).
"""

from __future__ import annotations

import inspect
import logging
import threading
from dataclasses import asdict, dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Any

from ibapi.client import EClient
from ibapi.contract import Contract
from ibapi.wrapper import EWrapper

LOGGER = logging.getLogger(__name__)


# --- Row types -------------------------------------------------------------


def _try_float(value: Any) -> float | None:
    """Coerce a stringy IBKR value to float if possible, else None.

    IBKR account-summary/ledger values arrive as strings. We keep the raw
    string in the row and expose the numeric version alongside so Excel can
    sort/sum without lexicographic surprises.
    """
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        return float(Decimal(text.replace(",", "")))
    except (InvalidOperation, ValueError, TypeError):
        return None


# secType -> user-facing kind used in the Positions sheet.
_SEC_TYPE_KIND: dict[str, str] = {
    "STK": "Stock",
    "CASH": "FX Cash",
    "OPT": "Option",
    "FUT": "Future",
    "FOP": "Future Option",
    "IND": "Index",
    "BAG": "Combo",
    "BOND": "Bond",
    "WAR": "Warrant",
    "FUND": "Fund",
    "CRYPTO": "Crypto",
    "CFD": "CFD",
}


def _instrument_kind(sec_type: str) -> str:
    return _SEC_TYPE_KIND.get((sec_type or "").upper(), sec_type or "Unknown")


@dataclass
class AccountSummaryRow:
    request_id: int
    account: str
    tag: str
    value: str
    value_numeric: float | None
    currency: str


@dataclass
class AccountValueRow:
    """Raw $LEDGER-* per-currency account values (one metric per row).

    Populated for every ``$LEDGER-*`` key emitted by IBKR when the currency
    is not ``BASE``. The ``metric`` field is the key with the ``$LEDGER-``
    prefix stripped for readability.
    """

    account: str
    metric: str
    currency: str
    value: str
    value_numeric: float | None


@dataclass
class PositionRow:
    account: str
    symbol: str
    security_type: str
    instrument_kind: str
    currency: str
    exchange: str
    primary_exchange: str
    contract_id: int
    quantity: float
    average_cost: float


@dataclass
class PortfolioRow:
    account: str
    symbol: str
    security_type: str
    instrument_kind: str
    currency: str
    exchange: str
    contract_id: int
    quantity: float
    market_price: float
    market_value: float
    average_cost: float
    unrealized_pnl: float
    realized_pnl: float


@dataclass
class ErrorRow:
    request_id: int
    code: int
    message: str
    advanced_rejection: str = ""
    error_time: int | None = None
    error_time_local: str = ""  # human-readable ISO timestamp
    kind: str = "warning"  # "info", "warning", "error"


# --- Constants -------------------------------------------------------------


# Codes IBKR emits as *status* messages, not real errors.
# See TWS API "Message Codes" reference.
INFORMATIONAL_CODES: frozenset[int] = frozenset(
    {
        2100,  # New account data requested
        2103,  # Market data farm connection is broken
        2104,  # Market data farm connection is OK
        2105,  # HMDS data farm connection is broken
        2106,  # HMDS data farm connection is OK
        2107,  # HMDS data farm connection is inactive but should be available
        2108,  # Market data farm connection is inactive but should be available
        2119,  # Market data farm is connecting
        2137,  # Cross Side Warning (subscription reset)
        2148,  # HMDS server connection was closed
        2158,  # Sec-def data farm connection is OK
        2168,  # Cross Side Warning
        2169,  # Cross Side Warning
    }
)


# Account summary tags requested from IBKR (see TWS Account Summary docs).
ACCOUNT_SUMMARY_TAGS: str = ",".join(
    [
        "AccountType",
        "NetLiquidation",
        "TotalCashValue",
        "SettledCash",
        "BuyingPower",
        "AvailableFunds",
        "ExcessLiquidity",
        "GrossPositionValue",
        "UnrealizedPnL",
        "RealizedPnL",
        "AccruedCash",
        "Cushion",
        "EquityWithLoanValue",
        "FullInitMarginReq",
        "FullMaintMarginReq",
    ]
)


# Per-account-value keys we surface in the "Cash By Currency" sheet.
#
# IBKR emits per-currency ledger data under the ``$LEDGER-`` prefix (as
# discovered by the diagnostic dump). The plain ``CashBalance`` key is only
# emitted with segment suffixes (``-C``/``-S``/``-P``) in the base currency,
# so filtering on that key alone yields nothing for multi-currency accounts.
LEDGER_PREFIX: str = "$LEDGER-"
# Metrics (post-prefix names) we forward to the Cash By Currency sheet.
LEDGER_METRICS: frozenset[str] = frozenset(
    {
        "CashBalance",
        "TotalCashBalance",
        "NetLiquidationByCurrency",
        "ExchangeRate",
        "AccruedCash",
        "RealizedPnL",
        "UnrealizedPnL",
        "StockMarketValue",
        "FxCashBalance",
    }
)


# Reserve a namespace of request IDs so future features (historical bars,
# fundamentals) don't collide with the ones used here.
REQ_ACCOUNT_SUMMARY = 9001


# --- Client ----------------------------------------------------------------


class IBKRClient(EWrapper, EClient):
    """Threaded EClient/EWrapper that snapshots a read-only account view."""

    def __init__(self) -> None:
        EClient.__init__(self, self)

        self.connected_event = threading.Event()
        self.managed_accounts_event = threading.Event()
        self.account_summary_end_event = threading.Event()
        self.positions_end_event = threading.Event()
        # One event per account_download_end; keyed by account name.
        self._account_download_events: dict[str, threading.Event] = {}
        # Mutex protects mutable collections shared with the API thread.
        self._lock = threading.RLock()

        self.managed_account_ids: list[str] = []
        # Every dict below is keyed to make handlers idempotent. ibapi >= 10.19
        # dual-dispatches many callbacks (legacy + *ProtoBuf) so plain
        # list.append() causes double rows. See analysis §2.1 (portfolio) and
        # the runtime finding for accountSummary / updateAccountValue.
        self._account_summary: dict[tuple[str, str, str], AccountSummaryRow] = {}
        # Keyed on (account, metric, currency) — one row per ledger metric.
        self._account_values: dict[tuple[str, str, str], AccountValueRow] = {}
        self._positions: dict[tuple[str, int], PositionRow] = {}
        self._portfolio: dict[tuple[str, int], PortfolioRow] = {}
        self.errors: list[ErrorRow] = []

        self._api_thread: threading.Thread | None = None

        # ibapi >= 10.19 added ``errorTime`` to ``EWrapper.error``. Detect once
        # so we tolerate either signature at import time.
        self._error_has_time: bool = "errorTime" in inspect.signature(EWrapper.error).parameters

    # -- helpers --

    @staticmethod
    def _to_float(value: Any) -> float:
        """Coerce ibapi ``Decimal``/``str``/``float`` payloads to ``float``."""
        if value is None:
            return 0.0
        try:
            return float(Decimal(str(value)))
        except (InvalidOperation, ValueError, TypeError):
            return 0.0

    def _reset_events(self) -> None:
        """Clear all events so a client instance can be reused."""
        self.connected_event.clear()
        self.managed_accounts_event.clear()
        self.account_summary_end_event.clear()
        self.positions_end_event.clear()
        with self._lock:
            for event in self._account_download_events.values():
                event.clear()

    # -- lifecycle --

    def connect_and_start(
        self,
        host: str,
        port: int,
        client_id: int,
        timeout: int,
    ) -> None:
        LOGGER.info("Connecting to IBKR at %s:%s with client ID %s", host, port, client_id)

        self._reset_events()
        self.connect(host, port, clientId=client_id)

        self._api_thread = threading.Thread(
            target=self.run,
            name="IBKRMessageLoop",
            daemon=True,
        )
        self._api_thread.start()

        if not self.connected_event.wait(timeout):
            self.disconnect()
            raise TimeoutError(
                f"IBKR connection timed out after {timeout} seconds. "
                "Confirm that TWS is running and API socket access is enabled."
            )

    def stop(self) -> None:
        try:
            if self.isConnected():
                self.disconnect()
        finally:
            if self._api_thread and self._api_thread.is_alive():
                self._api_thread.join(timeout=3)

    # -- collection --

    def collect(self, timeout: int) -> dict[str, list[dict[str, Any]]]:
        """Run the read-only snapshot sequence and return dict-of-rows."""

        self.reqManagedAccts()
        if not self.managed_accounts_event.wait(timeout):
            raise TimeoutError("Managed account request timed out.")
        if not self.managed_account_ids:
            raise RuntimeError("IBKR returned no managed account IDs.")

        LOGGER.info("Managed accounts: %s", self.managed_account_ids)
        if len(self.managed_account_ids) > 1:
            LOGGER.warning(
                "Multiple managed accounts detected (%s). Portfolio download "
                "will iterate through all of them.",
                len(self.managed_account_ids),
            )

        self.reqAccountSummary(REQ_ACCOUNT_SUMMARY, "All", ACCOUNT_SUMMARY_TAGS)
        try:
            if not self.account_summary_end_event.wait(timeout):
                raise TimeoutError("Account summary request timed out.")
        finally:
            # Cancel the subscription even on timeout so a subsequent run
            # (or ``stop()``) does not leave a live subscription on the wire.
            self._safe_cancel(self.cancelAccountSummary, REQ_ACCOUNT_SUMMARY)

        self.reqPositions()
        try:
            if not self.positions_end_event.wait(timeout):
                raise TimeoutError("Positions request timed out.")
        finally:
            self._safe_cancel(self.cancelPositions)

        for account in self.managed_account_ids:
            self._download_portfolio(account, timeout)

        return self._snapshot()

    @staticmethod
    def _safe_cancel(fn: Any, *args: Any) -> None:
        """Call an ibapi cancel* method; swallow errors so cleanup never masks the real one."""
        try:
            fn(*args)
        except Exception:  # noqa: BLE001 - best-effort cleanup on timeout paths
            LOGGER.debug("Ignored error during cleanup call %s", getattr(fn, "__name__", fn))

    def _download_portfolio(self, account: str, timeout: int) -> None:
        with self._lock:
            event = self._account_download_events.setdefault(account, threading.Event())
            event.clear()

        self.reqAccountUpdates(True, account)
        try:
            if not event.wait(timeout):
                LOGGER.warning(
                    "Portfolio download for %s did not signal completion within "
                    "%s seconds; results may be partial.",
                    account,
                    timeout,
                )
        finally:
            # Cancel the streaming subscription so subsequent accounts can proceed
            # even if we timed out or raised above.
            self._safe_cancel(self.reqAccountUpdates, False, account)

    def _snapshot(self) -> dict[str, list[dict[str, Any]]]:
        with self._lock:
            # Stable, deterministic ordering across runs. Users expect the
            # same row order on repeated snapshots so diffs stay meaningful.
            summary_rows = sorted(
                self._account_summary.values(),
                key=lambda r: (r.account, r.tag, r.currency or ""),
            )
            value_rows = sorted(
                self._account_values.values(),
                key=lambda r: (r.account, r.currency, r.metric),
            )
            position_rows = sorted(
                self._positions.values(),
                key=lambda r: (r.account, r.security_type, r.symbol),
            )
            portfolio_rows = sorted(
                self._portfolio.values(),
                key=lambda r: (r.account, r.security_type, r.symbol),
            )
            return {
                "accounts": [{"account": a} for a in self.managed_account_ids],
                "account_summary": [asdict(r) for r in summary_rows],
                "account_values": [asdict(r) for r in value_rows],
                "positions": [asdict(r) for r in position_rows],
                "portfolio": [asdict(r) for r in portfolio_rows],
                "errors": [asdict(r) for r in self.errors],
            }

    # -- EWrapper callbacks --

    def nextValidId(self, orderId: int) -> None:  # noqa: N802 (ibapi naming)
        LOGGER.info("API connection established. Next valid order ID: %s", orderId)
        self.connected_event.set()

    def managedAccounts(self, accountsList: str) -> None:  # noqa: N802,N803
        with self._lock:
            self.managed_account_ids = [
                account.strip() for account in accountsList.split(",") if account.strip()
            ]
        self.managed_accounts_event.set()

    def accountSummary(  # noqa: N802
        self,
        reqId: int,
        account: str,
        tag: str,
        value: str,
        currency: str,
    ) -> None:
        # Key on (account, tag, currency) so ibapi 10.19+ dual-dispatch
        # (legacy + accountSummaryProtoBuf) upserts instead of duplicating.
        with self._lock:
            self._account_summary[(account, tag, currency or "")] = AccountSummaryRow(
                request_id=reqId,
                account=account,
                tag=tag,
                value=value,
                value_numeric=_try_float(value),
                currency=currency,
            )

    def accountSummaryEnd(self, reqId: int) -> None:  # noqa: N802
        LOGGER.info("Account summary completed for request %s", reqId)
        self.account_summary_end_event.set()

    def updateAccountValue(  # noqa: N802
        self,
        key: str,
        val: str,
        currency: str,
        accountName: str,  # noqa: N803
    ) -> None:
        # IBKR routes per-currency ledger data through keys of the form
        # ``$LEDGER-<metric>`` with a real ISO currency. Plain ``CashBalance``
        # is only emitted with segment suffixes in the base currency.
        if not key.startswith(LEDGER_PREFIX):
            return
        metric = key[len(LEDGER_PREFIX) :]
        if metric not in LEDGER_METRICS:
            return
        if not currency or currency == "BASE":
            return
        with self._lock:
            self._account_values[(accountName, metric, currency)] = AccountValueRow(
                account=accountName,
                metric=metric,
                currency=currency,
                value=val,
                value_numeric=_try_float(val),
            )

    def position(
        self,
        account: str,
        contract: Contract,
        position: Decimal,
        avgCost: float,  # noqa: N803
    ) -> None:
        # Key on (account, conId) so ibapi 10.19+ dual-dispatch upserts.
        with self._lock:
            self._positions[(account, contract.conId)] = PositionRow(
                account=account,
                symbol=contract.symbol,
                security_type=contract.secType,
                instrument_kind=_instrument_kind(contract.secType),
                currency=contract.currency,
                exchange=contract.exchange or "",
                primary_exchange=contract.primaryExchange or "",
                contract_id=contract.conId,
                quantity=self._to_float(position),
                average_cost=self._to_float(avgCost),
            )

    def positionEnd(self) -> None:  # noqa: N802
        LOGGER.info("Positions collection completed")
        self.positions_end_event.set()

    def updatePortfolio(  # noqa: N802
        self,
        contract: Contract,
        position: Decimal,
        marketPrice: float,  # noqa: N803
        marketValue: float,  # noqa: N803
        averageCost: float,  # noqa: N803
        unrealizedPNL: float,  # noqa: N803
        realizedPNL: float,  # noqa: N803
        accountName: str,  # noqa: N803
    ) -> None:
        row = PortfolioRow(
            account=accountName,
            symbol=contract.symbol,
            security_type=contract.secType,
            instrument_kind=_instrument_kind(contract.secType),
            currency=contract.currency,
            exchange=contract.exchange or "",
            contract_id=contract.conId,
            quantity=self._to_float(position),
            market_price=self._to_float(marketPrice),
            market_value=self._to_float(marketValue),
            average_cost=self._to_float(averageCost),
            unrealized_pnl=self._to_float(unrealizedPNL),
            realized_pnl=self._to_float(realizedPNL),
        )
        with self._lock:
            self._portfolio[(accountName, contract.conId)] = row

    def accountDownloadEnd(self, accountName: str) -> None:  # noqa: N802,N803
        LOGGER.info("Portfolio download completed for %s", accountName)
        with self._lock:
            event = self._account_download_events.setdefault(accountName, threading.Event())
        event.set()

    # ibapi < 10.19 signature: (reqId, errorCode, errorString, advancedOrderRejectJson)
    # ibapi >= 10.19 signature: (reqId, errorTime, errorCode, errorString, advancedOrderRejectJson)
    # We accept both by matching positional args at runtime.
    def error(  # noqa: N802
        self,
        reqId: int,
        *args: Any,
        **kwargs: Any,
    ) -> None:
        error_time: int | None = None
        advanced: str = ""

        if self._error_has_time:
            if len(args) >= 3:
                error_time, error_code, error_string = args[0], args[1], args[2]
                if len(args) >= 4:
                    advanced = args[3]
            else:
                error_time = kwargs.get("errorTime")
                error_code = kwargs.get("errorCode")
                error_string = kwargs.get("errorString", "")
                advanced = kwargs.get("advancedOrderRejectJson", "")
        else:
            if len(args) >= 2:
                error_code, error_string = args[0], args[1]
                if len(args) >= 3:
                    advanced = args[2]
            else:
                error_code = kwargs.get("errorCode")
                error_string = kwargs.get("errorString", "")
                advanced = kwargs.get("advancedOrderRejectJson", "")

        try:
            code = int(error_code)
        except (TypeError, ValueError):
            code = -1

        kind = "info" if code in INFORMATIONAL_CODES else "warning"
        # Real client errors (bad request, bad contract, etc.) live in 100-499.
        if 100 <= code < 500:
            kind = "error"

        # Derive a human-readable timestamp from IBKR's epoch-millis time.
        # ibapi returns 0 for legacy code paths and a full ms epoch otherwise.
        error_time_local = ""
        if isinstance(error_time, int) and error_time > 0:
            try:
                error_time_local = (
                    datetime.fromtimestamp(error_time / 1000.0)
                    .astimezone()
                    .isoformat(timespec="milliseconds")
                )
            except (OverflowError, OSError, ValueError):
                error_time_local = ""

        entry = ErrorRow(
            request_id=reqId,
            code=code,
            message=str(error_string) if error_string is not None else "",
            advanced_rejection=str(advanced) if advanced else "",
            error_time=error_time,
            error_time_local=error_time_local,
            kind=kind,
        )
        with self._lock:
            self.errors.append(entry)

        if kind == "info":
            LOGGER.info("IBKR status %s: %s", code, entry.message)
        elif kind == "error":
            LOGGER.error("IBKR error %s: %s", code, entry.message)
        else:
            LOGGER.warning("IBKR warning %s: %s", code, entry.message)


__all__ = [
    "IBKRClient",
    "AccountSummaryRow",
    "AccountValueRow",
    "PositionRow",
    "PortfolioRow",
    "ErrorRow",
    "INFORMATIONAL_CODES",
    "ACCOUNT_SUMMARY_TAGS",
]

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
from decimal import Decimal, InvalidOperation
from typing import Any

from ibapi.client import EClient
from ibapi.contract import Contract
from ibapi.wrapper import EWrapper


LOGGER = logging.getLogger(__name__)


# --- Row types -------------------------------------------------------------


@dataclass
class AccountSummaryRow:
    request_id: int
    account: str
    tag: str
    value: str
    currency: str


@dataclass
class AccountValueRow:
    account: str
    key: str
    value: str
    currency: str


@dataclass
class PositionRow:
    account: str
    symbol: str
    security_type: str
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
CASH_VALUE_KEYS: frozenset[str] = frozenset(
    {"CashBalance", "TotalCashBalance", "RealizedPnL", "UnrealizedPnL"}
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
        self._account_summary: dict[
            tuple[str, str, str], AccountSummaryRow
        ] = {}
        self._account_values: dict[
            tuple[str, str, str], AccountValueRow
        ] = {}
        self._positions: dict[tuple[str, int], PositionRow] = {}
        self._portfolio: dict[tuple[str, int], PortfolioRow] = {}
        self.errors: list[ErrorRow] = []

        self._api_thread: threading.Thread | None = None

        # ibapi >= 10.19 added ``errorTime`` to ``EWrapper.error``. Detect once
        # so we tolerate either signature at import time.
        self._error_has_time: bool = (
            "errorTime" in inspect.signature(EWrapper.error).parameters
        )

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
        LOGGER.info(
            "Connecting to IBKR at %s:%s with client ID %s", host, port, client_id
        )

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
        if not self.account_summary_end_event.wait(timeout):
            raise TimeoutError("Account summary request timed out.")
        self.cancelAccountSummary(REQ_ACCOUNT_SUMMARY)

        self.reqPositions()
        if not self.positions_end_event.wait(timeout):
            raise TimeoutError("Positions request timed out.")
        self.cancelPositions()

        for account in self.managed_account_ids:
            self._download_portfolio(account, timeout)

        return self._snapshot()

    def _download_portfolio(self, account: str, timeout: int) -> None:
        with self._lock:
            event = self._account_download_events.setdefault(
                account, threading.Event()
            )
            event.clear()

        self.reqAccountUpdates(True, account)
        if not event.wait(timeout):
            LOGGER.warning(
                "Portfolio download for %s did not signal completion within "
                "%s seconds; results may be partial.",
                account,
                timeout,
            )
        # Cancel the streaming subscription so subsequent accounts can proceed.
        self.reqAccountUpdates(False, account)

    def _snapshot(self) -> dict[str, list[dict[str, Any]]]:
        with self._lock:
            return {
                "accounts": [{"account": a} for a in self.managed_account_ids],
                "account_summary": [
                    asdict(r) for r in self._account_summary.values()
                ],
                "account_values": [
                    asdict(r) for r in self._account_values.values()
                ],
                "positions": [asdict(r) for r in self._positions.values()],
                "portfolio": [asdict(r) for r in self._portfolio.values()],
                "errors": [asdict(r) for r in self.errors],
            }

    # -- EWrapper callbacks --

    def nextValidId(self, orderId: int) -> None:  # noqa: N802 (ibapi naming)
        LOGGER.info("API connection established. Next valid order ID: %s", orderId)
        self.connected_event.set()

    def managedAccounts(self, accountsList: str) -> None:  # noqa: N802,N803
        with self._lock:
            self.managed_account_ids = [
                account.strip()
                for account in accountsList.split(",")
                if account.strip()
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
            self._account_summary[(account, tag, currency or "")] = (
                AccountSummaryRow(
                    request_id=reqId,
                    account=account,
                    tag=tag,
                    value=value,
                    currency=currency,
                )
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
        # Only surface currency-bearing cash/PnL rows; skip base-currency
        # aggregates that duplicate the Account Summary sheet.
        if key not in CASH_VALUE_KEYS:
            return
        if not currency or currency in {"", "BASE"}:
            return
        with self._lock:
            self._account_values[(accountName, key, currency)] = AccountValueRow(
                account=accountName,
                key=key,
                value=val,
                currency=currency,
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
            event = self._account_download_events.setdefault(
                accountName, threading.Event()
            )
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

        entry = ErrorRow(
            request_id=reqId,
            code=code,
            message=str(error_string) if error_string is not None else "",
            advanced_rejection=str(advanced) if advanced else "",
            error_time=error_time,
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

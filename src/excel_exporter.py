"""Excel workbook exporter for the IBKR read-only snapshot.

Produces machine-owned ``IBKR_*`` tabs. In ``tax_workbook`` mode, preserves
all foreign Lei 14.754 sheets and adds ``IBKR_Reconciliacao``.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

from reconcile import reconcile_quantities


# Columns rendered with the "money" format (2 decimals, thousands separator).
_MONEY_COLUMNS: frozenset[str] = frozenset(
    {
        "average_cost",
        "market_price",
        "market_value",
        "unrealized_pnl",
        "realized_pnl",
        "value_numeric",
    }
)

# Columns rendered as quantities (up to 4 decimals to accommodate fractional
# shares and FX lots).
_QUANTITY_COLUMNS: frozenset[str] = frozenset(
    {"quantity", "qty_ibkr", "qty_fiscal", "delta"}
)

_MONEY_FORMAT: str = "#,##0.00;[Red]-#,##0.00"
_QUANTITY_FORMAT: str = "#,##0.####"

OWNED_SHEETS_STANDALONE: tuple[str, ...] = (
    "IBKR_Overview",
    "IBKR_Account_Summary",
    "IBKR_Cash_By_Currency",
    "IBKR_Positions",
    "IBKR_Portfolio",
    "IBKR_API_Messages",
)

OWNED_SHEETS_TAX: tuple[str, ...] = OWNED_SHEETS_STANDALONE + (
    "IBKR_Reconciliacao",
)

LEGACY_OWNED_SHEETS: tuple[str, ...] = (
    "Overview",
    "Account Summary",
    "Cash By Currency",
    "Positions",
    "Portfolio",
    "API Messages",
)


def extract_fiscal_quantities(workbook_path: Path) -> list[dict[str, Any]]:
    """Read fiscal quantities from MyProfit_2026 or Posicoes_Atuais."""
    wb = load_workbook(workbook_path, data_only=True, read_only=True)
    try:
        if "MyProfit_2026" in wb.sheetnames:
            ws = wb["MyProfit_2026"]
            rows: list[dict[str, Any]] = []
            for row in ws.iter_rows(min_row=5, values_only=True):
                if not row or len(row) < 3:
                    continue
                symbol, qty = row[1], row[2]
                if symbol is None or str(symbol).strip() == "":
                    continue
                rows.append({"symbol": symbol, "quantity": qty})
            return rows

        if "Posicoes_Atuais" in wb.sheetnames:
            ws = wb["Posicoes_Atuais"]
            rows = []
            for row in ws.iter_rows(min_row=5, values_only=True):
                if not row or len(row) < 4:
                    continue
                symbol, qty = row[0], row[3]
                if symbol is None or str(symbol).strip() == "":
                    continue
                rows.append({"symbol": symbol, "quantity": qty})
            return rows

        return []
    finally:
        wb.close()


class ExcelExporter:
    def __init__(
        self,
        output_path: Path,
        *,
        mode: str = "standalone",
        qty_tolerance: float = 0.0001,
    ) -> None:
        if mode not in {"standalone", "tax_workbook"}:
            raise ValueError(f"Unknown exporter mode: {mode!r}")
        self.output_path = Path(output_path)
        self.mode = mode
        self.qty_tolerance = float(qty_tolerance)
        self.owned_sheets: tuple[str, ...] = (
            OWNED_SHEETS_TAX if mode == "tax_workbook" else OWNED_SHEETS_STANDALONE
        )

    def _assert_not_locked(self) -> None:
        lock = self.output_path.parent / f"~${self.output_path.name}"
        if lock.exists():
            raise PermissionError(
                f"Workbook appears open in Excel ({lock.name}). "
                "Close Excel and retry."
            )

    def _load_or_create_workbook(self) -> Workbook:
        """Open an existing workbook (preserving foreign sheets) or create
        a fresh one in standalone mode.

        Sheet order in the returned workbook is:
        - all foreign sheets first (in their original order)
        - then the owned sheets in canonical order, appended by ``export()``.
        """
        overview_name = "IBKR_Overview"

        if self.mode == "tax_workbook":
            if not self.output_path.exists():
                raise FileNotFoundError(
                    f"Tax workbook not found: {self.output_path}. "
                    "Copy the reconciled Lei 14.754 file to this path first."
                )
            self._assert_not_locked()
            try:
                workbook = load_workbook(self.output_path)
            except Exception as exc:  # noqa: BLE001
                raise ValueError(
                    f"Tax workbook is unreadable or corrupt: {self.output_path}"
                ) from exc
        elif not self.output_path.exists():
            workbook = Workbook()
            active = workbook.active
            if active is not None and active.title in {"Sheet", "Sheet1"}:
                workbook.remove(active)
            workbook.create_sheet(overview_name)
            return workbook
        else:
            self._assert_not_locked()
            try:
                workbook = load_workbook(self.output_path)
            except Exception:  # noqa: BLE001 - corrupt file: fall back to fresh
                workbook = Workbook()
                default = workbook.active
                if default is not None:
                    workbook.remove(default)
                workbook.create_sheet(overview_name)
                return workbook

        delete_names = set(self.owned_sheets) | set(LEGACY_OWNED_SHEETS)
        for name in list(workbook.sheetnames):
            if name in delete_names:
                del workbook[name]

        workbook.create_sheet(overview_name)
        return workbook

    # -- helpers --

    @staticmethod
    def _humanize(column_name: str) -> str:
        exact = {
            "value_numeric": "Value (Numeric)",
            "error_time_local": "Error Time (Local)",
            "error_time": "Error Time (ms)",
            "instrument_kind": "Kind",
            "security_type": "SecType",
            "primary_exchange": "Primary Exchange",
            "qty_ibkr": "Qty IBKR",
            "qty_fiscal": "Qty Fiscal",
        }
        if column_name in exact:
            return exact[column_name]
        replacements = {
            "Pnl": "P&L",
            "Id": "ID",
            "Api": "API",
        }
        title = column_name.replace("_", " ").title()
        for source, target in replacements.items():
            title = title.replace(source, target)
        return title

    @staticmethod
    def _autosize(worksheet: Worksheet, cap: int = 40) -> None:
        for column_cells in worksheet.columns:
            first = column_cells[0]
            if first.column_letter is None:
                continue
            column_letter = get_column_letter(first.column)
            maximum_length = 0
            for cell in column_cells:
                value = "" if cell.value is None else str(cell.value)
                if len(value) > maximum_length:
                    maximum_length = len(value)
            worksheet.column_dimensions[column_letter].width = min(
                maximum_length + 3, cap
            )

    @staticmethod
    def _apply_number_format(cell: Any, column: str) -> None:
        if column in _MONEY_COLUMNS:
            cell.number_format = _MONEY_FORMAT
        elif column in _QUANTITY_COLUMNS:
            cell.number_format = _QUANTITY_FORMAT

    @staticmethod
    def _coerce_number(column: str, value: Any) -> Any:
        if column not in _MONEY_COLUMNS and column not in _QUANTITY_COLUMNS:
            return value
        if isinstance(value, (int, float)):
            return value
        if value is None or value == "":
            return None
        try:
            return float(str(value).replace(",", ""))
        except (TypeError, ValueError):
            return value

    @staticmethod
    def _make_table(
        worksheet: Worksheet,
        table_name: str,
        row_count: int,
        column_count: int,
    ) -> None:
        if row_count < 2 or column_count < 1:
            return
        end_column = get_column_letter(column_count)
        table = Table(
            displayName=table_name,
            ref=f"A1:{end_column}{row_count}",
        )
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        worksheet.add_table(table)

    def _write_records(
        self,
        workbook: Workbook,
        sheet_name: str,
        records: list[dict[str, Any]],
        table_name: str,
        preferred_order: Iterable[str] | None = None,
    ) -> None:
        worksheet = workbook.create_sheet(sheet_name)

        if not records:
            worksheet["A1"] = f"No {sheet_name.lower()} data returned."
            worksheet["A1"].font = Font(bold=True, italic=True)
            worksheet.column_dimensions["A"].width = 40
            return

        seen_keys: set[str] = set()
        columns: list[str] = []
        if preferred_order:
            for name in preferred_order:
                if name in records[0]:
                    columns.append(name)
                    seen_keys.add(name)
        for name in records[0].keys():
            if name not in seen_keys:
                columns.append(name)

        header_fill = PatternFill(
            start_color="FF1F4E78", end_color="FF1F4E78", fill_type="solid"
        )
        header_font = Font(bold=True, color="FFFFFFFF")

        for index, column in enumerate(columns, start=1):
            cell = worksheet.cell(row=1, column=index, value=self._humanize(column))
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for row_index, record in enumerate(records, start=2):
            for column_index, column in enumerate(columns, start=1):
                value = self._coerce_number(column, record.get(column))
                cell = worksheet.cell(row=row_index, column=column_index, value=value)
                self._apply_number_format(cell, column)

        worksheet.freeze_panes = "A2"
        self._make_table(
            worksheet=worksheet,
            table_name=table_name,
            row_count=worksheet.max_row,
            column_count=worksheet.max_column,
        )
        self._autosize(worksheet)

    def _write_overview(
        self,
        workbook: Workbook,
        data: dict[str, list[dict[str, Any]]],
    ) -> None:
        overview = workbook["IBKR_Overview"]
        overview.delete_rows(1, overview.max_row)

        generated_at = datetime.now().astimezone()
        errors = data.get("errors", [])
        error_kinds = {"info": 0, "warning": 0, "error": 0}
        for row in errors:
            error_kinds[row.get("kind", "warning")] = (
                error_kinds.get(row.get("kind", "warning"), 0) + 1
            )

        accounts = [
            str(row.get("account", "")).strip()
            for row in data.get("accounts", [])
            if row.get("account")
        ]
        account_label = ", ".join(accounts) if accounts else "(none)"

        overview_rows: list[tuple[str, Any]] = [
            ("IBKR Portfolio Workbook", ""),
            ("Generated", generated_at.isoformat(timespec="seconds")),
            ("Exporter mode", self.mode),
            ("Connection mode", "Read-only"),
            ("Accounts", account_label),
            ("Account summary rows", len(data.get("account_summary", []))),
            ("Cash / P&L per currency rows", len(data.get("account_values", []))),
            ("Positions", len(data.get("positions", []))),
            ("Portfolio rows", len(data.get("portfolio", []))),
            (
                "API messages",
                f"{len(errors)} "
                f"(info: {error_kinds['info']}, "
                f"warnings: {error_kinds['warning']}, "
                f"errors: {error_kinds['error']})",
            ),
        ]

        for row_number, row in enumerate(overview_rows, start=1):
            overview.cell(row=row_number, column=1, value=row[0])
            overview.cell(row=row_number, column=2, value=row[1])

        overview.merge_cells("A1:B1")
        overview["A1"].font = Font(size=16, bold=True)
        overview["A1"].alignment = Alignment(horizontal="left", vertical="center")
        overview.column_dimensions["A"].width = 32
        overview.column_dimensions["B"].width = 42

    def _write_reconciliation(
        self,
        workbook: Workbook,
        data: dict[str, list[dict[str, Any]]],
    ) -> None:
        fiscal = extract_fiscal_quantities(self.output_path)
        rows = reconcile_quantities(
            live=data.get("positions", []),
            fiscal=fiscal,
            tolerance=self.qty_tolerance,
        )
        self._write_records(
            workbook,
            "IBKR_Reconciliacao",
            rows,
            "IBKRReconciliacaoTable",
            preferred_order=(
                "symbol",
                "qty_ibkr",
                "qty_fiscal",
                "delta",
                "status",
            ),
        )

    # -- public --

    def export(self, data: dict[str, list[dict[str, Any]]]) -> Path:
        self.output_path.parent.mkdir(parents=True, exist_ok=True)

        workbook = self._load_or_create_workbook()
        self._write_overview(workbook, data)

        self._write_records(
            workbook,
            "IBKR_Account_Summary",
            data.get("account_summary", []),
            "IBKRAccountSummaryTable",
            preferred_order=(
                "account",
                "tag",
                "value",
                "value_numeric",
                "currency",
                "request_id",
            ),
        )

        self._write_records(
            workbook,
            "IBKR_Cash_By_Currency",
            data.get("account_values", []),
            "IBKRAccountValuesTable",
            preferred_order=(
                "account",
                "currency",
                "metric",
                "value_numeric",
                "value",
            ),
        )

        self._write_records(
            workbook,
            "IBKR_Positions",
            data.get("positions", []),
            "IBKRPositionsTable",
            preferred_order=(
                "account",
                "symbol",
                "instrument_kind",
                "security_type",
                "currency",
                "exchange",
                "primary_exchange",
                "contract_id",
                "quantity",
                "average_cost",
            ),
        )

        self._write_records(
            workbook,
            "IBKR_Portfolio",
            data.get("portfolio", []),
            "IBKRPortfolioTable",
            preferred_order=(
                "account",
                "symbol",
                "instrument_kind",
                "security_type",
                "currency",
                "exchange",
                "contract_id",
                "quantity",
                "market_price",
                "market_value",
                "average_cost",
                "unrealized_pnl",
                "realized_pnl",
            ),
        )

        self._write_records(
            workbook,
            "IBKR_API_Messages",
            data.get("errors", []),
            "IBKRAPIMessagesTable",
            preferred_order=(
                "kind",
                "code",
                "message",
                "error_time_local",
                "request_id",
                "error_time",
                "advanced_rejection",
            ),
        )

        if self.mode == "tax_workbook":
            self._write_reconciliation(workbook, data)

        workbook.save(self.output_path)
        return self.output_path


__all__ = [
    "ExcelExporter",
    "OWNED_SHEETS_STANDALONE",
    "OWNED_SHEETS_TAX",
    "LEGACY_OWNED_SHEETS",
    "extract_fiscal_quantities",
]

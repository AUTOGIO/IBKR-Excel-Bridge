"""Excel workbook exporter for the IBKR read-only snapshot.

Produces a workbook with tabs: Overview, Account Summary, Cash By Currency,
Positions, Portfolio, and API Messages. Uses openpyxl tables and applies
sensible number formats to currency/quantity columns.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet


# Columns rendered with the "money" format (2 decimals, thousands separator).
_MONEY_COLUMNS: frozenset[str] = frozenset(
    {
        "average_cost",
        "market_price",
        "market_value",
        "unrealized_pnl",
        "realized_pnl",
    }
)

# Columns rendered as quantities (up to 4 decimals to accommodate fractional
# shares and FX lots).
_QUANTITY_COLUMNS: frozenset[str] = frozenset({"quantity"})

_MONEY_FORMAT: str = "#,##0.00;[Red]-#,##0.00"
_QUANTITY_FORMAT: str = "#,##0.####"


class ExcelExporter:
    def __init__(self, output_path: Path) -> None:
        self.output_path = Path(output_path)

    # -- helpers --

    @staticmethod
    def _humanize(column_name: str) -> str:
        # Normalize acronyms so titles read naturally.
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
        # Force numeric columns into ``float`` even if they arrive as strings
        # (which the IBKR API often does for account-value payloads).
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

        # Determine column order: preferred columns first, then any extras.
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
        # ``auto_filter.ref`` is applied by the table below; setting it here
        # can trigger openpyxl warnings about overlapping filters.

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
        overview = workbook.active
        assert overview is not None
        overview.title = "Overview"

        generated_at = datetime.now().astimezone()
        errors = data.get("errors", [])
        error_kinds = {"info": 0, "warning": 0, "error": 0}
        for row in errors:
            error_kinds[row.get("kind", "warning")] = (
                error_kinds.get(row.get("kind", "warning"), 0) + 1
            )

        overview_rows: list[tuple[str, Any]] = [
            ("IBKR Portfolio Workbook", ""),
            ("Generated", generated_at.isoformat(timespec="seconds")),
            ("Connection mode", "Read-only (paper)"),
            ("Accounts", len(data.get("accounts", []))),
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

    # -- public --

    def export(self, data: dict[str, list[dict[str, Any]]]) -> Path:
        self.output_path.parent.mkdir(parents=True, exist_ok=True)

        workbook = Workbook()
        self._write_overview(workbook, data)

        self._write_records(
            workbook,
            "Account Summary",
            data.get("account_summary", []),
            "AccountSummaryTable",
            preferred_order=("account", "tag", "value", "currency", "request_id"),
        )

        self._write_records(
            workbook,
            "Cash By Currency",
            data.get("account_values", []),
            "AccountValuesTable",
            preferred_order=("account", "key", "currency", "value"),
        )

        self._write_records(
            workbook,
            "Positions",
            data.get("positions", []),
            "PositionsTable",
            preferred_order=(
                "account",
                "symbol",
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
            "Portfolio",
            data.get("portfolio", []),
            "PortfolioTable",
            preferred_order=(
                "account",
                "symbol",
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
            "API Messages",
            data.get("errors", []),
            "APIMessagesTable",
            preferred_order=(
                "kind",
                "code",
                "message",
                "request_id",
                "error_time",
                "advanced_rejection",
            ),
        )

        workbook.save(self.output_path)
        return self.output_path


__all__ = ["ExcelExporter"]

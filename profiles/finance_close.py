"""Finance close profile rules."""

from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import ClassVar

from core.reader import CellValue, RowData, SheetData, WorkbookData
from core.reporting import ReportIssue
from profiles.base import ProfileConfig

TARGET_SCHEMA: RowData = ("date", "account", "description", "debit", "credit", "balance")
REQUIRED_COLUMNS: tuple[str, ...] = ("date", "account", "debit", "credit", "balance")
NUMERIC_COLUMNS: tuple[str, ...] = ("debit", "credit", "balance")


class FinanceCloseProfile:
    name: ClassVar[str] = "finance_close"
    description: ClassVar[str] = "Finance close workbook consolidation"

    def validate(self, workbook: WorkbookData, config: ProfileConfig) -> tuple[ReportIssue, ...]:
        issues: list[ReportIssue] = []

        for sheet in workbook.sheets:
            sheet_issues: list[ReportIssue] = []
            header_map = _header_map(sheet)
            missing_columns = [column for column in REQUIRED_COLUMNS if column not in header_map]
            if missing_columns:
                issues.append(
                    ReportIssue(
                        code="missing_required_column",
                        message=(
                            f"Sheet '{sheet.name}' in '{workbook.path}' is missing required "
                            f"columns: {', '.join(missing_columns)}"
                        ),
                        path=str(workbook.path),
                        sheet=sheet.name,
                    )
                )
                continue

            sheet_issues.extend(_validate_row_types(workbook.path, sheet, header_map))
            issues.extend(sheet_issues)
            if not sheet_issues:
                balance_issue = _validate_balance(workbook.path, sheet, header_map, config)
                if balance_issue is not None:
                    issues.append(balance_issue)

        return tuple(issues)

    def transform(self, workbook: WorkbookData, config: ProfileConfig) -> WorkbookData:
        transformed_sheets: list[SheetData] = []

        for sheet in workbook.sheets:
            header_map = _header_map(sheet)
            rows = [TARGET_SCHEMA]
            for source_row in sheet.rows[1:]:
                rows.append(
                    (
                        _normalize_date(source_row[header_map["date"]]),
                        _string_value(source_row[header_map["account"]]),
                        _string_value(source_row[header_map["description"]])
                        if "description" in header_map
                        else "",
                        _number_value(source_row[header_map["debit"]]),
                        _number_value(source_row[header_map["credit"]]),
                        _number_value(source_row[header_map["balance"]]),
                    )
                )
            transformed_sheets.append(SheetData(name=sheet.name, rows=tuple(rows)))

        return WorkbookData(path=workbook.path, sheets=tuple(transformed_sheets))

    def postprocess(
        self, workbook: WorkbookData, config: ProfileConfig
    ) -> tuple[ReportIssue, ...]:
        return ()


def _header_map(sheet: SheetData) -> dict[str, int]:
    if not sheet.rows:
        return {}

    return {
        _normalize_header(value): index
        for index, value in enumerate(sheet.rows[0])
        if _normalize_header(value)
    }


def _normalize_header(value: CellValue) -> str:
    return str(value or "").strip().lower().replace(" ", "_").replace("-", "_")


def _validate_row_types(
    workbook_path: Path,
    sheet: SheetData,
    header_map: dict[str, int],
) -> tuple[ReportIssue, ...]:
    issues: list[ReportIssue] = []

    for row_number, row in enumerate(sheet.rows[1:], start=2):
        if not _is_date_value(_cell(row, header_map["date"])):
            issues.append(
                _type_issue(
                    workbook_path,
                    sheet.name,
                    row_number,
                    "date",
                    "a date in YYYY-MM-DD format",
                )
            )

        for column in NUMERIC_COLUMNS:
            if _number_value(_cell(row, header_map[column])) is None:
                issues.append(
                    _type_issue(workbook_path, sheet.name, row_number, column, "a number")
                )

    return tuple(issues)


def _validate_balance(
    workbook_path: Path,
    sheet: SheetData,
    header_map: dict[str, int],
    config: ProfileConfig,
) -> ReportIssue | None:
    tolerance = _decimal_option(config, "balance_tolerance", Decimal("0.01"))
    debit_total = Decimal("0")
    credit_total = Decimal("0")
    closing_balance = Decimal("0")

    for row in sheet.rows[1:]:
        debit_total += _decimal_value(_cell(row, header_map["debit"])) or Decimal("0")
        credit_total += _decimal_value(_cell(row, header_map["credit"])) or Decimal("0")
        closing_balance = _decimal_value(_cell(row, header_map["balance"])) or Decimal("0")

    expected_balance = debit_total - credit_total
    if abs(expected_balance - closing_balance) <= tolerance:
        return None

    return ReportIssue(
        code="balance_mismatch",
        message=(
            f"Sheet '{sheet.name}' in '{workbook_path}' has debit-credit total "
            f"{expected_balance} but closing balance {closing_balance}"
        ),
        path=str(workbook_path),
        sheet=sheet.name,
    )


def _type_issue(
    workbook_path: Path,
    sheet_name: str,
    row_number: int,
    column: str,
    expected: str,
) -> ReportIssue:
    return ReportIssue(
        code="invalid_column_type",
        message=(
            f"Sheet '{sheet_name}' in '{workbook_path}' row {row_number} column "
            f"'{column}' must be {expected}"
        ),
        path=str(workbook_path),
        sheet=sheet_name,
    )


def _cell(row: RowData, index: int) -> CellValue:
    return row[index] if index < len(row) else None


def _is_date_value(value: CellValue) -> bool:
    if isinstance(value, datetime | date):
        return True

    if not isinstance(value, str):
        return False

    try:
        datetime.strptime(value, "%Y-%m-%d")
    except ValueError:
        return False

    return True


def _normalize_date(value: CellValue) -> str:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return str(value)


def _number_value(value: CellValue) -> float | None:
    decimal_value = _decimal_value(value)
    if decimal_value is None:
        return None
    return float(decimal_value)


def _decimal_value(value: CellValue) -> Decimal | None:
    if value is None or isinstance(value, bool):
        return None

    try:
        return Decimal(str(value))
    except InvalidOperation:
        return None


def _decimal_option(config: ProfileConfig, key: str, default: Decimal) -> Decimal:
    value = config.options.get(key, default)
    try:
        return Decimal(str(value))
    except InvalidOperation:
        return default


def _string_value(value: CellValue) -> str:
    if value is None:
        return ""
    return str(value)

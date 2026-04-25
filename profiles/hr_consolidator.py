"""HR consolidator profile rules."""

from __future__ import annotations

import re
from datetime import date, datetime
from typing import Any, ClassVar

from core.reader import CellValue, RowData, SheetData, WorkbookData
from core.reporting import ReportIssue
from profiles.base import ProfileConfig

TARGET_SCHEMA: RowData = (
    "employee_id",
    "first_name",
    "last_name",
    "email",
    "department",
    "hire_date",
    "termination_date",
)
REQUIRED_COLUMNS: tuple[str, ...] = ("employee_id", "first_name", "last_name", "hire_date")
DATE_COLUMNS: tuple[str, ...] = ("hire_date", "termination_date")
PII_COLUMNS: frozenset[str] = frozenset({"first_name", "last_name", "email"})
DEFAULT_MASK_FIELDS: tuple[str, ...] = ("first_name", "last_name", "email")
EMPLOYEE_ID_PATTERN = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]{2,31}$")
DEFAULT_SYNONYMS: dict[str, tuple[str, ...]] = {
    "employee_id": ("employee_id", "employee", "employee_no", "employee_number", "personalnummer"),
    "first_name": ("first_name", "firstname", "given_name", "vorname"),
    "last_name": ("last_name", "lastname", "surname", "family_name", "nachname"),
    "email": ("email", "e_mail", "mail", "work_email"),
    "department": ("department", "dept", "team", "abteilung"),
    "hire_date": ("hire_date", "start_date", "entry_date", "eintrittsdatum"),
    "termination_date": ("termination_date", "end_date", "exit_date", "austrittsdatum"),
}


class HrConsolidatorProfile:
    name: ClassVar[str] = "hr_consolidator"
    description: ClassVar[str] = "HR workbook consolidation"

    def validate(self, workbook: WorkbookData, config: ProfileConfig) -> tuple[ReportIssue, ...]:
        issues: list[ReportIssue] = []

        for sheet in workbook.sheets:
            header_map = _header_map(sheet, config)
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

            for row_number, row in enumerate(sheet.rows[1:], start=2):
                employee_id = _string_value(_cell(row, header_map["employee_id"]))
                if not EMPLOYEE_ID_PATTERN.match(employee_id):
                    issues.append(
                        _validation_issue(
                            workbook,
                            sheet,
                            row_number,
                            "employee_id",
                            "must be 3-32 characters and contain only letters, numbers, '_' or '-'",
                        )
                    )

                for column in DATE_COLUMNS:
                    if column not in header_map:
                        continue
                    value = _cell(row, header_map[column])
                    if value in (None, "") and column == "termination_date":
                        continue
                    if not _is_date_value(value):
                        issues.append(
                            _validation_issue(
                                workbook,
                                sheet,
                                row_number,
                                column,
                                "must be a date in YYYY-MM-DD format",
                            )
                        )

        return tuple(issues)

    def transform(self, workbook: WorkbookData, config: ProfileConfig) -> WorkbookData:
        transformed_sheets: list[SheetData] = []
        mask_fields = _mask_fields(config)
        mask_token = _string_option(config, "mask_token", "***MASKED***")

        for sheet in workbook.sheets:
            header_map = _header_map(sheet, config)
            rows = [TARGET_SCHEMA]
            for row in sheet.rows[1:]:
                normalized_row = {
                    "employee_id": _string_value(_cell(row, header_map["employee_id"])),
                    "first_name": _string_value(_cell(row, header_map["first_name"])),
                    "last_name": _string_value(_cell(row, header_map["last_name"])),
                    "email": _string_value(_cell(row, header_map.get("email"))),
                    "department": _string_value(_cell(row, header_map.get("department"))),
                    "hire_date": _normalize_date(_cell(row, header_map["hire_date"])),
                    "termination_date": _normalize_date(
                        _cell(row, header_map.get("termination_date"))
                    ),
                }

                for field in mask_fields:
                    normalized_row[field] = mask_token

                rows.append(tuple(normalized_row[column] for column in TARGET_SCHEMA))

            transformed_sheets.append(SheetData(name=sheet.name, rows=tuple(rows)))

        return WorkbookData(path=workbook.path, sheets=tuple(transformed_sheets))

    def postprocess(
        self, workbook: WorkbookData, config: ProfileConfig
    ) -> tuple[ReportIssue, ...]:
        return ()


def _header_map(sheet: SheetData, config: ProfileConfig) -> dict[str, int]:
    if not sheet.rows:
        return {}

    synonyms = _synonyms(config)
    canonical_by_synonym = {
        synonym: canonical
        for canonical, values in synonyms.items()
        for synonym in values
    }
    header_map: dict[str, int] = {}

    for index, value in enumerate(sheet.rows[0]):
        normalized = _normalize_header(value)
        canonical = canonical_by_synonym.get(normalized)
        if canonical is not None and canonical not in header_map:
            header_map[canonical] = index

    return header_map


def _synonyms(config: ProfileConfig) -> dict[str, tuple[str, ...]]:
    synonyms = {key: tuple(values) for key, values in DEFAULT_SYNONYMS.items()}
    configured_synonyms = config.options.get("column_synonyms", {})
    if not isinstance(configured_synonyms, dict):
        return synonyms

    for canonical, values in configured_synonyms.items():
        if canonical not in synonyms:
            continue

        if isinstance(values, str):
            values = [values]
        if not isinstance(values, list):
            continue

        normalized_values = tuple(_normalize_header(value) for value in values)
        synonyms[canonical] = tuple({*synonyms[canonical], *normalized_values})

    return synonyms


def _mask_fields(config: ProfileConfig) -> tuple[str, ...]:
    configured = config.options.get("mask_fields", DEFAULT_MASK_FIELDS)
    if isinstance(configured, str):
        configured = [configured]
    if not isinstance(configured, list):
        return DEFAULT_MASK_FIELDS

    return tuple(
        field
        for field in (_normalize_header(value) for value in configured)
        if field in PII_COLUMNS
    )


def _validation_issue(
    workbook: WorkbookData,
    sheet: SheetData,
    row_number: int,
    column: str,
    message: str,
) -> ReportIssue:
    return ReportIssue(
        code="invalid_hr_record",
        message=(
            f"Sheet '{sheet.name}' in '{workbook.path}' row {row_number} "
            f"column '{column}' {message}"
        ),
        path=str(workbook.path),
        sheet=sheet.name,
    )


def _normalize_header(value: Any) -> str:
    return (
        str(value or "")
        .strip()
        .lower()
        .replace(" ", "_")
        .replace("-", "_")
        .replace(".", "")
    )


def _cell(row: RowData, index: int | None) -> CellValue:
    if index is None:
        return None
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
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return str(value).strip()


def _string_value(value: CellValue) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _string_option(config: ProfileConfig, key: str, default: str) -> str:
    value = config.options.get(key, default)
    if value is None:
        return default
    return str(value)

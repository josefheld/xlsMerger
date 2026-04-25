"""Supplier normalizer profile rules."""

from __future__ import annotations

from decimal import Decimal, InvalidOperation
from typing import Any, ClassVar

from core.reader import CellValue, RowData, SheetData, WorkbookData
from core.reporting import ReportIssue
from profiles.base import ProfileConfig

TARGET_SCHEMA: RowData = (
    "supplier",
    "invoice_id",
    "order_id",
    "invoice_date",
    "amount",
    "currency",
)
REQUIRED_COLUMNS: tuple[str, ...] = ("invoice_id", "order_id", "amount")
DEFAULT_SYNONYMS: dict[str, tuple[str, ...]] = {
    "supplier": ("supplier", "vendor", "lieferant", "supplier_name", "vendor_name"),
    "invoice_id": (
        "invoice_id",
        "invoice",
        "invoice_no",
        "invoice_number",
        "invoice_num",
        "rechnung_nr",
        "rechnungsnummer",
        "belegnummer",
    ),
    "order_id": (
        "order_id",
        "order",
        "order_no",
        "order_number",
        "purchase_order",
        "po_number",
        "po",
        "bestellnummer",
    ),
    "invoice_date": (
        "invoice_date",
        "date",
        "datum",
        "rechnungsdatum",
        "belegdatum",
    ),
    "amount": (
        "amount",
        "total",
        "total_amount",
        "gross_amount",
        "betrag",
        "summe",
    ),
    "currency": ("currency", "currency_code", "curr", "waehrung", "währung"),
}


class SupplierNormalizerProfile:
    name: ClassVar[str] = "supplier_normalizer"
    description: ClassVar[str] = "Supplier workbook normalization"

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
                if _number_value(_cell(row, header_map["amount"])) is None:
                    issues.append(
                        ReportIssue(
                            code="invalid_column_type",
                            message=(
                                f"Sheet '{sheet.name}' in '{workbook.path}' row {row_number} "
                                "column 'amount' must be a number"
                            ),
                            path=str(workbook.path),
                            sheet=sheet.name,
                        )
                    )

        return tuple(issues)

    def transform(self, workbook: WorkbookData, config: ProfileConfig) -> WorkbookData:
        transformed_sheets: list[SheetData] = []

        for sheet in workbook.sheets:
            header_map = _header_map(sheet, config)
            rows = [TARGET_SCHEMA]
            for row in sheet.rows[1:]:
                rows.append(
                    (
                        _supplier_value(row, header_map, config),
                        _string_value(_cell(row, header_map["invoice_id"])),
                        _string_value(_cell(row, header_map["order_id"])),
                        _string_value(_cell(row, header_map.get("invoice_date"))),
                        _number_value(_cell(row, header_map["amount"])),
                        _currency_value(row, header_map, config),
                    )
                )

            transformed_sheets.append(SheetData(name=sheet.name, rows=tuple(rows)))

        return WorkbookData(path=workbook.path, sheets=tuple(transformed_sheets))

    def postprocess(
        self, workbook: WorkbookData, config: ProfileConfig
    ) -> tuple[ReportIssue, ...]:
        issues: list[ReportIssue] = []

        for sheet in workbook.sheets:
            seen: dict[tuple[str, str], int] = {}
            for row_number, row in enumerate(sheet.rows[1:], start=2):
                invoice_id = _string_value(_cell(row, 1))
                order_id = _string_value(_cell(row, 2))
                key = (invoice_id, order_id)
                if key in seen:
                    issues.append(
                        ReportIssue(
                            code="duplicate_invoice_order",
                            message=(
                                f"Sheet '{sheet.name}' in '{workbook.path}' row {row_number} "
                                f"duplicates invoice_id '{invoice_id}' and order_id '{order_id}' "
                                f"from row {seen[key]}"
                            ),
                            path=str(workbook.path),
                            sheet=sheet.name,
                        )
                    )
                    continue

                seen[key] = row_number

        return tuple(issues)


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


def _supplier_value(
    row: RowData,
    header_map: dict[str, int],
    config: ProfileConfig,
) -> str:
    supplier = _string_value(_cell(row, header_map.get("supplier")))
    if not supplier:
        supplier = _string_option(config, "supplier_name")

    supplier_map = config.options.get("supplier_name_map", {})
    if not isinstance(supplier_map, dict):
        return supplier

    return str(supplier_map.get(supplier, supplier))


def _currency_value(
    row: RowData,
    header_map: dict[str, int],
    config: ProfileConfig,
) -> str:
    currency = _string_value(_cell(row, header_map.get("currency")))
    if currency:
        return currency.upper()

    return _string_option(config, "default_currency", "EUR").upper()


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


def _number_value(value: CellValue) -> float | None:
    if value is None or isinstance(value, bool):
        return None

    try:
        return float(Decimal(str(value)))
    except InvalidOperation:
        return None


def _string_value(value: CellValue) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _string_option(config: ProfileConfig, key: str, default: str = "") -> str:
    value = config.options.get(key, default)
    if value is None:
        return default
    return str(value).strip()

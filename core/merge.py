"""Merge workbook data using deterministic source ordering."""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from core.reader import RowData, SheetData, WorkbookData

HeaderStrategy = Literal["first_file", "every_file", "none"]
VALID_HEADER_STRATEGIES: frozenset[HeaderStrategy] = frozenset(
    {"first_file", "every_file", "none"}
)


class MergeError(Exception):
    """Raised when workbook data cannot be merged."""


@dataclass(frozen=True)
class SourceSheet:
    workbook_path: Path
    sheet: SheetData


def merge_workbooks(
    workbooks: Iterable[WorkbookData],
    *,
    header_strategy: HeaderStrategy = "first_file",
    output_path: str | Path = "merged.xlsx",
) -> WorkbookData:
    if header_strategy not in VALID_HEADER_STRATEGIES:
        supported = ", ".join(sorted(VALID_HEADER_STRATEGIES))
        raise MergeError(
            f"Unsupported header strategy '{header_strategy}': expected one of {supported}"
        )

    merged_rows_by_sheet: dict[str, list[RowData]] = {}
    header_seen_by_sheet: set[str] = set()

    for source in _iter_source_sheets(workbooks):
        rows = _rows_for_strategy(
            source.sheet.rows,
            sheet_name=source.sheet.name,
            header_strategy=header_strategy,
            header_seen_by_sheet=header_seen_by_sheet,
        )
        if not rows:
            continue

        merged_rows_by_sheet.setdefault(source.sheet.name, []).extend(rows)

    return WorkbookData(
        path=Path(output_path),
        sheets=tuple(
            SheetData(name=sheet_name, rows=tuple(rows))
            for sheet_name, rows in merged_rows_by_sheet.items()
        ),
    )


def _iter_source_sheets(workbooks: Iterable[WorkbookData]) -> tuple[SourceSheet, ...]:
    sources = [
        SourceSheet(workbook_path=workbook.path, sheet=sheet)
        for workbook in workbooks
        for sheet in workbook.sheets
    ]
    return tuple(
        sorted(
            sources,
            key=lambda source: (
                source.workbook_path.name.lower(),
                source.sheet.name.lower(),
                str(source.workbook_path).lower(),
            ),
        )
    )


def _rows_for_strategy(
    rows: tuple[RowData, ...],
    *,
    sheet_name: str,
    header_strategy: HeaderStrategy,
    header_seen_by_sheet: set[str],
) -> tuple[RowData, ...]:
    if not rows:
        return ()

    if header_strategy == "every_file":
        return rows

    if header_strategy == "none":
        return rows[1:]

    if sheet_name in header_seen_by_sheet:
        return rows[1:]

    header_seen_by_sheet.add(sheet_name)
    return rows

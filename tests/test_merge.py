from pathlib import Path

import pytest

from core.merge import MergeError, merge_workbooks
from core.reader import SheetData, WorkbookData


def workbook(path: str, sheet_name: str, rows: list[tuple[object, ...]]) -> WorkbookData:
    return WorkbookData(
        path=Path(path),
        sheets=(SheetData(name=sheet_name, rows=tuple(rows)),),
    )


def test_merge_processes_large_rows_and_columns_without_hard_limits() -> None:
    wide_header = tuple(f"col_{index}" for index in range(120))
    wide_row = tuple(range(120))
    rows = [wide_header, *[wide_row for _ in range(1_100)]]

    merged = merge_workbooks([workbook("large.xlsx", "Data", rows)])

    assert len(merged.sheets[0].rows) == 1_101
    assert len(merged.sheets[0].rows[0]) == 120
    assert merged.sheets[0].rows[-1][-1] == 119


def test_first_file_header_strategy_keeps_one_header_per_sheet() -> None:
    first = workbook("b.xlsx", "Data", [("name",), ("B",)])
    second = workbook("a.xlsx", "Data", [("name",), ("A",)])

    merged = merge_workbooks([first, second], header_strategy="first_file")

    assert merged.sheets[0].rows == (("name",), ("A",), ("B",))


def test_every_file_header_strategy_keeps_each_source_header() -> None:
    first = workbook("b.xlsx", "Data", [("name",), ("B",)])
    second = workbook("a.xlsx", "Data", [("name",), ("A",)])

    merged = merge_workbooks([first, second], header_strategy="every_file")

    assert merged.sheets[0].rows == (("name",), ("A",), ("name",), ("B",))


def test_none_header_strategy_drops_each_source_header() -> None:
    first = workbook("b.xlsx", "Data", [("name",), ("B",)])
    second = workbook("a.xlsx", "Data", [("name",), ("A",)])

    merged = merge_workbooks([first, second], header_strategy="none")

    assert merged.sheets[0].rows == (("A",), ("B",))


def test_merge_order_is_deterministic_by_filename_and_sheet_name() -> None:
    first = WorkbookData(
        path=Path("z.xlsx"),
        sheets=(
            SheetData(name="Totals", rows=(("header",), ("z-total",))),
            SheetData(name="Data", rows=(("header",), ("z-data",))),
        ),
    )
    second = WorkbookData(
        path=Path("a.xlsx"),
        sheets=(
            SheetData(name="Totals", rows=(("header",), ("a-total",))),
            SheetData(name="Data", rows=(("header",), ("a-data",))),
        ),
    )

    merged = merge_workbooks([first, second], header_strategy="none")

    assert [(sheet.name, sheet.rows) for sheet in merged.sheets] == [
        ("Data", (("a-data",), ("z-data",))),
        ("Totals", (("a-total",), ("z-total",))),
    ]


def test_invalid_header_strategy_has_clear_error() -> None:
    source = workbook("input.xlsx", "Data", [("name",), ("A",)])

    with pytest.raises(MergeError, match="Unsupported header strategy"):
        merge_workbooks([source], header_strategy="invalid")  # type: ignore[arg-type]

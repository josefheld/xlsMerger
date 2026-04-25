# core

Common merge/validation engine.

## Reader API

`core.reader.read_workbook()` reads `.xls` and `.xlsx` files into one internal
representation:

- `WorkbookData`
- `SheetData`
- immutable tuple-based row data

Reader failures raise `ReaderError` with a specific message for unsupported file types,
missing files, empty files, unreadable/locked files, and corrupt workbooks.

## Merge API

`core.merge.merge_workbooks()` merges `WorkbookData` objects without row or column limits.
Inputs are processed deterministically by workbook filename and sheet name.

Header handling is configurable:

- `first_file` keeps the first header row per sheet
- `every_file` keeps every source header row
- `none` drops the first row from every source sheet

## Reporting API

`core.reporting.RunReport` captures processed files, row counts, warnings, errors, and the
stable process exit code. Reports can be written as JSON or CSV with `write_report()`.

Exit codes are:

- `0` success
- `1` validation error
- `2` system error

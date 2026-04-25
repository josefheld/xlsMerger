# xlsMerger

Merge multiple Excel files.

## Current status

This repository started as a script to merge multiple Excel `.xls` files from the `Originale` folder into one output workbook.

## Day 1 setup completed

The project now contains an initial product-oriented scaffolding:

- `core/` for merge and validation engine logic
- `profiles/` for mode-specific rules
- `cli/` for command-line interfaces
- `tests/` for automated tests
- `examples/` for sample input/configuration files

## Repository structure

```text
.
├── cli/
├── core/
├── docs/
├── examples/
├── profiles/
├── tests/
├── xlsMerger.py
└── pyproject.toml
```

## Tooling

`pyproject.toml` includes initial tooling setup for:

- Ruff (linting/import sorting)
- Pytest (test discovery and execution)

## Quick checks

```bash
python -m pip install -e ".[dev]"
xlsmerger --help
xlsmerger --report report.json
xlsmerger --mode supplier_normalizer --config profile.yml --report report.json
python -m pytest
python -m ruff check .
```

## Profiles

Runs use a selectable profile mode:

- `finance_close`
- `supplier_normalizer`
- `hr_consolidator`

Select a mode with `--mode`. Optional YAML configuration is passed with `--config`.

```yaml
mode: supplier_normalizer
options:
  header_strategy: first_file
```

### `finance_close`

Required input columns:

- `date`
- `account`
- `debit`
- `credit`
- `balance`

Optional input column:

- `description`

Dates must use `YYYY-MM-DD` or native spreadsheet date values. `debit`, `credit`, and
`balance` must be numeric. The profile checks that `sum(debit) - sum(credit)` matches the
closing balance within `balance_tolerance` (default `0.01`).

The normalized output schema is:

```text
date, account, description, debit, credit, balance
```

```bash
xlsmerger close.xlsx --mode finance_close --output finance-output.xlsx --report report.json
```

## Reports and exit codes

Regular CLI runs create a machine-readable report. JSON is the default; CSV can be selected
with `--report-format csv`.

```bash
xlsmerger input.xlsx --report run-report.json
xlsmerger input.xlsx --report run-report.csv --report-format csv
```

Exit codes are stable:

- `0` success
- `1` validation error
- `2` system error

## Legacy notes

The legacy script and build spec are still present:

- `xlsMerger.py`
- `xlsMerger.spec`

`xlsMerger.py` is importable under Python 3 and will be migrated into the new package structure
incrementally during upcoming sprint tasks.

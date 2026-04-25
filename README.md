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
xlsmerger run --report report.json
xlsmerger run input.xlsx --dry-run --preview-rows 5 --report report.json
xlsmerger validate --mode supplier_normalizer --config profile.yml --log-level debug --report report.json
xlsmerger profiles list
python -m pytest
python -m ruff check .
```

## Profiles

Runs use a selectable profile mode:

- `finance_close`
- `supplier_normalizer`
- `hr_consolidator`

Select a mode with `--mode` on `run` or `validate`. Optional YAML configuration is passed
with `--config`.

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
xlsmerger run close.xlsx --mode finance_close --output finance-output.xlsx --report report.json
```

### `supplier_normalizer`

Required canonical columns:

- `invoice_id`
- `order_id`
- `amount`

Normalized output schema:

```text
supplier, invoice_id, order_id, invoice_date, amount, currency
```

Common supplier layouts are mapped through built-in synonyms such as `Invoice No`,
`Rechnungsnummer`, `PO Number`, `Bestellnummer`, `Total`, and `Betrag`. Additional synonyms
and supplier-specific naming rules can be configured:

```yaml
mode: supplier_normalizer
options:
  supplier_name: Acme Raw
  supplier_name_map:
    Acme Raw: ACME GmbH
  default_currency: EUR
  column_synonyms:
    invoice_id:
      - Bill ID
    order_id:
      - Order Ref
    amount:
      - Gross
```

Duplicate `(invoice_id, order_id)` pairs are marked as report warnings.

### `hr_consolidator`

Required canonical columns:

- `employee_id`
- `first_name`
- `last_name`
- `hire_date`

Normalized output schema:

```text
employee_id, first_name, last_name, email, department, hire_date, termination_date
```

By default, `first_name`, `last_name`, and `email` are masked in output with
`***MASKED***`. Masking is configurable:

```yaml
mode: hr_consolidator
options:
  mask_fields:
    - email
  mask_token: "[redacted]"
```

`employee_id` must be 3-32 characters and contain only letters, numbers, `_`, or `-`.
`hire_date` and `termination_date` must use `YYYY-MM-DD` or native spreadsheet date values.

## Reports and exit codes

Regular CLI runs create a machine-readable report. JSON is the default; CSV can be selected
with `--report-format csv`.

```bash
xlsmerger run input.xlsx --report run-report.json
xlsmerger run input.xlsx --report run-report.csv --report-format csv
```

Dry-runs validate and transform inputs without writing an output workbook. Reports are still
written. Use `--preview-rows N` to print the first N transformed rows per sheet and inspect
the normalized shape before writing files.

```bash
xlsmerger run input.xlsx --mode hr_consolidator --dry-run --preview-rows 5 --report report.json
```

Use `--log-level info`, `--log-level warn`, or `--log-level debug` on `run` and `validate`
to control console logging.

Every report error includes:

- `message` with `Cause:` and `Recommendation:`
- `cause`
- `recommendation`

Example profile configurations live in `examples/` and can be executed directly:

```bash
xlsmerger validate --mode finance_close --config examples/finance.yml --report /tmp/finance-report.json
xlsmerger validate --mode supplier_normalizer --config examples/supplier.yml --report /tmp/supplier-report.json
xlsmerger validate --mode hr_consolidator --config examples/hr.yml --report /tmp/hr-report.json
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

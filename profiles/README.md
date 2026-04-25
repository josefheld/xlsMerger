# profiles

Mode-specific business rules live behind a shared profile interface:

- `validate(workbook, config)`
- `transform(workbook, config)`
- `postprocess(workbook, config)`

Registered modes:

- `finance_close`
- `supplier_normalizer`
- `hr_consolidator`

Profile configuration is YAML:

```yaml
mode: finance_close
options: {}
```

Configuration is validated before input files are processed. Mode mismatches, unsupported
top-level keys, missing files, invalid YAML, and non-mapping `options` values fail with a
validation report.

## finance_close

Required columns:

- `date`
- `account`
- `debit`
- `credit`
- `balance`

Optional column:

- `description`

Validation checks:

- required columns are present
- `date` values are spreadsheet dates or `YYYY-MM-DD`
- `debit`, `credit`, and `balance` values are numeric
- `sum(debit) - sum(credit)` matches the closing balance within `balance_tolerance`

Output schema:

```text
date, account, description, debit, credit, balance
```

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

## supplier_normalizer

Required canonical columns:

- `invoice_id`
- `order_id`
- `amount`

Optional canonical columns:

- `supplier`
- `invoice_date`
- `currency`

The profile maps common English and German supplier layouts into:

```text
supplier, invoice_id, order_id, invoice_date, amount, currency
```

YAML options:

- `supplier_name`: default supplier name when no supplier column exists
- `supplier_name_map`: source-to-target supplier name normalization
- `default_currency`: currency used when no currency column exists
- `column_synonyms`: additional header synonyms per canonical column

Duplicate `(invoice_id, order_id)` pairs are reported with `duplicate_invoice_order`
warnings.

## hr_consolidator

Required canonical columns:

- `employee_id`
- `first_name`
- `last_name`
- `hire_date`

Optional canonical columns:

- `email`
- `department`
- `termination_date`

Output schema:

```text
employee_id, first_name, last_name, email, department, hire_date, termination_date
```

Validation checks:

- `employee_id` is 3-32 characters and contains only letters, numbers, `_`, or `-`
- `hire_date` and `termination_date` are spreadsheet dates or `YYYY-MM-DD`

PII masking:

- default masked fields: `first_name`, `last_name`, `email`
- configurable with `mask_fields`
- mask value configurable with `mask_token`

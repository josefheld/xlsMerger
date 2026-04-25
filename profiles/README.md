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

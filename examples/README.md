# examples

Profile configuration examples:

- `finance.yml`
- `supplier.yml`
- `hr.yml`

They are directly executable with the CLI:

```bash
xlsmerger validate --mode finance_close --config examples/finance.yml --report /tmp/finance-report.json
xlsmerger validate --mode supplier_normalizer --config examples/supplier.yml --report /tmp/supplier-report.json
xlsmerger validate --mode hr_consolidator --config examples/hr.yml --report /tmp/hr-report.json
```

Add input workbook paths before the options when validating real files.

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
python -m pytest
python -m ruff check .
```

## Legacy notes

The legacy script and build spec are still present:

- `xlsMerger.py`
- `xlsMerger.spec`

They will be migrated incrementally during upcoming sprint tasks.

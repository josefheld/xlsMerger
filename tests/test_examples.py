import json
from pathlib import Path

from cli import main as cli_main


def test_example_profile_configs_are_directly_executable(tmp_path: Path) -> None:
    examples = [
        ("finance_close", Path("examples/finance.yml")),
        ("supplier_normalizer", Path("examples/supplier.yml")),
        ("hr_consolidator", Path("examples/hr.yml")),
    ]

    for mode, config_path in examples:
        report_path = tmp_path / f"{mode}.json"

        exit_code = cli_main.main(
            [
                "validate",
                "--mode",
                mode,
                "--config",
                str(config_path),
                "--report",
                str(report_path),
            ]
        )

        payload = json.loads(report_path.read_text(encoding="utf-8"))
        assert exit_code == 0
        assert payload["errors"] == []
        assert payload["warnings"][0]["code"] == "no_input_files"

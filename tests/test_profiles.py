import json
from pathlib import Path

import pytest
from openpyxl import Workbook

from cli import main as cli_main
from core.reader import SheetData, WorkbookData
from core.reporting import ReportIssue
from profiles import ProfileConfig
from profiles.base import PassthroughProfile
from profiles.config import ProfileConfigError, load_profile_config
from profiles.registry import PROFILE_REGISTRY, list_profile_names


def write_xlsx(path: Path, rows: list[list[object]]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    for row in rows:
        worksheet.append(row)
    workbook.save(path)


def test_builtin_profile_modes_are_registered() -> None:
    assert list_profile_names() == (
        "finance_close",
        "hr_consolidator",
        "supplier_normalizer",
    )


def test_profile_config_loads_valid_yaml(tmp_path: Path) -> None:
    config_path = tmp_path / "profile.yml"
    config_path.write_text(
        "mode: supplier_normalizer\noptions:\n  header_strategy: first_file\n",
        encoding="utf-8",
    )

    config = load_profile_config(config_path, mode="supplier_normalizer")

    assert config == ProfileConfig(
        mode="supplier_normalizer",
        options={"header_strategy": "first_file"},
    )


def test_profile_config_rejects_mode_mismatch(tmp_path: Path) -> None:
    config_path = tmp_path / "profile.yml"
    config_path.write_text("mode: finance_close\noptions: {}\n", encoding="utf-8")

    with pytest.raises(ProfileConfigError, match="but '--mode hr_consolidator' was selected"):
        load_profile_config(config_path, mode="hr_consolidator")


def test_profile_config_rejects_non_mapping_options(tmp_path: Path) -> None:
    config_path = tmp_path / "profile.yml"
    config_path.write_text("mode: finance_close\noptions:\n  - invalid\n", encoding="utf-8")

    with pytest.raises(ProfileConfigError, match="field 'options' must be a mapping"):
        load_profile_config(config_path, mode="finance_close")


def test_cli_reports_unknown_mode_as_validation_error(tmp_path: Path) -> None:
    report_path = tmp_path / "report.json"

    exit_code = cli_main.main(["run", "--mode", "unknown", "--report", str(report_path)])

    payload = json.loads(report_path.read_text(encoding="utf-8"))
    assert exit_code == 1
    assert payload["errors"][0]["code"] == "profile_config_error"
    assert "Unknown mode 'unknown'" in payload["errors"][0]["message"]


def test_cli_uses_registered_profile_without_core_changes(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    class TestDropHeaderProfile(PassthroughProfile):
        name = "test_drop_header"
        description = "Test profile"

        def transform(self, workbook: WorkbookData, config: ProfileConfig) -> WorkbookData:
            return WorkbookData(
                path=workbook.path,
                sheets=tuple(
                    SheetData(name=sheet.name, rows=sheet.rows[1:])
                    for sheet in workbook.sheets
                ),
            )

        def postprocess(
            self, workbook: WorkbookData, config: ProfileConfig
        ) -> tuple[ReportIssue, ...]:
            return (
                ReportIssue(
                    code="test_profile_applied",
                    message=f"Applied {self.name}",
                ),
            )

    input_path = tmp_path / "input.xlsx"
    report_path = tmp_path / "report.json"
    write_xlsx(input_path, [["header"], ["value"]])
    monkeypatch.setitem(PROFILE_REGISTRY, TestDropHeaderProfile.name, TestDropHeaderProfile)

    exit_code = cli_main.main(
        [
            "run",
            str(input_path),
            "--mode",
            TestDropHeaderProfile.name,
            "--report",
            str(report_path),
        ]
    )

    payload = json.loads(report_path.read_text(encoding="utf-8"))
    assert exit_code == 0
    assert payload["files"][0]["rows"] == 1
    assert payload["warnings"][0]["code"] == "test_profile_applied"

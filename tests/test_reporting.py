import csv
import json
from pathlib import Path

import pytest
from openpyxl import Workbook

from cli import main as cli_main
from core.reporting import (
    ExitCode,
    ReportFile,
    ReportIssue,
    RunReport,
    report_exit_code,
    write_report,
)


def write_xlsx(path: Path, rows: list[list[object]]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    for row in rows:
        worksheet.append(row)
    workbook.save(path)


def test_json_report_export_contains_files_rows_warnings_and_errors(tmp_path: Path) -> None:
    report_path = tmp_path / "report.json"
    report = RunReport(
        exit_code=int(ExitCode.VALIDATION_ERROR),
        files=(ReportFile(path="input.xlsx", sheets=2, rows=42),),
        total_rows=42,
        warnings=(ReportIssue(code="sample_warning", message="warning message"),),
        errors=(ReportIssue(code="sample_error", message="error message", path="bad.xlsx"),),
    )

    write_report(report, report_path, report_format="json")

    payload = json.loads(report_path.read_text(encoding="utf-8"))
    assert payload["exit_code"] == 1
    assert payload["files"] == [{"path": "input.xlsx", "rows": 42, "sheets": 2}]
    assert payload["total_rows"] == 42
    assert payload["warnings"][0]["code"] == "sample_warning"
    assert payload["errors"][0]["message"] == "error message"


def test_csv_report_export_contains_summary_file_and_issue_rows(tmp_path: Path) -> None:
    report_path = tmp_path / "report.csv"
    report = RunReport(
        exit_code=int(ExitCode.SUCCESS),
        files=(ReportFile(path="input.xlsx", sheets=1, rows=3),),
        total_rows=3,
        warnings=(ReportIssue(code="sample_warning", message="warning message"),),
    )

    write_report(report, report_path, report_format="csv")

    rows = list(csv.DictReader(report_path.open(encoding="utf-8", newline="")))
    assert rows[0]["record_type"] == "summary"
    assert rows[0]["total_files"] == "1"
    assert rows[1]["record_type"] == "file"
    assert rows[1]["rows"] == "3"
    assert rows[2]["record_type"] == "warning"
    assert rows[2]["code"] == "sample_warning"


def test_cli_run_writes_machine_readable_report_for_success(tmp_path: Path) -> None:
    input_path = tmp_path / "input.xlsx"
    report_path = tmp_path / "run-report.json"
    write_xlsx(
        input_path,
        [
            ["employee_id", "first_name", "last_name", "hire_date"],
            ["EMP-001", "Ada", "Lovelace", "2026-01-01"],
        ],
    )

    exit_code = cli_main.main(
        ["run", str(input_path), "--mode", "hr_consolidator", "--report", str(report_path)]
    )

    payload = json.loads(report_path.read_text(encoding="utf-8"))
    assert exit_code == ExitCode.SUCCESS
    assert payload["exit_code"] == 0
    assert payload["files"][0]["path"] == str(input_path)
    assert payload["files"][0]["rows"] == 2


def test_cli_run_writes_machine_readable_report_for_validation_error(tmp_path: Path) -> None:
    input_path = tmp_path / "missing.xlsx"
    report_path = tmp_path / "run-report.json"

    exit_code = cli_main.main(["run", str(input_path), "--report", str(report_path)])

    payload = json.loads(report_path.read_text(encoding="utf-8"))
    assert exit_code == ExitCode.VALIDATION_ERROR
    assert payload["exit_code"] == 1
    assert payload["errors"][0]["code"] == "reader_error"
    assert "does not exist" in payload["errors"][0]["message"]


def test_cli_run_writes_csv_report_when_requested(tmp_path: Path) -> None:
    input_path = tmp_path / "input.xlsx"
    report_path = tmp_path / "run-report.csv"
    write_xlsx(
        input_path,
        [
            ["employee_id", "first_name", "last_name", "hire_date"],
            ["EMP-001", "Ada", "Lovelace", "2026-01-01"],
        ],
    )

    exit_code = cli_main.main(
        [
            "run",
            str(input_path),
            "--mode",
            "hr_consolidator",
            "--report",
            str(report_path),
            "--report-format",
            "csv",
        ]
    )

    rows = list(csv.DictReader(report_path.open(encoding="utf-8", newline="")))
    assert exit_code == ExitCode.SUCCESS
    assert rows[0]["record_type"] == "summary"
    assert rows[1]["record_type"] == "file"


def test_cli_validate_writes_report_without_transforming(tmp_path: Path) -> None:
    input_path = tmp_path / "input.xlsx"
    report_path = tmp_path / "validate-report.json"
    write_xlsx(
        input_path,
        [
            ["employee_id", "first_name", "last_name", "hire_date"],
            ["EMP-001", "Ada", "Lovelace", "2026-01-01"],
        ],
    )

    exit_code = cli_main.main(
        [
            "validate",
            str(input_path),
            "--mode",
            "hr_consolidator",
            "--report",
            str(report_path),
        ]
    )

    payload = json.loads(report_path.read_text(encoding="utf-8"))
    assert exit_code == ExitCode.SUCCESS
    assert payload["exit_code"] == 0
    assert payload["files"][0]["rows"] == 2
    assert payload["errors"] == []


def test_cli_system_error_uses_stable_exit_code(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    input_path = tmp_path / "input.xlsx"
    report_path = tmp_path / "run-report.json"
    write_xlsx(input_path, [["name"], ["A"]])

    def raise_runtime_error(path: Path) -> None:
        raise RuntimeError("boom")

    monkeypatch.setattr(cli_main, "read_workbook", raise_runtime_error)

    exit_code = cli_main.main(["run", str(input_path), "--report", str(report_path)])

    payload = json.loads(report_path.read_text(encoding="utf-8"))
    assert exit_code == ExitCode.SYSTEM_ERROR
    assert payload["exit_code"] == 2
    assert payload["errors"][0]["code"] == "system_error"


def test_exit_code_mapping_is_stable() -> None:
    assert int(ExitCode.SUCCESS) == 0
    assert int(ExitCode.VALIDATION_ERROR) == 1
    assert int(ExitCode.SYSTEM_ERROR) == 2
    assert report_exit_code(has_validation_errors=False) == ExitCode.SUCCESS
    assert report_exit_code(has_validation_errors=True) == ExitCode.VALIDATION_ERROR
    assert (
        report_exit_code(has_validation_errors=False, has_system_errors=True)
        == ExitCode.SYSTEM_ERROR
    )

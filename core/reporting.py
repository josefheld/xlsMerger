"""Machine-readable run reports and stable process exit codes."""

from __future__ import annotations

import csv
import json
from dataclasses import asdict, dataclass
from enum import IntEnum
from pathlib import Path
from typing import Literal


class ExitCode(IntEnum):
    SUCCESS = 0
    VALIDATION_ERROR = 1
    SYSTEM_ERROR = 2


ReportFormat = Literal["json", "csv"]


@dataclass(frozen=True)
class ReportFile:
    path: str
    sheets: int
    rows: int


@dataclass(frozen=True)
class ReportIssue:
    code: str
    message: str
    path: str | None = None
    sheet: str | None = None


@dataclass(frozen=True)
class RunReport:
    exit_code: int
    files: tuple[ReportFile, ...] = ()
    total_rows: int = 0
    warnings: tuple[ReportIssue, ...] = ()
    errors: tuple[ReportIssue, ...] = ()


def report_exit_code(*, has_validation_errors: bool, has_system_errors: bool = False) -> ExitCode:
    if has_system_errors:
        return ExitCode.SYSTEM_ERROR
    if has_validation_errors:
        return ExitCode.VALIDATION_ERROR
    return ExitCode.SUCCESS


def write_report(
    report: RunReport,
    path: str | Path,
    *,
    report_format: ReportFormat | None = None,
) -> None:
    report_path = Path(path)
    selected_format = report_format or _infer_report_format(report_path)
    report_path.parent.mkdir(parents=True, exist_ok=True)

    if selected_format == "json":
        _write_json_report(report, report_path)
        return

    if selected_format == "csv":
        _write_csv_report(report, report_path)
        return

    raise ValueError(f"Unsupported report format '{selected_format}': expected json or csv")


def _infer_report_format(path: Path) -> ReportFormat:
    if path.suffix.lower() == ".csv":
        return "csv"
    return "json"


def _write_json_report(report: RunReport, path: Path) -> None:
    path.write_text(json.dumps(asdict(report), indent=2, sort_keys=True) + "\n", encoding="utf-8")


def _write_csv_report(report: RunReport, path: Path) -> None:
    with path.open("w", encoding="utf-8", newline="") as report_file:
        writer = csv.DictWriter(
            report_file,
            fieldnames=[
                "record_type",
                "path",
                "sheet",
                "sheets",
                "rows",
                "severity",
                "code",
                "message",
                "total_files",
                "total_rows",
                "exit_code",
            ],
        )
        writer.writeheader()
        writer.writerow(
            {
                "record_type": "summary",
                "total_files": len(report.files),
                "total_rows": report.total_rows,
                "exit_code": report.exit_code,
            }
        )

        for file_report in report.files:
            writer.writerow(
                {
                    "record_type": "file",
                    "path": file_report.path,
                    "sheets": file_report.sheets,
                    "rows": file_report.rows,
                }
            )

        for warning in report.warnings:
            writer.writerow(_issue_to_csv_row("warning", warning))

        for error in report.errors:
            writer.writerow(_issue_to_csv_row("error", error))


def _issue_to_csv_row(record_type: str, issue: ReportIssue) -> dict[str, str | None]:
    return {
        "record_type": record_type,
        "path": issue.path,
        "sheet": issue.sheet,
        "severity": record_type,
        "code": issue.code,
        "message": issue.message,
    }

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
    cause: str | None = None
    recommendation: str | None = None


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
    report = with_error_troubleshooting(report)
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
                "cause",
                "recommendation",
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
        "cause": issue.cause,
        "recommendation": issue.recommendation,
    }


def with_error_troubleshooting(report: RunReport) -> RunReport:
    return RunReport(
        exit_code=report.exit_code,
        files=report.files,
        total_rows=report.total_rows,
        warnings=report.warnings,
        errors=tuple(_with_troubleshooting(issue) for issue in report.errors),
    )


def _with_troubleshooting(issue: ReportIssue) -> ReportIssue:
    cause = issue.cause or _default_cause(issue.code)
    recommendation = issue.recommendation or _default_recommendation(issue.code)
    message = issue.message
    if "Cause:" not in message:
        message = f"{message} Cause: {cause}"
    if "Recommendation:" not in message:
        message = f"{message} Recommendation: {recommendation}"

    return ReportIssue(
        code=issue.code,
        message=message,
        path=issue.path,
        sheet=issue.sheet,
        cause=cause,
        recommendation=recommendation,
    )


def _default_cause(code: str) -> str:
    causes = {
        "missing_required_column": (
            "The input sheet does not contain every column required by the selected profile."
        ),
        "invalid_column_type": (
            "A cell value does not match the type required by the selected profile."
        ),
        "balance_mismatch": "The finance totals do not reconcile to the closing balance.",
        "invalid_hr_record": "An HR row contains an invalid employee ID or date value.",
        "reader_error": "The input workbook could not be opened or parsed.",
        "profile_config_error": "The selected mode or YAML profile configuration is invalid.",
        "profile_error": (
            "The selected profile failed while validating or transforming workbook data."
        ),
        "output_error": "The output workbook could not be written.",
        "system_error": "An unexpected system error interrupted the run.",
    }
    return causes.get(code, "The run produced a validation or processing error.")


def _default_recommendation(code: str) -> str:
    recommendations = {
        "missing_required_column": (
            "Add the missing columns or configure supported column synonyms, then rerun."
        ),
        "invalid_column_type": "Correct the value in the reported row and column, then rerun.",
        "balance_mismatch": (
            "Check debit, credit, and balance values for the reported sheet, then rerun."
        ),
        "invalid_hr_record": (
            "Fix the reported employee ID or date value using the documented format, then rerun."
        ),
        "reader_error": (
            "Check that the file exists, is readable, and is a valid .xls or .xlsx workbook."
        ),
        "profile_config_error": (
            "Fix the mode or YAML configuration file and rerun the same command."
        ),
        "profile_error": (
            "Review the reported file and profile configuration; rerun with --log-level debug "
            "if needed."
        ),
        "output_error": (
            "Check the output path permissions and close any open workbook with the same name, "
            "then rerun."
        ),
        "system_error": "Rerun with --log-level debug and inspect the reported exception.",
    }
    return recommendations.get(code, "Review the file, sheet, and message details, then rerun.")

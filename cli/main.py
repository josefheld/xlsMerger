"""Command-line entry point for xlsmerger."""

from __future__ import annotations

import argparse
from pathlib import Path

from core.reader import ReaderError, read_workbook
from core.reporting import (
    ExitCode,
    ReportFile,
    ReportFormat,
    ReportIssue,
    RunReport,
    report_exit_code,
    write_report,
)
from profiles import (
    ProfileConfigError,
    ProfileRegistryError,
    get_profile,
    list_profile_names,
    load_profile_config,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="xlsmerger",
        description="Merge and normalize Excel files.",
    )
    parser.add_argument("inputs", nargs="*", type=Path, help="Excel files to read")
    parser.add_argument(
        "--mode",
        default="finance_close",
        help=f"Profile mode to apply ({', '.join(list_profile_names())})",
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=None,
        help="YAML profile configuration",
    )
    parser.add_argument(
        "--report",
        type=Path,
        default=Path("report.json"),
        help="Machine-readable report path",
    )
    parser.add_argument(
        "--report-format",
        choices=["json", "csv"],
        default=None,
        help="Report format; inferred from --report when omitted",
    )
    parser.add_argument("--version", action="version", version="xlsmerger 0.1.0")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    report = build_run_report(args.inputs, mode=args.mode, config_path=args.config)
    write_report(report, args.report, report_format=args.report_format)
    return report.exit_code


def build_run_report(
    inputs: list[Path],
    *,
    mode: str = "finance_close",
    config_path: Path | None = None,
) -> RunReport:
    files: list[ReportFile] = []
    warnings: list[ReportIssue] = []
    errors: list[ReportIssue] = []
    has_system_errors = False

    try:
        profile = get_profile(mode)
        profile_config = load_profile_config(config_path, mode=mode)
    except (ProfileRegistryError, ProfileConfigError) as exc:
        errors.append(
            ReportIssue(
                code="profile_config_error",
                message=str(exc),
                path=str(config_path) if config_path else None,
            )
        )
        return RunReport(
            exit_code=int(ExitCode.VALIDATION_ERROR),
            errors=tuple(errors),
        )

    if not inputs:
        warnings.append(
            ReportIssue(
                code="no_input_files",
                message="No input files were provided; no workbooks were processed",
            )
        )

    for input_path in inputs:
        try:
            workbook = read_workbook(input_path)
        except ReaderError as exc:
            errors.append(
                ReportIssue(
                    code="reader_error",
                    message=str(exc),
                    path=str(input_path),
                )
            )
            continue
        except Exception as exc:
            has_system_errors = True
            errors.append(
                ReportIssue(
                    code="system_error",
                    message=f"Unexpected error while reading '{input_path}': {exc}",
                    path=str(input_path),
                )
            )
            continue

        try:
            validation_errors = profile.validate(workbook, profile_config)
            transformed_workbook = profile.transform(workbook, profile_config)
            profile_warnings = profile.postprocess(transformed_workbook, profile_config)
        except Exception as exc:
            has_system_errors = True
            errors.append(
                ReportIssue(
                    code="profile_error",
                    message=(
                        f"Unexpected profile error in mode '{mode}' "
                        f"while processing '{input_path}': {exc}"
                    ),
                    path=str(input_path),
                )
            )
            continue

        errors.extend(validation_errors)
        warnings.extend(profile_warnings)
        row_count = sum(len(sheet.rows) for sheet in transformed_workbook.sheets)
        files.append(
            ReportFile(
                path=str(transformed_workbook.path),
                sheets=len(transformed_workbook.sheets),
                rows=row_count,
            )
        )

    exit_code = report_exit_code(
        has_validation_errors=bool(errors) and not has_system_errors,
        has_system_errors=has_system_errors,
    )
    return RunReport(
        exit_code=int(exit_code),
        files=tuple(files),
        total_rows=sum(file.rows for file in files),
        warnings=tuple(warnings),
        errors=tuple(errors),
    )


__all__ = ["ExitCode", "ReportFormat", "build_parser", "build_run_report", "main"]


if __name__ == "__main__":
    raise SystemExit(main())

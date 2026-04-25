"""Command-line entry point for xlsmerger."""

from __future__ import annotations

import argparse
import logging
from dataclasses import dataclass
from pathlib import Path

from core.merge import MergeError, merge_workbooks
from core.reader import ReaderError, WorkbookData, read_workbook
from core.reporting import (
    ExitCode,
    ReportFile,
    ReportFormat,
    ReportIssue,
    RunReport,
    report_exit_code,
    write_report,
)
from core.writer import WriterError, write_workbook
from profiles import (
    ProfileConfigError,
    ProfileRegistryError,
    get_profile,
    iter_profiles,
    list_profile_names,
    load_profile_config,
)

LOGGER = logging.getLogger(__name__)
LOG_LEVELS = {"debug": logging.DEBUG, "info": logging.INFO, "warn": logging.WARNING}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="xlsmerger",
        description="Merge and normalize Excel files.",
    )
    parser.add_argument("--version", action="version", version="xlsmerger 0.1.0")
    subparsers = parser.add_subparsers(dest="command", required=True)

    run_parser = subparsers.add_parser(
        "run",
        help="Read, validate, transform, and optionally write output",
        description="Read, validate, transform, and optionally write output.",
    )
    add_processing_arguments(run_parser)
    run_parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Optional .xlsx output workbook path",
    )
    run_parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate and transform without writing the output workbook",
    )
    run_parser.add_argument(
        "--preview-rows",
        type=int,
        default=0,
        help="Print the first N transformed rows per sheet",
    )
    run_parser.set_defaults(func=run_command)

    validate_parser = subparsers.add_parser(
        "validate",
        help="Validate input workbooks without writing output",
        description="Validate input workbooks without writing output.",
    )
    add_processing_arguments(validate_parser)
    validate_parser.set_defaults(func=validate_command)

    profiles_parser = subparsers.add_parser(
        "profiles",
        help="Inspect available profile modes",
        description="Inspect available profile modes.",
    )
    profiles_subparsers = profiles_parser.add_subparsers(
        dest="profiles_command", required=True
    )
    profiles_list_parser = profiles_subparsers.add_parser(
        "list",
        help="List registered profile modes",
        description="List registered profile modes.",
    )
    profiles_list_parser.set_defaults(func=profiles_list_command)

    return parser


@dataclass(frozen=True)
class ProcessingResult:
    report: RunReport
    transformed_workbooks: tuple[WorkbookData, ...]


def add_processing_arguments(parser: argparse.ArgumentParser) -> None:
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
    parser.add_argument(
        "--log-level",
        choices=["info", "warn", "debug"],
        default="warn",
        help="Console logging level",
    )


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


def run_command(args: argparse.Namespace) -> int:
    configure_logging(args.log_level)
    LOGGER.info("Starting run command with mode '%s'", args.mode)
    result = build_processing_result(
        args.inputs,
        mode=args.mode,
        config_path=args.config,
        output_path=args.output,
        dry_run=args.dry_run,
    )
    report = result.report
    write_report(report, args.report, report_format=args.report_format)
    log_report_summary(report)
    print_run_statistics(report, dry_run=args.dry_run, output_path=args.output)
    if args.preview_rows > 0:
        print_preview(result.transformed_workbooks, row_limit=args.preview_rows)
    return report.exit_code


def validate_command(args: argparse.Namespace) -> int:
    configure_logging(args.log_level)
    LOGGER.info("Starting validate command with mode '%s'", args.mode)
    report = build_run_report(
        args.inputs,
        mode=args.mode,
        config_path=args.config,
        apply_transforms=False,
    )
    write_report(report, args.report, report_format=args.report_format)
    log_report_summary(report)
    return report.exit_code


def profiles_list_command(args: argparse.Namespace) -> int:
    for profile_type in iter_profiles():
        print(f"{profile_type.name}\t{profile_type.description}")
    return int(ExitCode.SUCCESS)


def configure_logging(log_level: str) -> None:
    logging.basicConfig(
        level=LOG_LEVELS[log_level],
        format="%(levelname)s %(message)s",
        force=True,
    )


def log_report_summary(report: RunReport) -> None:
    if report.errors:
        LOGGER.warning("Run completed with %s error(s)", len(report.errors))
    elif report.warnings:
        LOGGER.warning("Run completed with %s warning(s)", len(report.warnings))
    else:
        LOGGER.info("Run completed successfully")


def build_run_report(
    inputs: list[Path],
    *,
    mode: str = "finance_close",
    config_path: Path | None = None,
    output_path: Path | None = None,
    apply_transforms: bool = True,
    dry_run: bool = False,
) -> RunReport:
    return build_processing_result(
        inputs,
        mode=mode,
        config_path=config_path,
        output_path=output_path,
        apply_transforms=apply_transforms,
        dry_run=dry_run,
    ).report


def build_processing_result(
    inputs: list[Path],
    *,
    mode: str = "finance_close",
    config_path: Path | None = None,
    output_path: Path | None = None,
    apply_transforms: bool = True,
    dry_run: bool = False,
) -> ProcessingResult:
    files: list[ReportFile] = []
    warnings: list[ReportIssue] = []
    errors: list[ReportIssue] = []
    transformed_workbooks: list[WorkbookData] = []
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
        return ProcessingResult(
            report=RunReport(
                exit_code=int(ExitCode.VALIDATION_ERROR),
                errors=tuple(errors),
            ),
            transformed_workbooks=(),
        )

    if not inputs:
        warnings.append(
            ReportIssue(
                code="no_input_files",
                message="No input files were provided; no workbooks were processed",
            )
        )

    for input_path in inputs:
        LOGGER.debug("Reading workbook '%s'", input_path)
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
            LOGGER.debug("Validating workbook '%s' with mode '%s'", input_path, mode)
            validation_errors = profile.validate(workbook, profile_config)
        except Exception as exc:
            has_system_errors = True
            errors.append(
                ReportIssue(
                    code="profile_error",
                    message=(
                        f"Unexpected profile error in mode '{mode}' "
                        f"while validating '{input_path}': {exc}"
                    ),
                    path=str(input_path),
                )
            )
            continue

        if validation_errors:
            errors.extend(validation_errors)
            files.append(
                ReportFile(
                    path=str(workbook.path),
                    sheets=len(workbook.sheets),
                    rows=sum(len(sheet.rows) for sheet in workbook.sheets),
                )
            )
            continue

        if not apply_transforms:
            files.append(
                ReportFile(
                    path=str(workbook.path),
                    sheets=len(workbook.sheets),
                    rows=sum(len(sheet.rows) for sheet in workbook.sheets),
                )
            )
            continue

        try:
            LOGGER.debug("Transforming workbook '%s' with mode '%s'", input_path, mode)
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

        warnings.extend(profile_warnings)
        if profile_warnings:
            LOGGER.debug(
                "Profile produced %s warning(s) for workbook '%s'",
                len(profile_warnings),
                input_path,
            )
        row_count = sum(len(sheet.rows) for sheet in transformed_workbook.sheets)
        files.append(
            ReportFile(
                path=str(transformed_workbook.path),
                sheets=len(transformed_workbook.sheets),
                rows=row_count,
            )
        )
        transformed_workbooks.append(transformed_workbook)

    if (
        dry_run
        and output_path is not None
        and transformed_workbooks
        and not errors
        and not has_system_errors
    ):
        warnings.append(
            ReportIssue(
                code="dry_run_output_skipped",
                message=f"Dry-run enabled; output workbook '{output_path}' was not written",
                path=str(output_path),
            )
        )

    if (
        output_path is not None
        and not dry_run
        and transformed_workbooks
        and not errors
        and not has_system_errors
    ):
        header_strategy = str(profile_config.options.get("header_strategy", "first_file"))
        try:
            LOGGER.debug("Writing output workbook '%s'", output_path)
            merged_workbook = merge_workbooks(
                transformed_workbooks,
                header_strategy=header_strategy,  # type: ignore[arg-type]
                output_path=output_path,
            )
            write_workbook(merged_workbook, output_path)
        except (MergeError, WriterError) as exc:
            has_system_errors = True
            errors.append(
                ReportIssue(
                    code="output_error",
                    message=f"Output workbook '{output_path}' could not be written: {exc}",
                    path=str(output_path),
                )
            )

    exit_code = report_exit_code(
        has_validation_errors=bool(errors) and not has_system_errors,
        has_system_errors=has_system_errors,
    )
    return ProcessingResult(
        report=RunReport(
            exit_code=int(exit_code),
            files=tuple(files),
            total_rows=sum(file.rows for file in files),
            warnings=tuple(warnings),
            errors=tuple(errors),
        ),
        transformed_workbooks=tuple(transformed_workbooks),
    )


def print_run_statistics(
    report: RunReport, *, dry_run: bool, output_path: Path | None
) -> None:
    print("Run statistics")
    print(f"  files: {len(report.files)}")
    print(f"  rows: {report.total_rows}")
    print(f"  warnings: {len(report.warnings)}")
    print(f"  errors: {len(report.errors)}")
    print(f"  exit_code: {report.exit_code}")
    if dry_run:
        print("  dry_run: true")
    if output_path is not None:
        print(f"  output: {output_path}")


def print_preview(workbooks: tuple[WorkbookData, ...], *, row_limit: int) -> None:
    print(f"Preview first {row_limit} rows")
    for workbook in workbooks:
        print(f"Workbook: {workbook.path}")
        for sheet in workbook.sheets:
            print(f"Sheet: {sheet.name}")
            for row in sheet.rows[:row_limit]:
                print("  " + "\t".join("" if value is None else str(value) for value in row))


__all__ = [
    "ExitCode",
    "ReportFormat",
    "build_processing_result",
    "build_parser",
    "build_run_report",
    "configure_logging",
    "main",
    "profiles_list_command",
    "run_command",
    "validate_command",
]


if __name__ == "__main__":
    raise SystemExit(main())

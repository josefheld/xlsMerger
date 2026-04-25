"""Basic smoke tests for the new project scaffolding."""

import pytest

import xlsMerger
from cli.main import main


def test_project_scaffold_is_present() -> None:
    assert True


def test_cli_help_starts(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as exc_info:
        main(["--help"])

    assert exc_info.value.code == 0
    output = capsys.readouterr().out
    assert "xlsmerger" in output
    assert "run" in output
    assert "validate" in output
    assert "profiles" in output


@pytest.mark.parametrize("command", ["run", "validate"])
def test_cli_command_help_is_available(
    command: str, capsys: pytest.CaptureFixture[str]
) -> None:
    with pytest.raises(SystemExit) as exc_info:
        main([command, "--help"])

    assert exc_info.value.code == 0
    output = capsys.readouterr().out
    assert "--mode" in output
    assert "--config" in output
    assert "--report" in output
    assert "--log-level" in output
    if command == "run":
        assert "--dry-run" in output
        assert "--preview-rows" in output


def test_profiles_list_command_outputs_modes(capsys: pytest.CaptureFixture[str]) -> None:
    exit_code = main(["profiles", "list"])

    output = capsys.readouterr().out
    assert exit_code == 0
    assert "finance_close" in output
    assert "supplier_normalizer" in output
    assert "hr_consolidator" in output


def test_legacy_script_imports_under_python3() -> None:
    assert callable(xlsMerger.open_xls_as_xlsx)

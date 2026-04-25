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
    assert "xlsmerger" in capsys.readouterr().out


def test_legacy_script_imports_under_python3() -> None:
    assert callable(xlsMerger.open_xls_as_xlsx)

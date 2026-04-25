"""Profile interface for mode-specific validation and transforms."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, ClassVar, Protocol

from core.reader import WorkbookData
from core.reporting import ReportIssue


@dataclass(frozen=True)
class ProfileConfig:
    mode: str
    options: dict[str, Any]


class Profile(Protocol):
    name: ClassVar[str]
    description: ClassVar[str]

    def validate(
        self, workbook: WorkbookData, config: ProfileConfig
    ) -> tuple[ReportIssue, ...]: ...

    def transform(self, workbook: WorkbookData, config: ProfileConfig) -> WorkbookData: ...

    def postprocess(
        self, workbook: WorkbookData, config: ProfileConfig
    ) -> tuple[ReportIssue, ...]: ...


class PassthroughProfile:
    """Base implementation for profiles that do not yet change workbook data."""

    name: ClassVar[str]
    description: ClassVar[str]

    def validate(self, workbook: WorkbookData, config: ProfileConfig) -> tuple[ReportIssue, ...]:
        return ()

    def transform(self, workbook: WorkbookData, config: ProfileConfig) -> WorkbookData:
        return workbook

    def postprocess(self, workbook: WorkbookData, config: ProfileConfig) -> tuple[ReportIssue, ...]:
        return ()

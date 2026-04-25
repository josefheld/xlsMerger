"""Profile registry for mode selection."""

from __future__ import annotations

from collections.abc import Iterable

from profiles.base import PassthroughProfile, Profile


class ProfileRegistryError(Exception):
    """Raised when a profile mode is not registered."""


class FinanceCloseProfile(PassthroughProfile):
    name = "finance_close"
    description = "Finance close workbook consolidation"


class SupplierNormalizerProfile(PassthroughProfile):
    name = "supplier_normalizer"
    description = "Supplier workbook normalization"


class HrConsolidatorProfile(PassthroughProfile):
    name = "hr_consolidator"
    description = "HR workbook consolidation"


PROFILE_REGISTRY: dict[str, type[Profile]] = {
    FinanceCloseProfile.name: FinanceCloseProfile,
    SupplierNormalizerProfile.name: SupplierNormalizerProfile,
    HrConsolidatorProfile.name: HrConsolidatorProfile,
}


def list_profile_names() -> tuple[str, ...]:
    return tuple(sorted(PROFILE_REGISTRY))


def iter_profiles() -> Iterable[type[Profile]]:
    for name in list_profile_names():
        yield PROFILE_REGISTRY[name]


def get_profile(mode: str) -> Profile:
    try:
        profile_type = PROFILE_REGISTRY[mode]
    except KeyError as exc:
        supported = ", ".join(list_profile_names())
        raise ProfileRegistryError(
            f"Unknown mode '{mode}': expected one of {supported}"
        ) from exc

    return profile_type()

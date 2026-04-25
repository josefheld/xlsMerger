"""YAML profile configuration loading and validation."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml

from profiles.base import ProfileConfig


class ProfileConfigError(Exception):
    """Raised when profile configuration is missing or invalid."""


def load_profile_config(path: str | Path | None, *, mode: str) -> ProfileConfig:
    if path is None:
        return ProfileConfig(mode=mode, options={})

    config_path = Path(path)
    try:
        payload = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise ProfileConfigError(f"Profile configuration '{config_path}' does not exist") from exc
    except PermissionError as exc:
        raise ProfileConfigError(
            f"Profile configuration '{config_path}' is not readable"
        ) from exc
    except yaml.YAMLError as exc:
        raise ProfileConfigError(
            f"Profile configuration '{config_path}' is not valid YAML: {exc}"
        ) from exc

    if payload is None:
        payload = {}

    if not isinstance(payload, dict):
        raise ProfileConfigError(
            f"Profile configuration '{config_path}' must be a YAML mapping"
        )

    unknown_keys = sorted(set(payload) - {"mode", "options"})
    if unknown_keys:
        joined = ", ".join(str(key) for key in unknown_keys)
        raise ProfileConfigError(
            f"Profile configuration '{config_path}' contains unsupported keys: {joined}"
        )

    configured_mode = payload.get("mode", mode)
    if not isinstance(configured_mode, str) or not configured_mode:
        raise ProfileConfigError(
            f"Profile configuration '{config_path}' field 'mode' must be a non-empty string"
        )

    if configured_mode != mode:
        raise ProfileConfigError(
            f"Profile configuration '{config_path}' is for mode '{configured_mode}', "
            f"but '--mode {mode}' was selected"
        )

    options = payload.get("options", {})
    if options is None:
        options = {}

    if not isinstance(options, dict):
        raise ProfileConfigError(
            f"Profile configuration '{config_path}' field 'options' must be a mapping"
        )

    return ProfileConfig(mode=mode, options=dict(options))


def dump_example_config(mode: str, options: dict[str, Any] | None = None) -> dict[str, Any]:
    return {"mode": mode, "options": options or {}}

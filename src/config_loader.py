"""Shared settings loader.

Loads ``config/settings.json`` and, if present, deep-merges
``config/settings.local.json`` on top of it. The local override file is
gitignored so real account identifiers, Flex tokens, and other personal
values never enter version control.

Precedence (highest wins): ``settings.local.json`` values override
``settings.json`` values for the same key path. Nested dicts merge key by
key; scalar and list values from the local file replace those from the
base file entirely.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


CONFIG_FILENAME = "settings.json"
LOCAL_OVERRIDE_FILENAME = "settings.local.json"


def _deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    merged: dict[str, Any] = dict(base)
    for key, value in override.items():
        prior = merged.get(key)
        if isinstance(prior, dict) and isinstance(value, dict):
            merged[key] = _deep_merge(prior, value)
        else:
            merged[key] = value
    return merged


def load_config(project_root: Path) -> dict[str, Any]:
    """Return the effective settings dict for ``project_root``.

    Raises ``FileNotFoundError`` if the base ``settings.json`` is missing.
    Silently ignores a missing local override file.
    """
    config_dir = Path(project_root) / "config"
    base_path = config_dir / CONFIG_FILENAME
    local_path = config_dir / LOCAL_OVERRIDE_FILENAME

    if not base_path.exists():
        raise FileNotFoundError(f"Configuration not found: {base_path}")

    with base_path.open("r", encoding="utf-8") as handle:
        config = json.load(handle)

    if local_path.exists():
        with local_path.open("r", encoding="utf-8") as handle:
            local = json.load(handle)
        if not isinstance(local, dict):
            raise ValueError(
                f"{local_path} must contain a JSON object; got {type(local).__name__}"
            )
        config = _deep_merge(config, local)

    return config


__all__ = ["load_config", "CONFIG_FILENAME", "LOCAL_OVERRIDE_FILENAME"]

"""Tests for the shared settings.json + settings.local.json loader."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from config_loader import load_config  # noqa: E402


def _write(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload), encoding="utf-8")


def test_load_config_without_local_returns_base(tmp_path: Path) -> None:
    _write(tmp_path / "config" / "settings.json", {"ibkr": {"host": "127.0.0.1"}})
    assert load_config(tmp_path) == {"ibkr": {"host": "127.0.0.1"}}


def test_load_config_deep_merges_local_over_base(tmp_path: Path) -> None:
    _write(
        tmp_path / "config" / "settings.json",
        {
            "ibkr": {"host": "127.0.0.1", "port": 7497, "expected_account": ""},
            "excel": {"output_mode": "standalone"},
        },
    )
    _write(
        tmp_path / "config" / "settings.local.json",
        {
            "ibkr": {"expected_account": "UACCT123"},
            "excel": {"tax_workbook": "data/output/UACCT123.xlsx"},
        },
    )
    cfg = load_config(tmp_path)
    assert cfg["ibkr"] == {
        "host": "127.0.0.1",
        "port": 7497,
        "expected_account": "UACCT123",
    }
    assert cfg["excel"] == {
        "output_mode": "standalone",
        "tax_workbook": "data/output/UACCT123.xlsx",
    }


def test_load_config_missing_base_raises(tmp_path: Path) -> None:
    with pytest.raises(FileNotFoundError, match="settings.json"):
        load_config(tmp_path)


def test_load_config_local_must_be_object(tmp_path: Path) -> None:
    _write(tmp_path / "config" / "settings.json", {"ibkr": {}})
    (tmp_path / "config" / "settings.local.json").write_text("[1, 2, 3]", encoding="utf-8")
    with pytest.raises(ValueError, match="must contain a JSON object"):
        load_config(tmp_path)

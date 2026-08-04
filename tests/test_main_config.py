"""Tests for main.py config helpers (no TWS required)."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from main import assert_expected_account, resolve_output_path  # noqa: E402


def test_resolve_standalone(tmp_path: Path) -> None:
    path = resolve_output_path(
        {"output_mode": "standalone", "output_file": "data/output/IBKR_Portfolio.xlsx"},
        tmp_path,
    )
    assert path == tmp_path / "data" / "output" / "IBKR_Portfolio.xlsx"


def test_resolve_tax_workbook(tmp_path: Path) -> None:
    path = resolve_output_path(
        {
            "output_mode": "tax_workbook",
            "tax_workbook": "data/output/TRIBUTACAO_WORKING.xlsx",
        },
        tmp_path,
    )
    assert path == tmp_path / "data" / "output" / "TRIBUTACAO_WORKING.xlsx"


def test_resolve_unknown_mode_raises(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="output_mode"):
        resolve_output_path({"output_mode": "nope"}, tmp_path)


def test_assert_expected_account_ok() -> None:
    assert_expected_account(
        {"accounts": [{"account": "UACCT123"}]},
        "UACCT123",
    )


def test_assert_expected_account_blank_skips() -> None:
    assert_expected_account({"accounts": [{"account": "X"}]}, "")


def test_assert_expected_account_mismatch() -> None:
    with pytest.raises(ValueError, match="expected_account"):
        assert_expected_account(
            {"accounts": [{"account": "DUPAPER1"}]},
            "UACCT123",
        )

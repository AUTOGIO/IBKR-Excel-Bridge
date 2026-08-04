"""IBKR Flex Web Service client (stub for Phase 3a).

When ``flex.enabled`` is true and credentials are set, a future slice will
download statements into ``data/statements/`` for the CSV ingest pipeline.
"""

from __future__ import annotations

from pathlib import Path


def download_flex_statement(
    *,
    token: str,
    query_id: str,
    out_dir: Path,
    base_url: str = ("https://ndcdyn.interactivebrokers.com/AccountManagement/FlexWebService"),
) -> Path:
    """Download a Flex query result into ``out_dir``.

    Not implemented in this slice — drop CSVs into ``data/statements/`` instead.
    """
    raise NotImplementedError(
        "Flex Web Service download is stubbed. "
        "Export the Flex/Activity CSV manually into "
        f"{out_dir} and run ./scripts/ingest.zsh. "
        f"(token set={bool(token)}, query_id={query_id!r}, base_url={base_url})"
    )


__all__ = ["download_flex_statement"]

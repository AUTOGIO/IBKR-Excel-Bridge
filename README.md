# IBKR → Excel Bridge

Read-only macOS pipeline: connect to Interactive Brokers (TWS / IB Gateway), snapshot account and portfolio data, and write structured `IBKR_*` Excel sheets. No order placement.

**Run:** `./scripts/setup.zsh` once, then open TWS/Gateway (paper, Read-Only API), then `./scripts/run.zsh`. Drop statement CSVs into `data/statements/` and use `./scripts/ingest.zsh` / `./scripts/promote_events.zsh` when needed.

**Where things live:** `src/` code · `scripts/` runners · `config/` settings · `data/` inputs & Excel output · `docs/` guides · `tests/` tests · `archive/` old files. Day-to-day steps: [`docs/OPERATING_GUIDE.md`](docs/OPERATING_GUIDE.md). Layout rules: [`AGENTS.md`](AGENTS.md).

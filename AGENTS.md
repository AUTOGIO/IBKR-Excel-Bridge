# AGENTS.md — repo layout rules

Keep this personal project simple and predictable. Prefer **move** over copy. Prefer **edit** over new files. Do not redesign features unless asked.

## Allowed top-level folders

| Folder | Contents |
| --- | --- |
| `src/` | Application code (Python). Do not also create `app/`. |
| `scripts/` | Runnable helpers (`.zsh`, `.sh`, `.command`). |
| `config/` | Non-secret settings (e.g. `settings.json`). Never commit secrets. |
| `data/` | CSV, Excel, exports, raw inputs (`data/statements/`, `data/output/`, …). |
| `assets/` | Images, icons, logos (create only when needed). |
| `docs/` | Guides, design notes. AI prompts go in `docs/prompts/`. |
| `tests/` | Tests only. |
| `archive/` | Obsolete files we are not deleting yet. |
| `logs/` | Runtime logs only (gitignored). |

## Root

Root may contain only: `README.md`, `AGENTS.md`, `.gitignore`, and toolchain files
(`requirements.txt`, `*.code-workspace`, etc.).

Do not invent new top-level folders without asking. Do not put personal machine inventory here.
No filename versioning (`Foo_v1.0.md` → `docs/foo.md`; old copy → `archive/` if unsure).

# AGENTS.md

## Setup

```bash
uv sync
```

## Test

```bash
uv run pytest -q
```

## Notes

- Package management uses **uv** — always `uv add` / `uv run`, never pip.
- `.venv` is gitignored, so it is skipped by search tools.

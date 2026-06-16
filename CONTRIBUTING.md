# Contributing

Thank you for considering a contribution to gcal-to-outlook.

## Reporting Issues

Open a GitHub Issue. Include:
- Windows version
- Python version (if running from source)  
- Steps to reproduce
- Relevant lines from sync.log (never include token files or credentials)

## Pull Requests

1. Fork the repository and create a branch from `main`.
2. Keep changes focused: one concern per PR.
3. Write commit messages in English, imperative mood ("Add X", "Fix Y", "Remove Z").
4. Run `pytest` and `ruff check src tests`, then run `python src/app.py once` and confirm it exits without errors before submitting.
5. If you change config structure, update `config.example.json` and the README table.

## Code Style

- Python 3.10+, standard library preferred.
- No commented-out code. No debug prints left in.
- Comments only when the WHY is non-obvious.

## What Not to Include

Never commit: `google_credentials.json`, `google_token.json`, `ms_token_cache.bin`, `config.json`, `sync_state.db`, `*.log`, `*.exe`.
The `.gitignore` already excludes all of these.

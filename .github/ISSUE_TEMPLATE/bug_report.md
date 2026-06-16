---
name: Bug report
about: Report something that does not work as expected
title: "[Bug] "
labels: bug
assignees: ''
---

## Environment

- Windows version (10 or 11, plus build if you know it):
- Install method: installer (`GCalSync-Setup.exe`) or source (`python src/app.py`)
- Python version (only if running from source, `python --version`):
- Microsoft mode: COM (local Outlook) or Graph (Microsoft 365)
- App version (see release tag, or `1.0.0`):

## Steps to reproduce

1.
2.
3.

## Expected behavior

What you expected to happen.

## Actual behavior

What actually happened. Include any error dialog text verbatim.

## Relevant `sync.log` lines

Paste the lines around the failure. The log lives next to the executable
(`%LOCALAPPDATA%\GCalSync\sync.log`) or in the project root when running from source.

```
(paste here)
```

> WARNING: redact secrets before pasting. Strip access tokens, refresh tokens,
> client IDs, and any account credentials. Never attach `google_credentials.json`,
> `google_token.json`, `ms_token_cache.bin`, or `config.json` to an issue.

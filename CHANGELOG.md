# Changelog

All notable changes to this project are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.0.0] - 2026-06-16

First public release.

### Added

- One-way sync from Google Calendar to Microsoft Outlook and Teams. Changes flow
  from Google to Outlook only, so there are no feedback loops.
- Incremental sync using the Google `syncToken`, so each cycle fetches only the
  events that changed since the last run.
- Two Microsoft integration modes: COM mode drives the local Outlook desktop app
  through `pywin32`, and Graph mode calls the Microsoft 365 REST API with MSAL.
- SQLite mapping store that pairs each Google event ID with its Outlook event ID,
  which keeps updates and deletions reliable.
- Tray GUI monitor showing connected accounts, last sync time, and results, with
  a manual "Sync now" trigger, a log viewer, and an autostart checkbox.
- First-run setup wizard that writes `config.json` and handles both logins, so
  end users never edit JSON by hand.
- Windows installer (`GCalSync-Setup.exe`) built with Inno Setup. Per-user install
  to `%LOCALAPPDATA%\GCalSync` with no admin rights, and a native uninstaller in
  "Apps & features".
- Duplicate-event recovery tool for cleaning up events left behind by an
  interrupted or repeated sync.

[Unreleased]: https://github.com/vinicq/gcal-to-outlook/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/vinicq/gcal-to-outlook/releases/tag/v1.0.0

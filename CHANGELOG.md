# Changelog

All notable changes to this project are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.0.4] - 2026-06-17

### Fixed

- Autostart now opens the app to the system tray instead of starting it
  invisibly. Before, "Start with Windows" launched the headless sync loop, which
  synced but showed no window or tray icon, so it looked like nothing started. A
  new `tray` mode shows the tray icon and runs the sync loop in one process.
- Autostart settles on a single mechanism (a launcher in the Windows Startup
  folder). The setup wizard no longer registers a separate scheduled task, and
  enabling autostart removes any leftover task from older installs, so two
  launchers can no longer both fire at login.
- The background sync loop reschedules itself even if a cycle fails, so a single
  failed sync can no longer stop it silently.

### Added

- The window shows the installed version and checks GitHub for a newer release
  on startup. When one exists, a clickable "new version available" notice opens
  the download page. The check is local-only, sends no telemetry, and never
  blocks the app if it fails.

## [1.0.3] - 2026-06-16

### Added

- The window, taskbar, and tray now use the application icon (bundled with the
  build), not the default Python/tk icon.
- The official release installer is turnkey: the OAuth client is written from a
  repository secret at build time and bundled, so a user only logs in with their
  own Google account. The credential is never committed to the repository.
- Troubleshooting section in the README for the Outlook programmatic-access
  prompt and the `invalid_client` OAuth error.

### Changed

- The duplicate-recovery lookup scans the calendar once per run and reads the
  body only of our own `[GCal]` events, instead of a full scan per event. This
  lowers Outlook object-model access. Note: it does not remove the Outlook
  programmatic-access prompt, which is driven by the machine's antivirus status,
  not by this app (see Troubleshooting).
- "Sharing with Others" in the README now documents the secret-based bundling
  and the trade-offs of distributing a Desktop OAuth client.

## [1.0.2] - 2026-06-16

### Fixed

- Duplicate events. After a lost or reset local database, the sync recreated
  events that were already in Outlook. It now looks up its own events by the
  `[sync-id]` marker before creating, so it adopts the existing event instead
  of duplicating. Create new, update only when Google reports a change, delete
  when removed in Google.
- Outlook status showed "Disconnected (CoInitialize...)" in the monitor even
  when sync worked. The status check runs on a worker thread and now
  initializes COM on that thread (CO_E_NOTINITIALIZED).

### Added

- App, installer, and executable icon.
- Monitor: an "Export log" button, a plain "Offline" status when an account is
  unreachable, and a note that Outlook for Windows must be installed.

## [1.0.1] - 2026-06-16

### Fixed

- Setup wizard crashed at the login step in the installed build. It tried to run
  a separate `sync.exe` that the single-executable build never produces. The
  wizard now re-invokes `GCalSync.exe` with the mode argument, and the login
  child runs without grabbing the wizard's console window. Covered by a
  regression test on the command resolver.

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

[Unreleased]: https://github.com/vinicq/gcal-to-outlook/compare/v1.0.4...HEAD
[1.0.4]: https://github.com/vinicq/gcal-to-outlook/compare/v1.0.3...v1.0.4
[1.0.3]: https://github.com/vinicq/gcal-to-outlook/compare/v1.0.2...v1.0.3
[1.0.2]: https://github.com/vinicq/gcal-to-outlook/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/vinicq/gcal-to-outlook/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/vinicq/gcal-to-outlook/releases/tag/v1.0.0

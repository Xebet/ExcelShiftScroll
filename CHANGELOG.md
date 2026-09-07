# Changelog

All notable changes follow [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) principles. This project uses semantic versioning.

## [0.2.1] - 2026-09-07

### Fixed

- Only process `HC_ACTION` wheel notifications; never replay `HC_NOREMOVE` peeks.
- Bind queued COM fallback to its original Excel root window; drop expired (>500 ms), paused, cancelled, or mismatched-window work.
- Reset fractional input on target/settings changes and native delivery to avoid cross-path double counting.
- Restore defaults for missing JSON fields without overriding explicit disabled settings.
- Save settings transactionally with a per-path named mutex and unique temporary files; preserve current settings and show a Ribbon warning on save failure.
- Verify installer checksums before replacing files; serialize installation, retain previous-XLL backups and roll back failed post-install verification.

### Added

- Actual XLL load path and Excel bitness in About, plus an Open add-in folder control.
- Regression tests for hook notification filtering, deferred scrolling, settings failures/concurrency and installer transactions.

The add-in remains entirely offline: no updater, telemetry, or remote code.

## [0.2.0] - 2026-09-03

### Changed

- Shift + wheel now uses Excel's native `WM_MOUSEHWHEEL` path so supported Excel builds provide their normal smooth, pixel-level transition.
- High-resolution wheel and touchpad deltas are forwarded without waiting for a complete 120-unit detent.
- Scroll-distance choices are scaled against the Windows horizontal-wheel setting; exact-column `SmallScroll` remains as a safe compatibility fallback.

### Fixed

- The About dialog now reads the assembly version instead of displaying a stale hard-coded version.

## [0.1.1] - 2026-09-03

### Fixed

- Added a current-user installer that copies the architecture-matched XLL from a release ZIP to `%LocalAppData%\ExcelShiftScroll\AddIn` and removes Mark-of-the-Web from that exact copied file.
- Added a matching binary-only uninstaller and bilingual `INSTALL.txt` with recovery steps for stale temporary-directory add-in entries.
- Strengthened installation documentation for Microsoft 365's default blocking of untrusted XLL files without disabling Office security controls.

## [0.1.0] - 2026-09-03

### Added

- Shift + vertical wheel horizontal scrolling inside Excel worksheet windows.
- Exact one-modifier gating that preserves normal wheel, Ctrl + wheel, Ctrl + Shift + wheel, Alt/Win combinations, and native horizontal wheel input.
- Configurable 1, 2, 3, 5, or 10 columns per detent, direction reversal, pause/resume, and restore-defaults Ribbon controls.
- Fail-closed worksheet hit testing and Excel-main-thread `SmallScroll` dispatch.
- Per-user JSON settings and opt-in privacy-preserving local diagnostics.
- x86 and x64 packed XLL outputs, tests, documentation, packaging, checksums, and GitHub Release automation.

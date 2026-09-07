# Version 0.2.1 verification record

- Date: 2026-09-07.
- Scope: offline reliability fixes; no online updater or network behavior added.

## Local verification

- .NET SDK 8.0.424, locked dependency restore; Release build with zero warnings/errors.
- 49 automated .NET Framework 4.8 tests passed, zero failures/skips.
- 10 isolated installer scenarios passed in both PowerShell 7 and Windows PowerShell 5.1: fresh install, upgrade backup, damaged/missing checksum, locked destination, concurrent installer, Excel already running, Excel starting during staging, post-commit rollback, and failed-first-install cleanup.
- Installer failure cases use controlled mocks for process checks and post-install verification; actual file locking, replacement, backup, and rollback run on temporary files. No installed user XLL is replaced by these tests.
- Three independent Windows PowerShell processes performed 180 settings save/read cycles; JSON remained valid and no temporary files remained.
- Isolated x64 Excel smoke test: `RegisterXLL=true`, version `0.2.1`, status `Mouse hook active`; test process exited cleanly.
- Actual Excel native horizontal-wheel messages: `+120` moved ScrollColumn 10 to 13; `-120` moved 10 to 7.

## Not claimed as passed

- This session's synthetic Shift+wheel end-to-end probe could not obtain foreground focus: Windows retained another foreground HWND. It stopped before injecting wheel/key events. Native wheel movement above passed, but does not substitute for the full gesture test.
- The probe now checks actual foreground HWND and restores the pointer on preparation failure; failed probes also run process-exit verification.
- Perceived smoothness/animation timing, frozen/split panes, mixed-DPI monitors, older Excel versions and actual 32-bit Excel loading still require hands-on testing. Pure policy and queue tests are not claims of full UI coverage.

## Release validation

The Windows tag workflow builds from a clean checkout, runs the automated and installer tests, and packages both x64 and x86 XLLs with SHA-256 files. Publication and downloaded-asset verification are recorded in the release handoff after the workflow finishes.

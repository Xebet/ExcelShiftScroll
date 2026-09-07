# ExcelShiftScroll

[![Build, test, and release](https://github.com/Xebet/ExcelShiftScroll/actions/workflows/build-release.yml/badge.svg)](https://github.com/Xebet/ExcelShiftScroll/actions/workflows/build-release.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

ExcelShiftScroll adds the Windows-standard **Shift + mouse wheel** gesture to desktop Microsoft Excel:

- Shift + wheel up scrolls left.
- Shift + wheel down scrolls right.
- It is enabled by default, uses Excel's native smooth-scrolling path, and defaults to a three-column-equivalent distance per wheel detent.

[简体中文](README.zh-CN.md)

![Demo placeholder: animation will show Shift plus wheel scrolling horizontally in Excel](docs/assets/demo-placeholder.svg)

> A real demonstration GIF will replace this placeholder after the compatibility test matrix is recorded.

## What it preserves

Normal wheel input still scrolls vertically. Ctrl + wheel still zooms. Ctrl + Shift + wheel remains Excel's native behavior. Alt/Windows-key combinations, native horizontal wheel events, non-Excel applications, Ribbon/formula-bar/task-pane/dialog/VBA-editor regions, and background Excel windows are not converted.

The add-in has no external process, no network access, no telemetry, and no workbook-content access.

## Supported systems

Primary target: Windows 11, 64-bit Microsoft 365 desktop Excel. The build also produces a 32-bit XLL and is designed for Windows 10 and Excel 2019, 2021, and 2024. Excel for the web, macOS, and mobile are not supported. See [compatibility notes](docs/compatibility.md) for verified and pending combinations.

## Download

Download the ZIP matching your Excel bitness from [GitHub Releases](https://github.com/Xebet/ExcelShiftScroll/releases/latest), then verify it against `SHA256SUMS.txt`.

To check Excel bitness: open **File → Account → About Excel**. The first line says either 32-bit or 64-bit. Windows bitness is not a substitute; a 64-bit Windows installation can run 32-bit Excel.

## Install

1. Exit every Excel process. Do not load the XLL from inside a ZIP or `%TEMP%`.
2. Right-click the downloaded ZIP, choose **Properties**, select **Unblock** if shown, and then extract it.
3. In the extracted folder, right-click `Install-CurrentUser.ps1` and choose **Run with PowerShell**. It verifies the included XLL checksum, stages and unblocks the file, then installs under `%LocalAppData%\ExcelShiftScroll\AddIn`. It retains the previous XLL as a timestamped `.bak` and restores it if post-install verification fails. It does not change the registry, Trust Center, or Excel's add-in list.
4. Start Excel and open **File → Options → Add-ins**.
5. At the bottom, select **Excel Add-ins**, choose **Go**, then **Browse**. This is not the Office Store **My Add-ins** page.
6. Select `ExcelShiftScroll64.xll` for 64-bit Excel or `ExcelShiftScroll32.xll` for 32-bit Excel from the stable installed folder.
7. Place the pointer over the worksheet grid, hold Shift, and turn the vertical wheel.

Manual alternative: extract to a stable local folder, right-click the XLL itself, choose **Properties → General → Unblock → OK**, and browse to that exact file. See the packaged `INSTALL.txt`. [Microsoft documents that current Excel blocks XLL files from untrusted locations by default](https://support.microsoft.com/en-US/Excel/excel-is-blocking-untrusted-xll-add-ins-by-default).

No administrator rights or separate .NET runtime installation is required. The release is currently unsigned, so Windows SmartScreen or Office may show a publisher warning. Do not disable security features; use the repository Release page, unblock only the file you verified, and compare its SHA-256.

### Upgrade or restore a previous version

Close Excel and run the new package's installer. If you already load from the stable installed path, reopen Excel; otherwise uncheck the old XLL entry and browse to the installed path once. **About** shows the actual loaded XLL path and version: a directory name is not reliable version evidence. An Excel add-in entry and a COM Ribbon helper with the same name can be normal; do not delete the helper merely because the name is duplicated.

To restore a backup, close Excel, preserve the current XLL separately, copy the desired timestamped `.bak` in the installed folder back to its original `ExcelShiftScroll64.xll` or `ExcelShiftScroll32.xll` name, and reopen Excel. Settings are preserved. Backups are not automatically pruned. SHA-256 detects corruption, not publisher authenticity; binaries remain unsigned. No online update feature is included.

## Controls

Open Excel's **Add-ins** Ribbon tab and use the **Shift Scroll** group:

- enable or pause conversion immediately;
- choose a 1, 2, 3, 5, or 10-column-equivalent distance per detent;
- reverse direction;
- restore defaults;
- show version, runtime status, Excel bitness, and the actual loaded XLL path;
- open the loaded add-in's folder.

Settings are per-user at `%LocalAppData%\ExcelShiftScroll\settings.json`; they are never stored in a workbook.

Missing JSON fields use defaults while explicit `false` values are preserved. Failed saves leave the current settings unchanged and show a warning. Separate Excel processes serialize file writes; the last successful whole-settings save wins, and existing processes do not live-reload one another's settings.

## Uninstall

1. In Excel, open **File → Options → Add-ins**.
2. Select **Excel Add-ins**, choose **Go**, and clear ExcelShiftScroll.
3. Exit every Excel process. The in-process hook is removed during add-in shutdown; there is no background process.
4. Run the packaged `Uninstall-CurrentUser.ps1` to remove the installed XLLs. It preserves settings and backups. Delete old extracted copies separately if you no longer need them. Optionally delete `%LocalAppData%\ExcelShiftScroll` to remove settings, backups, and opt-in diagnostics.

## Troubleshooting

- **“Not a valid add-in”** usually means XLL bitness does not match Excel. Recheck **About Excel**.
- **Excel says the source is untrusted**: an XLL loaded from a browser-download or temporary folder probably carries Mark-of-the-Web. Exit Excel, run the packaged current-user installer or use **Properties → Unblock** on the stable local XLL, then browse to it again. Do not disable `BlockXLLFromInternet` or weaken Trust Center globally.
- **A stale `%TEMP%` entry remains**: try checking the missing entry once; if Excel offers to delete it from the list, choose **Yes**, then browse to the stable installed copy.
- **Ribbon group missing**: check **File → Options → Add-ins → Disabled Items**. Ribbon callback failures can cause Office to disable a COM helper.
- **No horizontal scroll**: ensure the pointer is over the worksheet grid, only Shift is pressed, the add-in is enabled, and the sheet has horizontally scrollable columns.
- **Scrolling is not smooth**: ExcelShiftScroll forwards precision-preserving horizontal wheel deltas to Excel's native scrolling engine. Zero/page horizontal-wheel settings or a failed native post use the exact-column fallback. Older Excel builds may render native messages without smooth transitions; the add-in does not detect animation support or implement its own easing.
- **No action over formula bar/Ribbon/dialog/task pane** is intentional.
- **Corporate policy blocks XLL files**: ask the administrator to approve the verified file or deploy a signed build. This project does not bypass policy.
- **Need diagnostics**: edit `diagnosticsEnabled` to `true` in the settings file while Excel is closed. Logs contain only timestamps, fixed event names, and exception types. Re-disable it after diagnosis.

If Excel disables the add-in after a crash, preserve the Windows/Excel version, XLL bitness, and sanitized diagnostics, then open an issue. Never upload a sensitive workbook.

## Build from source

Requirements: Windows, .NET 8 SDK (used as the build SDK), and PowerShell. The add-in itself targets .NET Framework 4.8, which ships with supported Windows versions.

```powershell
dotnet restore ExcelShiftScroll.sln --configfile NuGet.Config
dotnet build ExcelShiftScroll.sln --configuration Release --no-restore
dotnet test ExcelShiftScroll.sln --configuration Release --no-build --no-restore
./tests/install/Test-Installer.ps1
./build/Package-Release.ps1 -Version 0.2.1
```

Packed XLLs are produced under `src/ExcelShiftScroll/bin/Release/net48/publish`. Release ZIPs and hashes are produced under `artifacts/release`. CI repeats these steps on a Windows runner and publishes tag builds matching `v*`.

See [architecture](docs/architecture.md), [testing](docs/testing.md), and the [ADRs](docs/adr/) for design evidence.

## Privacy and security

ExcelShiftScroll is offline. It does not read workbook names, paths, formulas, values, selections, or VBA; record keyboard characters, pointer locations, or wheel history; send telemetry; download or execute code; change Trust Center; or install a service. The thread-scoped mouse hook exists only inside the hosting Excel process and is unhooked on shutdown. See [SECURITY.md](SECURITY.md).

## Acknowledgements and clean-room development

[OfficeScroll](https://github.com/T800G/OfficeScroll) is acknowledged as a related historical project and a user-experience/compatibility reference. ExcelShiftScroll is an independent clean-room implementation. No OfficeScroll CPOL source, binary, resource, translation, decompilation output, or line-by-line rewrite was used. See the [clean-room record](docs/clean-room.md).

## Dependencies and licenses

The runtime uses Excel-DNA 1.9.0 under the zlib license. Build/test dependencies and notices are listed in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md). ExcelShiftScroll source is licensed under [MIT](LICENSE).

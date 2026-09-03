# Version 0.1.0 verification record

- Date: 2026-09-03
- Host-reported Windows registry product: Windows 10 Home China, display version 25H2, build 26200.9168 (the legacy `ProductName` field can differ from marketed Windows naming)
- Excel executable: Microsoft 365 desktop Excel 64-bit, 16.0.20326.20112
- Build SDK: isolated .NET SDK 8.0.424
- Target/runtime: .NET Framework 4.8, Excel-DNA 1.9.0

## Passed

- `dotnet restore` with committed lock files and the repository NuGet configuration.
- Release solution build: 0 warnings, 0 errors.
- Packed Excel-DNA outputs: x86 and x64 XLL created.
- Automated tests: 22 passed, 0 failed, 0 skipped.
- `dotnet format --verify-no-changes`: passed.
- Development `WindowProbe` build: 0 warnings, 0 errors.
- x64 isolated Excel smoke test: `RegisterXLL=true`; `ExcelShiftScroll.Version()` returned `0.1.0`; `ExcelShiftScroll.Status()` returned `Mouse hook active`.
- Smoke-test Excel process exited within 10 seconds; no newly created Excel process remained. A pre-existing Excel PID was explicitly excluded and was neither modified nor closed.
- Both release ZIPs contain the architecture-matched XLL, English and Chinese READMEs, MIT license, third-party notices, and internal XLL SHA-256.
- ZIP SHA-256 verification passed:
  - `f6237a2211c24bb59ae6befa95e140f5709436feea0019f7b15abf6121f25b2a  ExcelShiftScroll-v0.1.0-x64.zip`
  - `7920ab407aca071c0c06c96d53c4de7b104293bc70ea5e9685db448f3bff8509  ExcelShiftScroll-v0.1.0-x86.zip`

## Not performed in this automated session

- Physical Shift + wheel gestures and proof of direction/column distance.
- Ribbon interaction by a human operator.
- Frozen panes, split panes, multiple workbook windows, large sheets, formula editing, modal dialogs, menus, Ribbon/formula-bar/name-box/task-pane/VBA-editor pointer surfaces.
- Multi-monitor mixed-DPI behavior.
- High-resolution wheel and precision touchpad hardware.
- 32-bit Excel XLL loading (no 32-bit Excel host was available).
- Windows 10 and Excel 2019/2021/2024 host matrices.

These remain unchecked in `docs/testing.md`; no pass is claimed for them.

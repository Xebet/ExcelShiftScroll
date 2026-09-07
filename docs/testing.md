# Testing

The latest maintenance verification is recorded in [v0.2.1 results](test-results-0.2.1.md).
Run `tests/install/Test-Installer.ps1` under both PowerShell 7 and Windows PowerShell 5.1, and `tests/settings/Test-ConcurrentSettings.ps1` after a Release build. These tests use isolated temporary folders, not the current user's installed add-in or settings. The native Excel integration probe requires PowerShell 7 and an interactive desktop that permits foreground focus.

## Automated tests

Run on Windows:

```powershell
dotnet restore ExcelShiftScroll.sln --configfile NuGet.Config
dotnet build ExcelShiftScroll.sln --configuration Release --no-restore
dotnet test ExcelShiftScroll.sln --configuration Release --no-build --no-restore
```

The suite covers:

- Shift + wheel-up → left and wheel-down → right;
- reverse direction and allowed column counts;
- no Shift, Ctrl + Shift, Alt + Shift, and Win + Shift pass-through;
- background Excel and non-worksheet pass-through;
- native horizontal-wheel pass-through;
- disabled setting pass-through;
- precision-delta accumulation;
- native horizontal-message direction, signed multi-monitor coordinates, cleared modifier state, distance scaling, retained sub-unit deltas, and COM fallback selection;
- corrupt settings fallback, validation, and persistence;
- repeated hook initialization/release and failed installation;
- queued-scroll coalescing, shutdown dropping, and rejection after disposal.

CI also verifies a warning-free Release build, both packed XLL architectures, installer/uninstaller and instruction inclusion, ZIP contents, and SHA-256 generation.

The Windows-only integration check opens an isolated temporary workbook, first posts signed `WM_MOUSEHWHEEL` messages to its `EXCEL7` pane, then briefly synthesizes Shift + vertical-wheel input over that pane to exercise the complete hook-to-native-scroll path. It verifies `ScrollColumn` in both directions, restores the pointer, releases Shift in a `finally` block, saves no document, and closes only the Excel process it created:

```powershell
./tests/manual/Test-NativeHorizontalWheel.ps1 -XllPath ./src/ExcelShiftScroll/bin/Release/net48/publish/ExcelShiftScroll-AddIn64-packed.xll
```

Evidence snapshots: [0.1.0](test-results-0.1.0.md), [0.1.1 trust-installation fix](test-results-0.1.1.md), and [0.2.0 native smooth scrolling](test-results-0.2.0.md).

## Development window probe

`tools/WindowProbe` enumerates only handle, process/thread id, class name, and ancestry for Excel windows. It never reads titles or workbook data. Use it to validate window hierarchy before changing the fail-closed allowlist:

```powershell
dotnet run --project tools/WindowProbe/WindowProbe.csproj
```

## Manual Excel validation checklist

Record Windows version, Excel exact version and bitness, display configuration, input device, XLL hash, outcome, and notes. Never mark an item passed without direct observation.

### Installation and lifecycle

- [ ] A clean 64-bit Microsoft 365 Excel session loads `ExcelShiftScroll64.xll`.
- [ ] The Add-ins Ribbon group appears and About reports **Mouse hook active**.
- [ ] Disabling/re-enabling via Excel Add-ins does not require a Windows restart.
- [ ] Normal Excel shutdown completes without an error dialog.
- [ ] Task Manager shows no new residual Excel or helper process after all Excel windows close.
- [ ] A wrong-bitness XLL fails without destabilizing Excel.

### Core input behavior

- [ ] Blank workbook: Shift + wheel-up moves left; Shift + wheel-down moves right.
- [ ] Default movement is approximately three columns per detent and uses Excel's smooth transition.
- [ ] 1, 2, 3, 5, and 10-column-equivalent settings change native scroll distance proportionally.
- [ ] High-resolution deltas smaller than 120 produce incremental pixel movement rather than waiting for a full detent.
- [ ] Windows page-at-a-time horizontal-wheel configuration uses the exact-column compatibility fallback.
- [ ] Reverse direction flips both directions.
- [ ] Pause passes input through; resume restores handling immediately.
- [ ] Normal wheel remains vertical.
- [ ] Ctrl + wheel remains Excel zoom.
- [ ] Ctrl + Shift + wheel remains native and is not doubled.
- [ ] Alt + Shift and Win + Shift do not trigger horizontal scrolling.
- [ ] A native horizontal wheel/trackpad gesture is not doubled.

### Excel states and surfaces

- [ ] Large worksheet and rapid continuous scrolling remain responsive.
- [ ] Frozen panes.
- [ ] Split panes.
- [ ] Multiple workbooks and multiple Excel windows.
- [ ] Cell edit mode, including formula entry.
- [ ] Modal Excel dialog and dropdown menu.
- [ ] Ribbon, Quick Access Toolbar, name box, formula bar, status bar, scroll bars, task pane.
- [ ] VBA editor.
- [ ] Excel minimized and Excel visible but behind another application.

### Displays and devices

- [ ] Single monitor at 100% scale.
- [ ] Multiple monitors with equal DPI.
- [ ] Multiple monitors with different DPI values, including negative virtual-screen coordinates.
- [ ] Standard detented mouse wheel.
- [ ] High-resolution wheel.
- [ ] Precision touchpad vertical scrolling while Shift is held.

## Test evidence policy

Automated tests do not prove real Excel UI behavior. Release notes must distinguish: build/test results; temporary automated XLL load; and hands-on wheel/device tests. Unsupported or unexecuted rows remain explicitly pending.

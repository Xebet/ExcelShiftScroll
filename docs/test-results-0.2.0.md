# Version 0.2.0 verification record

- Date: 2026-09-03
- Purpose: replace abrupt whole-column COM scrolling with Excel's native smooth horizontal-wheel path while retaining a safe compatibility fallback.
- Environment: Windows 11 x64 build 26200; Microsoft 365 Excel x64 file version 16.0.20326.20112.

## Passed

- Release solution build: 0 warnings, 0 errors.
- Automated tests: 28 passed, 0 failed, 0 skipped.
- New native-dispatch coverage verifies direction, cleared modifier flags, signed multi-monitor coordinates, configured-distance scaling, retained sub-unit precision deltas, page/zero-setting fallback, and post failure handling.
- x64 isolated Microsoft 365 Excel smoke test: `RegisterXLL=true`; version `0.2.0`; hook status `Mouse hook active`.
- The isolated XLL smoke-test process exited cleanly; no pre-existing Excel process was modified or closed.
- Native Excel integration probe: an ordinary `WM_MOUSEHWHEEL` delta of `+120` moved `ScrollColumn` from 10 to 13, and `-120` moved it from 10 to 7.
- The probe caught and prevented an Excel-specific incompatibility: including `MK_SHIFT` in the converted horizontal message snapped the view to the first column, so the release intentionally clears modifier flags after Shift has triggered conversion.
- Full input-path probe with the v0.2.0 XLL loaded: synthesized Shift + wheel-up moved `ScrollColumn` from 10 to 7; Shift + wheel-down moved it from 10 to 13. The script restored the pointer and force-released Shift in a `finally` block.
- The native integration-test process exited cleanly and did not retain a workbook or Excel process.
- Both x64 and x86 packages passed required-entry, architecture, XLL checksum, and outer ZIP checksum verification.

## Published release checksums

The assets were downloaded again from GitHub after the tag workflow completed. Their hashes match the `SHA256SUMS.txt` uploaded by that workflow:

- `7ca7daa393d248e3de3851abc88f556ec4a25602fb630b20b18f33680b7edec4  ExcelShiftScroll-v0.2.0-x64.zip`
- `c06b864962debe91d2000f01a888d120e13afcf05e2576b5feab5053272571f1  ExcelShiftScroll-v0.2.0-x86.zip`
- `67b30c722ce81a4d7e1e461012053948e5b95e7c22c34c58af7e60e34e38453b  ExcelShiftScroll64.xll`
- `c10f4a0c33c40d3fe363de1553d74de85f4b0c40bf787c486bee4899c5ec12df  ExcelShiftScroll32.xll`

ZIP bytes are not reproducible between the local and CI packaging environments because archive metadata differs, so only the published workflow checksum file is authoritative for downloaded assets.

## Still requiring hands-on judgment

Automation proves that Excel's native horizontal-wheel engine receives the correct precision-preserving messages and moves in both directions. Perceived animation quality must still be judged with the user's physical wheel or touchpad because a COM property sample cannot measure rendered easing. Frozen panes, split panes, mixed-DPI monitors, 32-bit Excel, and older Excel releases remain in the manual compatibility matrix.

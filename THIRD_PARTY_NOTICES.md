# Third-party notices

ExcelShiftScroll itself is licensed under the MIT License.

## Runtime dependency

- **Excel-DNA 1.9.0** — Copyright Excel-DNA contributors; distributed under the zlib license. Source and license: <https://github.com/Excel-DNA/ExcelDna>.

The packed `.xll` release files include the Excel-DNA loader and the managed ExcelShiftScroll assembly. The `Microsoft.NETFramework.ReferenceAssemblies` package is a build-only dependency and is not redistributed in the packed add-in.

## Test-only dependencies

- Microsoft.NET.Test.Sdk 17.11.1 — MIT License.
- xUnit.net 2.9.2 — Apache License 2.0.
- xunit.runner.visualstudio 2.8.2 — Apache License 2.0.

These test packages are not included in release ZIP files.

## OfficeScroll clean-room statement

[OfficeScroll](https://github.com/T800G/OfficeScroll) was reviewed only as a historical user-experience and compatibility reference. No OfficeScroll source code, binaries, resources, translations, decompilation output, or line-by-line reimplementation are included. ExcelShiftScroll is an independent clean-room implementation and is not derived from OfficeScroll's CPOL-licensed code.

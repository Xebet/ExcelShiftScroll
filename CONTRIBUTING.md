# Contributing

Thank you for helping improve ExcelShiftScroll.

1. Open an issue describing the bug or proposed behavior.
2. Fork the repository and create a focused branch.
3. Keep input policy separate from Win32 and Excel COM code.
4. Add or update tests for behavior changes.
5. Run `dotnet restore ExcelShiftScroll.sln --configfile NuGet.Config`, `dotnet build ExcelShiftScroll.sln -c Release --no-restore`, and `dotnet test ExcelShiftScroll.sln -c Release --no-build --no-restore` on Windows.
6. Complete relevant checks in `docs/testing.md`; never claim manual Excel coverage that was not performed.
7. Submit a pull request with the motivation, test evidence, Excel version/bitness, and any compatibility risk.

Do not submit secrets, certificates, workbook data, personal paths, generated `bin`/`obj` content, or code copied from OfficeScroll. Security reports should follow `SECURITY.md` instead of public issue discussion.

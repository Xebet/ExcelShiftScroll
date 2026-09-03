# Security policy

## Supported versions

Security fixes are provided for the latest published release.

## Reporting a vulnerability

Use GitHub's private vulnerability reporting feature for this repository. Do not attach sensitive workbooks, credentials, certificates, or personal paths. Include the Excel version, Windows version, add-in bitness, reproduction steps, and impact. Maintainers will acknowledge a report when reviewed and coordinate disclosure when appropriate.

## Mouse-hook scope and privacy

ExcelShiftScroll installs a `WH_MOUSE` hook only on the Excel UI thread that loads the add-in. It is not a desktop-global hook and is removed in `AutoClose`. The callback considers only vertical wheel messages, Shift/Ctrl/Alt/Windows modifier state, foreground-window process identity, the window ancestry under the pointer, and wheel delta. It does not record character keys, pointer coordinates, or wheel history.

The add-in:

- makes no network connections;
- contains no telemetry, advertising, analytics, updater, or remote-code loader;
- does not read workbook names, paths, cell values, formulas, or VBA projects;
- does not modify workbooks, Trust Center, macro security, or Office policy;
- does not require administrator rights or install a background process.

Optional local diagnostics are disabled by default. When enabled, JSON-lines records contain only UTC time, a fixed event name, and an exception type. They are stored under `%LocalAppData%\ExcelShiftScroll\diagnostics.jsonl` and never include exception messages, workbook data, file paths, pointer coordinates, or input history.

## Signing

Official 0.1.0 artifacts are not Authenticode-signed unless a release explicitly says otherwise. The project does not bypass SmartScreen or Office security warnings. Verify `SHA256SUMS.txt` and obtain releases only from this repository.

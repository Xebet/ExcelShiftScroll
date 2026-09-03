# Version 0.1.1 verification record

- Date: 2026-09-03
- Purpose: correct the Microsoft 365 untrusted-XLL installation failure caused by browser Mark-of-the-Web and temporary-folder loading.

## Passed

- Release solution build: 0 warnings, 0 errors.
- Automated tests: 22 passed, 0 failed, 0 skipped.
- x64 isolated Microsoft 365 Excel smoke test: `RegisterXLL=true`; version `0.1.1`; hook status `Mouse hook active`.
- The isolated Excel process exited cleanly within the 30-second teardown threshold; a pre-existing user Excel PID was neither modified nor closed.
- Mark-of-the-Web test fixture: a `Zone.Identifier` stream with `ZoneId=3` was attached to a copied XLL; `Unblock-File` removed the stream from that exact file (`zoneBefore=True`, `zoneAfter=False`). The temporary fixture was then deleted.
- Both x64 and x86 ZIPs contain `INSTALL.txt`, `Install-CurrentUser.ps1`, `Uninstall-CurrentUser.ps1`, the architecture-matched XLL, both READMEs, licenses, and internal checksums.
- Package verification passed for ZIP hashes, XLL hashes, required-entry names, and architecture mapping.
- Local ZIP SHA-256:
  - `d8033917d203ccba18af303cf0c631e004ccce2a3d29d8235cd9f426c5371941  ExcelShiftScroll-v0.1.1-x64.zip`
  - `30cd366b12e3f312d3d1de845c9bf14e9eab2132ef07f035bdc4a79ad7d58f3a  ExcelShiftScroll-v0.1.1-x86.zip`

## Safety properties of the installer

- Refuses to run while any Excel process is active.
- Requires exactly one matching XLL beside the script.
- Copies only that XLL to `%LocalAppData%\ExcelShiftScroll\AddIn`.
- Removes Mark-of-the-Web only from the copied XLL and verifies the stream is absent.
- Does not edit the registry, Trust Center, trusted locations, Office policies, or Excel's add-in list.

## Still requiring the user

After fully closing Excel, the user must run the installer (or manually unblock the stable XLL) and browse to the installed path from **Manage: Excel Add-ins → Go**. Persistent registration remains an explicit Excel UI action; the project does not silently alter Office configuration.

# Clean-room development record

ExcelShiftScroll was designed from the public behavior requested for Shift + wheel and from official Microsoft Win32/Excel-DNA documentation.

OfficeScroll was considered only as evidence that users value Office-wide horizontal-wheel behavior and may encounter compatibility differences. During ExcelShiftScroll development:

- no OfficeScroll source file was copied, translated, opened for implementation study, decompiled, rewritten, or added;
- no OfficeScroll binary, image, icon, text resource, or other asset was included;
- no claim is made that ExcelShiftScroll is a fork or derivative;
- the input architecture was independently derived from documented `WH_MOUSE`, foreground/hit testing, Excel-DNA `QueueAsMacro`, and Excel `Window.SmallScroll` behavior;
- all new source in this repository is offered under MIT, while Excel-DNA retains its own zlib license.

If future work requires reusing CPOL-covered implementation material, development must stop, the license impact must be reviewed publicly, and the material must not enter the MIT repository without explicit authorization and compatible licensing.

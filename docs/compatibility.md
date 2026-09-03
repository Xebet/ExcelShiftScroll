# Compatibility

## Intended support

| Platform | Status | Notes |
|---|---|---|
| Windows 11 + Microsoft 365 Excel x64 | Primary | Build and temporary load validation environment. Full gesture matrix still requires hands-on input testing. |
| Windows 10 + Excel 2019/2021/2024 x64 | Designed | Uses .NET Framework 4.8 and long-standing Win32/Excel APIs; manual matrix pending. |
| 32-bit Excel | Build produced | Packed x86 XLL is built and checksum-tested; loading requires a 32-bit Excel host. |
| Multiple workbooks / modern SDI Excel windows | Designed | Thread-scoped hook plus per-event foreground/root checks. |
| Frozen and split panes | Designed | `SmallScroll` acts on the active Excel window; manual behavior validation pending. |
| Mixed-DPI multi-monitor setups | Designed | No custom coordinate scaling; class hit testing uses the Win32 screen point. Manual matrix pending. |
| Precision wheels and touchpads | Partial | Fractional vertical-wheel deltas accumulate. Native horizontal-wheel messages pass through. Device-specific testing pending. |
| Excel for web, macOS, mobile | Unsupported | Requires Win32 and Excel-DNA. |

## Fail-closed worksheet detection

The pointer must resolve into an `EXCEL7` ancestor in the active Excel foreground root. This rejects Ribbon, formula bar, name box, dialogs, menus, task panes, and the VBA editor on known versions. Unknown future worksheet class names are rejected instead of guessed. If a supported Excel update changes its window hierarchy, open an issue with class-only output from `tools/WindowProbe`; do not include workbook titles.

## Coexistence

- Normal wheel, Ctrl + wheel, Ctrl + Shift + wheel, Alt/Win combinations, and native horizontal input are passed to the remaining hook chain.
- The hook is scoped to the Excel UI thread and calls `CallNextHookEx` for every event it does not intentionally consume.
- Other applications and other Excel processes are not hooked.
- There is no resident helper process or driver.

## Runtime choice

The add-in targets .NET Framework 4.8 so supported Windows installations need no extra .NET Desktop Runtime. It uses Excel-DNA 1.9.0, whose current stable release supports .NET Framework 4.7.2+ and modern .NET. See [ADR 0001](adr/0001-target-net48.md).

# ADR 0002: Thread-scoped mouse hook plus queued COM scrolling

- Status: Accepted
- Date: 2026-09-03

## Technical spike

A minimal spike compared the three viable capture/scroll families against scope, safety, exact column movement, Excel busy-state behavior, and teardown:

| Option | Strengths | Risks | Decision |
|---|---|---|---|
| Desktop-global `WH_MOUSE_LL` | Sees raw wheel input before dispatch; easy cross-window capture | Observes the whole desktop, must gate every event, greater blast radius, global-hook timeout/removal behavior | Rejected |
| Worksheet subclassing/message replacement | Direct per-window message access | Excel recreates child windows; version-specific lifecycle; subclass mistakes can crash Excel; many panes/windows to track | Rejected |
| Excel UI-thread `WH_MOUSE` | Same-process/thread scope; no DLL injection; receives wheel before dispatch; one teardown handle | Still requires strict surface hit testing; managed delegate must remain rooted | Selected |
| Convert to `WM_MOUSEHWHEEL` | Avoids COM and naturally supports busy Excel | Column distance follows Excel/system settings and is not the requested exact configurable count | Rejected as primary/fallback only for future study |
| Queue `Window.SmallScroll` | Exact column count; works with active workbook/window; no synthetic-event recursion | COM cannot run inside hook; a gesture may be dropped while Excel shuts down | Selected through `QueueAsMacro` |

The development-only `tools/WindowProbe` supports class-hierarchy validation without titles or workbook access. The production code keeps Win32, policy, and COM layers independently testable.

## Decision

Install `WH_MOUSE` with `dwThreadId = GetCurrentThreadId()` during `AutoOpen` and `hMod = NULL`, as allowed for hook code in the current process. Hold the delegate strongly and always unhook during `AutoClose` or failed initialization.

For `WM_MOUSEWHEEL`, sample only modifier state, wheel delta, foreground ownership, and hit-window ancestry. Require Shift alone and an `EXCEL7` ancestor in the active Excel root. Queue signed column deltas through `ExcelAsyncUtil.QueueAsMacro`; call `ActiveWindow.SmallScroll` only in that safe macro context.

References:

- Microsoft `SetWindowsHookEx`: <https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-setwindowshookexw>
- Excel-DNA async work: <https://excel-dna.net/docs/guides-advanced/performing-asynchronous-work/>
- Excel-DNA COM guidance: <https://excel-dna.net/docs/guides-basic/excel-programming-interfaces/using-the-excel-com-automation-interfaces/>

## Failure and fallback policy

- Any callback exception calls the next hook and does not log to disk.
- Unknown class names pass through; there is no broad class-name fallback.
- Queue/COM failures are caught and may drop a gesture, but never retry with a synthetic event or busy wait.
- Native `WM_MOUSEHWHEEL` always passes through, preventing duplicate horizontal movement.
- The system can adopt a verified new worksheet class in a future release only with probe evidence and manual surface regression tests.

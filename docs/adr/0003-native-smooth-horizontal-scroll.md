# ADR 0003: Native horizontal-wheel dispatch for smooth scrolling

- Status: Accepted
- Date: 2026-09-03

## Context

The original `Window.SmallScroll` output is reliable and gives an exact column count, but each call snaps the viewport immediately. Current Excel versions have an improved scrolling engine that can retain partial row or column positions and accept precision-device deltas. The add-in should use that engine instead of attempting to recreate its animation with timers and repeated COM calls.

## Decision

Keep the existing thread-scoped `WH_MOUSE` input hook and fail-closed worksheet hit test. For an eligible Shift + `WM_MOUSEWHEEL` event:

1. preserve the original delta, including values smaller than `WHEEL_DELTA`;
2. reverse its sign because positive vertical means up while positive horizontal means right;
3. scale it against `SPI_GETWHEELSCROLLCHARS` so the Ribbon distance remains meaningful;
4. asynchronously post an ordinary `WM_MOUSEHWHEEL` to the hit worksheet window with signed screen coordinates, clearing the Shift message flag because Shift was only the source-gesture trigger;
5. let that posted horizontal message pass through the hook unchanged.

Posting is constant-time and avoids calling Excel or another window procedure from the hook. It also lets Excel control pixel movement, easing, frozen/split pane behavior, and version-specific rendering. No timer, animation loop, worker thread, synthetic global input, or retained COM object is introduced.

An integration probe against Microsoft 365 Excel confirmed that an unmodified `WM_MOUSEHWHEEL` moves three columns per 120-unit message in the documented direction. Forwarding `MK_SHIFT` caused Excel-specific special behavior that snapped the view to the first column, so the converted message intentionally contains no modifier flags.

If Windows is configured for zero characters or page-at-a-time horizontal scrolling, the native distance cannot be mapped safely. If `SystemParametersInfo` or `PostMessage` fails, use the previous queued `ActiveWindow.SmallScroll` path for that complete detent. This fallback preserves functionality, although it is intentionally not smooth.

## Trade-offs

- The selected distance is column-equivalent rather than an absolute number of rendered columns on sheets with mixed column widths.
- Smoothness depends on the Excel build and Windows mouse settings because Excel owns the animation.
- The COM fallback remains necessary for unusual system settings and native-message failures.

## References

- Microsoft `WM_MOUSEHWHEEL`: <https://learn.microsoft.com/windows/win32/inputdev/wm-mousehwheel>
- Microsoft Excel improved scrolling: <https://support.microsoft.com/en-US/Excel/move-or-scroll-through-a-worksheet>

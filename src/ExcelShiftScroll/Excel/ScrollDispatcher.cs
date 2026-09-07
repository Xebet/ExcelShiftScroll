using System;
using System.Diagnostics;
using ExcelShiftScroll.Diagnostics;

namespace ExcelShiftScroll.Excel;

public sealed class ScrollDispatcher : IDisposable
{
    private readonly object _gate = new();
    private readonly IExcelMacroQueue _macroQueue;
    private readonly IExcelScroller _scroller;
    private readonly IDiagnosticLog _log;
    private readonly Func<bool> _isEnabled;
    private readonly Func<long> _milliseconds;
    private int _pendingColumns;
    private IntPtr _target;
    private long _started;
    private bool _scheduled;
    private bool _disposed;

    public ScrollDispatcher(IExcelMacroQueue macroQueue, IExcelScroller scroller,
        IDiagnosticLog? log = null, Func<bool>? isEnabled = null, Func<long>? milliseconds = null)
    {
        _macroQueue = macroQueue ?? throw new ArgumentNullException(nameof(macroQueue));
        _scroller = scroller ?? throw new ArgumentNullException(nameof(scroller));
        _log = log ?? NullDiagnosticLog.Instance;
        _isEnabled = isEnabled ?? (() => true);
        _milliseconds = milliseconds ?? (() => (long)(Stopwatch.GetTimestamp() * 1000.0 / Stopwatch.Frequency));
    }

    public bool Enqueue(int columnDelta, IntPtr target)
    {
        lock (_gate)
        {
            if (_disposed || target == IntPtr.Zero) { return false; }
            if (columnDelta == 0) { return true; }
            var now = _milliseconds();
            if (_pendingColumns == 0 || _target != target || now - _started > 500)
            {
                _pendingColumns = 0;
                _target = target;
                _started = now;
            }
            _pendingColumns = (int)Math.Max(-1000L, Math.Min(1000L, (long)_pendingColumns + columnDelta));
            if (_scheduled) { return true; }
            _scheduled = true;
            try
            {
                _macroQueue.Queue(DrainOnExcelThread);
                return true;
            }
            catch
            {
                _scheduled = false;
                _pendingColumns = 0;
                // Enqueue runs inside the input hook: no disk logging here.
                return false;
            }
        }
    }

    public void CancelPending()
    {
        lock (_gate) { _pendingColumns = 0; }
    }

    public void Dispose()
    {
        lock (_gate) { _disposed = true; _pendingColumns = 0; }
    }

    private void DrainOnExcelThread()
    {
        lock (_gate)
        {
            var columns = _pendingColumns;
            _pendingColumns = 0;
            _scheduled = false;
            try
            {
                if (!_disposed && columns != 0 && _milliseconds() - _started <= 500 && _isEnabled())
                {
                    _scroller.ScrollColumns(columns, _target);
                }
            }
            catch (Exception exception) { _log.Write("scroll_execute_failed", exception); }
        }
    }
}

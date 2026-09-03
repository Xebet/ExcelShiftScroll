using System;
using System.Threading;
using ExcelShiftScroll.Diagnostics;

namespace ExcelShiftScroll.Excel;

public sealed class ScrollDispatcher : IDisposable
{
    private readonly IExcelMacroQueue _macroQueue;
    private readonly IExcelScroller _scroller;
    private readonly IDiagnosticLog _log;
    private int _pendingColumns;
    private int _scheduled;
    private int _disposed;

    public ScrollDispatcher(
        IExcelMacroQueue macroQueue,
        IExcelScroller scroller,
        IDiagnosticLog? log = null)
    {
        _macroQueue = macroQueue ?? throw new ArgumentNullException(nameof(macroQueue));
        _scroller = scroller ?? throw new ArgumentNullException(nameof(scroller));
        _log = log ?? NullDiagnosticLog.Instance;
    }

    public bool Enqueue(int columnDelta)
    {
        if (columnDelta == 0)
        {
            return true;
        }

        if (Volatile.Read(ref _disposed) != 0)
        {
            return false;
        }

        Interlocked.Add(ref _pendingColumns, columnDelta);
        return EnsureScheduled();
    }

    public void Dispose()
    {
        Interlocked.Exchange(ref _disposed, 1);
        Interlocked.Exchange(ref _pendingColumns, 0);
    }

    private bool EnsureScheduled()
    {
        if (Interlocked.CompareExchange(ref _scheduled, 1, 0) != 0)
        {
            return true;
        }

        try
        {
            _macroQueue.Queue(DrainOnExcelThread);
            return true;
        }
        catch (Exception exception)
        {
            Interlocked.Exchange(ref _scheduled, 0);
            Interlocked.Exchange(ref _pendingColumns, 0);
            _log.Write("scroll_queue_failed", exception);
            return false;
        }
    }

    private void DrainOnExcelThread()
    {
        try
        {
            if (Volatile.Read(ref _disposed) == 0)
            {
                var columns = Interlocked.Exchange(ref _pendingColumns, 0);
                if (columns != 0)
                {
                    _scroller.ScrollColumns(columns);
                }
            }
        }
        catch (Exception exception)
        {
            _log.Write("scroll_execute_failed", exception);
        }
        finally
        {
            Interlocked.Exchange(ref _scheduled, 0);
            if (Volatile.Read(ref _disposed) == 0 && Volatile.Read(ref _pendingColumns) != 0)
            {
                EnsureScheduled();
            }
        }
    }
}

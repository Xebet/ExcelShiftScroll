using System;

namespace ExcelShiftScroll.Input;

public sealed class HookLifetime : IDisposable
{
    private readonly object _gate = new();
    private readonly Func<IntPtr> _install;
    private readonly Func<IntPtr, bool> _uninstall;
    private IntPtr _handle;
    private bool _disposed;

    public HookLifetime(Func<IntPtr> install, Func<IntPtr, bool> uninstall)
    {
        _install = install ?? throw new ArgumentNullException(nameof(install));
        _uninstall = uninstall ?? throw new ArgumentNullException(nameof(uninstall));
    }

    public bool IsInstalled
    {
        get
        {
            lock (_gate)
            {
                return _handle != IntPtr.Zero;
            }
        }
    }

    public void Install()
    {
        lock (_gate)
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(HookLifetime));
            }

            if (_handle != IntPtr.Zero)
            {
                return;
            }

            var handle = _install();
            if (handle == IntPtr.Zero)
            {
                throw new InvalidOperationException("Windows did not install the Excel mouse hook.");
            }

            _handle = handle;
        }
    }

    public void Dispose()
    {
        lock (_gate)
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            if (_handle != IntPtr.Zero)
            {
                _uninstall(_handle);
                _handle = IntPtr.Zero;
            }
        }
    }
}

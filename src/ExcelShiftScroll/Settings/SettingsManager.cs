using System;

namespace ExcelShiftScroll.Settings;

public sealed class SettingsManager
{
    private readonly object _gate = new();
    private readonly ISettingsStore _store;
    private ScrollSettings _current;

    public SettingsManager(ISettingsStore store)
    {
        _store = store ?? throw new ArgumentNullException(nameof(store));
        _current = _store.Load().ValidatedCopy();
    }

    public event EventHandler? Changed;

    public ScrollSettings Current
    {
        get
        {
            lock (_gate)
            {
                return _current.Copy();
            }
        }
    }

    public void Update(Action<ScrollSettings> update)
    {
        if (update is null)
        {
            throw new ArgumentNullException(nameof(update));
        }

        lock (_gate)
        {
            var next = _current.Copy();
            update(next);
            _current = next.ValidatedCopy();
            _store.Save(_current);
        }

        Changed?.Invoke(this, EventArgs.Empty);
    }

    public void Reset() => Update(settings =>
    {
        var defaults = ScrollSettings.Defaults();
        settings.Enabled = defaults.Enabled;
        settings.ColumnsPerDetent = defaults.ColumnsPerDetent;
        settings.ReverseDirection = defaults.ReverseDirection;
        settings.DiagnosticsEnabled = defaults.DiagnosticsEnabled;
    });
}

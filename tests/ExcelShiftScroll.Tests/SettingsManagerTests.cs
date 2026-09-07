using System.IO;
using ExcelShiftScroll.Settings;
using Xunit;

namespace ExcelShiftScroll.Tests;

public sealed class SettingsManagerTests
{
    [Fact]
    public void SaveFailurePreservesMemoryAndDoesNotRaiseChanged()
    {
        var manager = new SettingsManager(new FailingStore());
        var changed = false;
        manager.Changed += (_, _) => changed = true;
        Assert.Throws<IOException>(() => manager.Update(s => s.Enabled = false));
        Assert.True(manager.Current.Enabled);
        Assert.False(changed);
    }

    private sealed class FailingStore : ISettingsStore
    {
        public ScrollSettings Load() => ScrollSettings.Defaults();
        public void Save(ScrollSettings settings) => throw new IOException("Test save failure");
    }
}

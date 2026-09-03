using System;
using System.IO;
using ExcelShiftScroll.Settings;
using Xunit;

namespace ExcelShiftScroll.Tests;

public sealed class JsonSettingsStoreTests
{
    [Fact]
    public void CorruptedSettingsRecoverToDefaults()
    {
        WithTemporarySettingsPath(path =>
        {
            File.WriteAllText(path, "{ definitely not valid JSON");

            var settings = new JsonSettingsStore(path).Load();

            Assert.True(settings.Enabled);
            Assert.Equal(3, settings.ColumnsPerDetent);
            Assert.False(settings.ReverseDirection);
            Assert.False(settings.DiagnosticsEnabled);
        });
    }

    [Fact]
    public void SavedSettingsRoundTrip()
    {
        WithTemporarySettingsPath(path =>
        {
            var store = new JsonSettingsStore(path);
            store.Save(new ScrollSettings
            {
                Enabled = false,
                ColumnsPerDetent = 10,
                ReverseDirection = true,
                DiagnosticsEnabled = true,
            });

            var settings = store.Load();

            Assert.False(settings.Enabled);
            Assert.Equal(10, settings.ColumnsPerDetent);
            Assert.True(settings.ReverseDirection);
            Assert.True(settings.DiagnosticsEnabled);
        });
    }

    [Fact]
    public void UnsupportedColumnCountRecoversToDefault()
    {
        WithTemporarySettingsPath(path =>
        {
            var store = new JsonSettingsStore(path);
            store.Save(new ScrollSettings { ColumnsPerDetent = 7 });

            Assert.Equal(3, store.Load().ColumnsPerDetent);
        });
    }

    private static void WithTemporarySettingsPath(Action<string> test)
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelShiftScroll.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            test(Path.Combine(directory, "settings.json"));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }
}

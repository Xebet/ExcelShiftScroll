using System;
using System.IO;
using System.Threading.Tasks;
using ExcelShiftScroll.Settings;
using Xunit;

namespace ExcelShiftScroll.Tests;

public sealed class JsonSettingsStoreTests
{
    [Theory]
    [InlineData("{}", true, 3)]
    [InlineData("{\"columnsPerDetent\":5}", true, 5)]
    [InlineData("{\"enabled\":false}", false, 3)]
    [InlineData("{\"futureField\":123}", true, 3)]
    public void MissingFieldsUseDefaultsWithoutOverridingExplicitValues(string json, bool enabled, int columns)
    {
        WithTemporarySettingsPath(path =>
        {
            File.WriteAllText(path, json);
            var settings = new JsonSettingsStore(path).Load();
            Assert.Equal(enabled, settings.Enabled);
            Assert.Equal(columns, settings.ColumnsPerDetent);
        });
    }

    [Fact]
    public void ConcurrentStoresLeaveValidJsonAndNoTemporaryFiles()
    {
        WithTemporarySettingsPath(path =>
        {
            Parallel.For(0, 40, i =>
            {
                var store = new JsonSettingsStore(path);
                store.Save(new ScrollSettings { Enabled = false, ColumnsPerDetent = 5 });
                Assert.False(store.Load().Enabled);
            });
            Assert.Equal(5, new JsonSettingsStore(path).Load().ColumnsPerDetent);
            Assert.Empty(Directory.GetFiles(Path.GetDirectoryName(path)!, "*.tmp"));
        });
    }

    [Fact]
    public void FailedReplacementPreservesPreviousFileAndCleansTemporaryFile()
    {
        WithTemporarySettingsPath(path =>
        {
            var store = new JsonSettingsStore(path);
            store.Save(new ScrollSettings { Enabled = false });
            using (File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                Assert.Throws<IOException>(() => store.Save(new ScrollSettings { Enabled = true }));
            }
            Assert.False(store.Load().Enabled);
            Assert.Empty(Directory.GetFiles(Path.GetDirectoryName(path)!, "*.tmp"));
        });
    }

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

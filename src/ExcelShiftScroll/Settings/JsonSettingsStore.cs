using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Security.Cryptography;
using System.Threading;

namespace ExcelShiftScroll.Settings;

public sealed class JsonSettingsStore : ISettingsStore
{
    private readonly string _path;
    private readonly string _mutexName;

    public JsonSettingsStore(string path)
    {
        _path = Path.GetFullPath(path ?? throw new ArgumentNullException(nameof(path)));
        using var hash = SHA256.Create();
        _mutexName = @"Local\ExcelShiftScroll.Settings." + BitConverter.ToString(
            hash.ComputeHash(Encoding.UTF8.GetBytes(_path.ToUpperInvariant()))).Replace("-", "");
    }

    public static string DefaultPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "ExcelShiftScroll",
        "settings.json");

    public ScrollSettings Load()
    {
        try
        {
            if (!File.Exists(_path))
            {
                return ScrollSettings.Defaults();
            }

            using var stream = new FileStream(_path, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete);
            var serializer = new DataContractJsonSerializer(typeof(ScrollSettings));
            return (serializer.ReadObject(stream) as ScrollSettings)?.ValidatedCopy()
                ?? ScrollSettings.Defaults();
        }
        catch (Exception exception) when (
            exception is IOException ||
            exception is UnauthorizedAccessException ||
            exception is SerializationException)
        {
            return ScrollSettings.Defaults();
        }
    }

    public void Save(ScrollSettings settings)
    {
        if (settings is null)
        {
            throw new ArgumentNullException(nameof(settings));
        }

        var directory = Path.GetDirectoryName(_path)
            ?? throw new InvalidOperationException("The settings path has no directory.");
        Directory.CreateDirectory(directory);

        var temporaryPath = _path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        using var mutex = new Mutex(false, _mutexName);
        var acquired = false;
        try
        {
            try { acquired = mutex.WaitOne(TimeSpan.FromSeconds(2)); }
            catch (AbandonedMutexException) { acquired = true; }
            if (!acquired) { throw new IOException("Settings are being saved by another Excel process. Try again."); }

            var serializer = new DataContractJsonSerializer(typeof(ScrollSettings));
            using (var stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                serializer.WriteObject(stream, settings.ValidatedCopy());
                stream.Flush(true);
            }

            if (File.Exists(_path)) { File.Replace(temporaryPath, _path, null); }
            else { File.Move(temporaryPath, _path); }
        }
        finally
        {
            try { File.Delete(temporaryPath); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
            if (acquired) { mutex.ReleaseMutex(); }
        }
    }
}

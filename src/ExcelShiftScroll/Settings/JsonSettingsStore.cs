using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace ExcelShiftScroll.Settings;

public sealed class JsonSettingsStore : ISettingsStore
{
    private readonly string _path;

    public JsonSettingsStore(string path)
    {
        _path = path ?? throw new ArgumentNullException(nameof(path));
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

            using var stream = File.OpenRead(_path);
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

        var temporaryPath = _path + ".tmp";
        var serializer = new DataContractJsonSerializer(typeof(ScrollSettings));
        using (var stream = File.Create(temporaryPath))
        {
            serializer.WriteObject(stream, settings.ValidatedCopy());
        }

        if (File.Exists(_path))
        {
            File.Replace(temporaryPath, _path, null);
        }
        else
        {
            File.Move(temporaryPath, _path);
        }
    }
}

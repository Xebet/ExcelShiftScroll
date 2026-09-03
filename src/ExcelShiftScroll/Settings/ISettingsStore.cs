namespace ExcelShiftScroll.Settings;

public interface ISettingsStore
{
    ScrollSettings Load();
    void Save(ScrollSettings settings);
}

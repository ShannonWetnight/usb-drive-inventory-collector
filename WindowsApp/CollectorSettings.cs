using System.Text.Json;

namespace USBDriveInventoryCollector;

internal static class CollectorSettings
{
    private static readonly string FilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "USBDriveInventoryCollector", "settings.json");

    public static string DefaultWorkbookPath => Path.Combine(AppContext.BaseDirectory, "Output", "Inventory.xlsx");
    public static string DefaultLogsDirectory => Path.Combine(AppContext.BaseDirectory, "Output", "Logs");
    public static bool SoundsEnabled()
    {
        try { return Load().SoundsEnabled ?? true; }
        catch { return true; }
    }

    public static string WorkbookPath()
    {
        try { var path = Load().WorkbookPath; if (!string.IsNullOrWhiteSpace(path)) return Path.GetFullPath(path); }
        catch { /* A damaged preference must not prevent opening the default workbook. */ }
        return DefaultWorkbookPath;
    }

    public static string LogsDirectory()
    {
        try { var path = Load().LogsDirectory; if (!string.IsNullOrWhiteSpace(path)) return Path.GetFullPath(path); }
        catch { /* Fall back to the default log directory. */ }
        return DefaultLogsDirectory;
    }

    public static void SaveWorkbookPath(string path)
    {
        SavePaths(path, LogsDirectory());
    }

    public static void SavePaths(string workbookPath, string logsDirectory)
        => Save(new SettingsData { WorkbookPath = Path.GetFullPath(workbookPath), LogsDirectory = Path.GetFullPath(logsDirectory), SoundsEnabled = SoundsEnabled() });

    public static void SaveSoundsEnabled(bool enabled)
        => Save(new SettingsData { WorkbookPath = WorkbookPath(), LogsDirectory = LogsDirectory(), SoundsEnabled = enabled });

    private static void Save(SettingsData settings)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
        var temp = FilePath + ".tmp";
        File.WriteAllText(temp, JsonSerializer.Serialize(settings));
        File.Move(temp, FilePath, true);
    }

    private static SettingsData Load() => File.Exists(FilePath) ? JsonSerializer.Deserialize<SettingsData>(File.ReadAllText(FilePath)) ?? new SettingsData() : new SettingsData();

    private sealed class SettingsData
    {
        public string? WorkbookPath { get; set; }
        public string? LogsDirectory { get; set; }
        public bool? SoundsEnabled { get; set; }
    }
}

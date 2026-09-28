using System.Text.Json;

namespace USBDriveInventoryCollector;

internal static class CollectorSettings
{
    private static readonly string FilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "USBDriveInventoryCollector", "settings.json");

    public static string DefaultWorkbookPath => Path.Combine(AppContext.BaseDirectory, "Output", "Inventory.xlsx");

    public static string WorkbookPath()
    {
        try
        {
            if (File.Exists(FilePath))
            {
                var settings = JsonSerializer.Deserialize<SettingsData>(File.ReadAllText(FilePath));
                if (!string.IsNullOrWhiteSpace(settings?.WorkbookPath)) return Path.GetFullPath(settings.WorkbookPath);
            }
        }
        catch { /* A damaged preference must not prevent opening the default workbook. */ }
        return DefaultWorkbookPath;
    }

    public static void SaveWorkbookPath(string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
        var temp = FilePath + ".tmp";
        File.WriteAllText(temp, JsonSerializer.Serialize(new SettingsData { WorkbookPath = Path.GetFullPath(path) }));
        File.Move(temp, FilePath, true);
    }

    private sealed class SettingsData
    {
        public string? WorkbookPath { get; set; }
    }
}

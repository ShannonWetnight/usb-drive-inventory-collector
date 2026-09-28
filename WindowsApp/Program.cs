using Microsoft.Win32;
using System.Management;
using System.Runtime.InteropServices;
using System.Security.Principal;

namespace USBDriveInventoryCollector;

internal static class Program
{
    [DllImport("kernel32.dll")] private static extern uint SetErrorMode(uint mode);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool AllocConsole();
    [STAThread]
    private static void Main(string[] args)
    {
        var oldMode = SetErrorMode(0x0001 | 0x8000);
        SetErrorMode(oldMode | 0x0001 | 0x8000);
        var terminal = args.Length == 1 && args[0].Equals("--terminal", StringComparison.OrdinalIgnoreCase);
        try
        {
            Directory.CreateDirectory(CollectorSettings.LogsDirectory());
            if (terminal)
            {
                if (!AllocConsole()) throw new IOException("Could not open a terminal window.");
                Console.SetOut(new StreamWriter(Console.OpenStandardOutput()) { AutoFlush = true });
                Console.SetError(new StreamWriter(Console.OpenStandardError()) { AutoFlush = true });
                Console.SetIn(new StreamReader(Console.OpenStandardInput()));
                TerminalCollector.Run();
            }
            else
            {
                ApplicationConfiguration.Initialize();
                Application.Run(new CollectorForm());
            }
        }
        catch (Exception ex)
        {
            var path = Path.Combine(CollectorSettings.LogsDirectory(), "startup-error.log");
            try { File.AppendAllText(path, $"{DateTime.Now:O} {ex}\n"); } catch { }
            if (terminal) { Console.Error.WriteLine($"The collector could not start: {ex}\nLog: {path}"); Console.WriteLine("Press [Enter] to close."); Console.ReadLine(); }
            else MessageBox.Show($"The collector could not start. {ex.Message}\n\nDetails: {path}", "USB Drive Inventory Collector", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally { SetErrorMode(oldMode); }
    }
}

internal sealed class AutoPlayGuard : IDisposable
{
    private const string KeyPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers";
    private bool _active, _hadValue, _keyExisted;
    private int _original;
    public string TryOffer(IWin32Window owner) => TryOffer(message =>
        MessageBox.Show(owner, message, "Temporary AutoPlay Setting", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes);

    public string TryOffer(Func<string, bool> confirm)
    {
        string? desktop = null;
        try { using var searcher = new ManagementObjectSearcher("SELECT UserName FROM Win32_ComputerSystem"); using var items = searcher.Get(); desktop = items.Cast<ManagementObject>().FirstOrDefault()?["UserName"]?.ToString(); }
        catch { }
        var ownSid = WindowsIdentity.GetCurrent().User;
        SecurityIdentifier? desktopSid = null;
        try { if (desktop is not null) desktopSid = (SecurityIdentifier)new NTAccount(desktop).Translate(typeof(SecurityIdentifier)); } catch { }
        if (desktopSid is null || ownSid is null || !desktopSid.Equals(ownSid)) return "AutoPlay prompt skipped: the signed-in desktop user differs from this account.";
        using var key = Registry.CurrentUser.OpenSubKey(KeyPath);
        _keyExisted = key is not null;
        _hadValue = key?.GetValueNames().Contains("DisableAutoplay") ?? false;
        if (_hadValue)
        {
            if (key!.GetValueKind("DisableAutoplay") != RegistryValueKind.DWord) return "AutoPlay has an unexpected setting type; no change was made.";
            _original = (int)key.GetValue("DisableAutoplay")!;
            if (_original == 1) return "AutoPlay is already disabled.";
        }
        if (!confirm("AutoPlay may open drive folders or show pop-ups. Disable it until the collector exits?"))
            return "AutoPlay was left unchanged.";
        using var writable = Registry.CurrentUser.CreateSubKey(KeyPath) ?? throw new IOException("Could not open AutoPlay preference.");
        _active = true; // restore even if the write partially fails
        writable.SetValue("DisableAutoplay", 1, RegistryValueKind.DWord);
        return "AutoPlay temporarily disabled for this Windows user.";
    }
    public void Dispose()
    {
        if (!_active) return;
        using (var key = Registry.CurrentUser.OpenSubKey(KeyPath, true))
        {
            if (key is null || !key.GetValueNames().Contains("DisableAutoplay") || key.GetValueKind("DisableAutoplay") != RegistryValueKind.DWord || (int)key.GetValue("DisableAutoplay")! != 1) return;
            if (_hadValue) key.SetValue("DisableAutoplay", _original, RegistryValueKind.DWord);
            else key.DeleteValue("DisableAutoplay", false);
        }
        _active = false;
        if (!_keyExisted)
        {
            using var parent = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Explorer", true);
            using var child = parent?.OpenSubKey("AutoplayHandlers");
            if (parent is not null && child is not null && child.ValueCount == 0 && child.SubKeyCount == 0)
            { child.Close(); parent.DeleteSubKey("AutoplayHandlers", false); }
        }
    }
}

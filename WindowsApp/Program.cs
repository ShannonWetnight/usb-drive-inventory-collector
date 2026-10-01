using Microsoft.Win32;
using System.Management;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Xml.Linq;

namespace USBDriveInventoryCollector;

internal static class Program
{
    [DllImport("kernel32.dll")] private static extern uint SetErrorMode(uint mode);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool AllocConsole();
    [STAThread]
    private static void Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        if (args.Length == 2 && args[0] == "--verify-workbook")
        {
            try
            {
                var book = new InventoryBook(args[1]);
                book.OpenOrCreate();
                var sample = new DriveRecord { ["Manufacturer"] = "Example", ["Model"] = "TEST-DRIVE", ["SerialNumber"] = "TEST-SERIAL", ["Capacity"] = "1 TB", ["Type"] = "SATA HDD" };
                book.Add(sample);
                var reopened = new InventoryBook(args[1]); reopened.OpenOrCreate();
                if (reopened.Records.Count != 1 || reopened.Records[0]["SerialNumber"] != "TEST-SERIAL" || reopened.Records[0]["Manufacturer"] != "Example") throw new InvalidDataException("Workbook round trip failed.");
                var legacyPath = Path.Combine(Path.GetDirectoryName(args[1])!, "Collector-Legacy-Workbook-Check.xlsx");
                File.Copy(args[1], legacyPath, true);
                using (var archive = ZipFile.Open(legacyPath, ZipArchiveMode.Update))
                {
                    var sheet = archive.GetEntry("xl/worksheets/sheet1.xml")!;
                    XDocument xml;
                    using (var stream = sheet.Open()) xml = XDocument.Load(stream);
                    var header = xml.Descendants().First(c => c.Name.LocalName == "c" && (string?)c.Attribute("r") == "A1");
                    header.Descendants().First(t => t.Name.LocalName == "t").Value = "Make";
                    sheet.Delete();
                    using var output = archive.CreateEntry("xl/worksheets/sheet1.xml").Open();
                    xml.Save(output);
                }
                var legacy = new InventoryBook(legacyPath); legacy.OpenOrCreate();
                if (legacy.Records.Count != 1 || legacy.Records[0]["Manufacturer"] != "Example" || legacy.Columns[0] != "Manufacturer") throw new InvalidDataException("Legacy workbook import failed.");
                using (var archive = ZipFile.OpenRead(legacyPath))
                using (var stream = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open())
                {
                    var migrated = XDocument.Load(stream);
                    if (!migrated.Descendants().Any(t => t.Name.LocalName == "t" && t.Value == "Manufacturer") || migrated.Descendants().Any(t => t.Name.LocalName == "t" && t.Value == "Make")) throw new InvalidDataException("Legacy header migration failed.");
                }
                VerifyEntryRemoval();
            }
            catch (Exception ex) { File.WriteAllText(args[1] + ".error.txt", ex.ToString()); Environment.ExitCode = 1; }
            return;
        }
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
    private static void VerifyEntryRemoval()
    {
        var directory = Path.Combine(Path.GetTempPath(), "Collector-Removal-Check-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var path = Path.Combine(directory, "Inventory.xlsx");
            var book = new InventoryBook(path);
            book.OpenOrCreate();
            book.ChangeColumns([..InventoryBook.Core, "FirmwareVersion"]);
            foreach (var serial in new[] { "FIRST", "MIDDLE", "LAST" })
                book.Add(new DriveRecord { ["Manufacturer"] = "Example", ["Model"] = serial + "-MODEL", ["SerialNumber"] = serial, ["FirmwareVersion"] = serial + "-FW" });

            void Check(params string[] serials)
            {
                var reopened = new InventoryBook(path);
                reopened.OpenOrCreate();
                if (!book.Records.Select(r => r["SerialNumber"]).SequenceEqual(serials) ||
                    !reopened.Records.Select(r => r["SerialNumber"]).SequenceEqual(serials) ||
                    !reopened.Columns.SequenceEqual(book.Columns) ||
                    reopened.Records.Any(r => r["FirmwareVersion"] != r["SerialNumber"] + "-FW"))
                    throw new InvalidDataException("Removal changed record order, values, or columns.");
                using var archive = ZipFile.OpenRead(path);
                using var stream = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
                var xml = XDocument.Load(stream);
                var rows = xml.Descendants().Where(e => e.Name.LocalName == "row").ToArray();
                if (!rows.Select(r => (string?)r.Attribute("r")).SequenceEqual(Enumerable.Range(1, serials.Length + 1).Select(i => i.ToString())))
                    throw new InvalidDataException("Removal did not renumber workbook rows.");
                var filter = xml.Descendants().FirstOrDefault(e => e.Name.LocalName == "autoFilter");
                if (serials.Length == 0 ? filter is not null : (string?)filter?.Attribute("ref") != $"A1:F{serials.Length + 1}")
                    throw new InvalidDataException("Removal did not update the workbook filter.");
            }

            foreach (var index in new[] { -1, book.Records.Count })
            {
                var rejected = false;
                try { book.Remove(index); }
                catch (ArgumentOutOfRangeException) { rejected = true; }
                if (!rejected) throw new InvalidDataException("Invalid removal index was accepted.");
            }
            Check("FIRST", "MIDDLE", "LAST");
            var previous = book.Records.ToArray();
            var originalBytes = File.ReadAllBytes(path);
            using (var locked = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                var failed = false;
                try { book.Remove(1); }
                catch (IOException) { failed = true; }
                if (!failed || !book.Records.SequenceEqual(previous)) throw new InvalidDataException("Failed removal did not restore the entry.");
            }
            if (!File.ReadAllBytes(path).SequenceEqual(originalBytes)) throw new InvalidDataException("Failed removal changed the workbook.");
            Check("FIRST", "MIDDLE", "LAST");
            book.Remove(1);
            Check("FIRST", "LAST");
            if (book.HasSerial("MIDDLE")) throw new InvalidDataException("Removed serial is still considered a duplicate.");
            book.Add(previous[1]);
            Check("FIRST", "LAST", "MIDDLE");
            book.Remove(2);
            Check("FIRST", "LAST");
            book.Remove(0);
            Check("LAST");
            book.Remove(0);
            Check();
            book.Add(previous[0]);
            Check("FIRST");
        }
        finally { Directory.Delete(directory, true); }
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

using System.Diagnostics;
using System.Text.RegularExpressions;

namespace USBDriveInventoryCollector;

internal sealed class TerminalCollector
{
    private InventoryBook _book = new(CollectorSettings.WorkbookPath());
    private readonly AutoPlayGuard _autoPlay = new();
    private readonly HashSet<int> _connected = [];
    private readonly string _logPath = Path.Combine(AppContext.BaseDirectory, "Output", "Logs", $"USB-Drive-Inventory-Collector-{DateTime.Now:yyyyMMdd-HHmmss}.log");
    private DriveProbe? _probe;
    private string _smartVersion = "N/A";
    private bool _paused;
    private volatile bool _stop;

    public static void Run() => new TerminalCollector().Start();

    private void Start()
    {
        try
        {
            _book.OpenOrCreate();
            Console.Title = "USB Drive Inventory Collector – Terminal";
            Console.WriteLine("USB Drive Inventory Collector – Terminal");
            Console.WriteLine("========================================");
            Console.WriteLine($"Workbook: {_book.Path}");
            Console.WriteLine($"Log: {_logPath}");
            var smart = DriveProbe.FindSmartctl();
            if (smart is null && Confirm("smartmontools is needed for automatic scanning. Install it with WinGet now?"))
            {
                using var install = Process.Start(new ProcessStartInfo("winget.exe") { UseShellExecute = false,
                    Arguments = "install --id smartmontools.smartmontools -e --source winget --accept-package-agreements --accept-source-agreements" });
                install?.WaitForExit();
                Environment.SetEnvironmentVariable("PATH", (Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Machine) ?? "") + ";" + (Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.User) ?? ""));
                smart = DriveProbe.FindSmartctl();
            }
            if (smart is not null)
            {
                _probe = new DriveProbe(smart, Log);
                try { _smartVersion = _probe.VersionAsync(CancellationToken.None).GetAwaiter().GetResult(); } catch (Exception ex) { Log(ex.ToString()); }
            }
            else Write("Automatic scanning is unavailable until smartmontools is installed. Manual Drive Entry is available.");
            Write(_autoPlay.TryOffer(Confirm));
            Help();
            Console.CancelKeyPress += OnCancel;
            try { ScanLoop(); }
            finally { Console.CancelKeyPress -= OnCancel; }
        }
        finally
        {
            try { _autoPlay.Dispose(); } catch (Exception ex) { Log(ex.ToString()); Console.Error.WriteLine("AutoPlay could not be restored. Check Windows AutoPlay settings."); }
            Write("Inventory collector stopped.");
            Console.WriteLine($"Inventory: {_book.Path}\nLog: {_logPath}");
            Console.Write("Press [Enter] to close Terminal."); Console.ReadLine();
        }
    }

    private void ScanLoop()
    {
        Write("Waiting for a USB drive. Press [M] for manual entry.");
        while (!_stop)
        {
            if (Console.KeyAvailable)
            {
                var key = char.ToUpperInvariant(Console.ReadKey(true).KeyChar);
                if (key == 'Q') return;
                if (key == 'M') ManualEntry(false);
                else if (key == 'L') ManualEntry(true);
                else if (key == 'S') Setup();
                else if (key == 'D') Details();
                else if (key == 'P') { _paused = !_paused; Write(_paused ? "Scanning paused. Press [P] to resume." : "Scanning resumed."); }
                else if (key == 'H') Help();
            }
            if (!_paused && _probe is not null)
            {
                try
                {
                    var disks = DriveProbe.Disks();
                    var present = disks.Select(d => d.Number).ToHashSet();
                    foreach (var old in _connected.Where(n => !present.Contains(n)).ToList()) { _connected.Remove(old); Write($"Disk {old} removed. Ready for the next drive."); }
                    foreach (var disk in disks.Where(d => !_connected.Contains(d.Number)))
                    {
                        _connected.Add(disk.Number);
                        Write($"USB drive detected on Disk {disk.Number}. Reading drive identity...");
                        try
                        {
                            var record = _probe.IdentifyAsync(disk, CancellationToken.None).GetAwaiter().GetResult();
                            if (_book.HasSerial(record["SerialNumber"])) Write("This serial number is already in the workbook. No row added.");
                            else { var row = _book.Add(record); Recorded(record, row); }
                            Console.WriteLine("Remove this drive and insert the next drive. Press [M] for manual entry.");
                        }
                        catch (Exception ex) { Write($"Disk {disk.Number} could not be read: {ex.Message}"); Log(ex.ToString()); Console.WriteLine("Remove and reinsert to retry, or press [M] for manual entry."); }
                    }
                }
                catch (Exception ex) { Write("Scan error: " + ex.Message); Log(ex.ToString()); }
            }
            Thread.Sleep(250);
        }
    }

    private void OnCancel(object? sender, ConsoleCancelEventArgs args) { args.Cancel = true; _stop = true; }

    private void ManualEntry(bool copyLast)
    {
        if (copyLast && _book.Records.Count == 0) { Write("There is no saved drive to copy."); return; }
        var draft = copyLast ? Draft.From(_book.Records[^1]) : new Draft();
        var copy = copyLast;
        while (true)
        {
            Console.Clear(); Console.WriteLine("MANUAL DRIVE ENTRY\n==================");
            Console.WriteLine("Press [Enter] for N/A or type :cancel to return to scanning.\n");
            if (copy) { Console.WriteLine($"Copying: {draft.Make} / {draft.Model} / {draft.Capacity} / {draft.Type}"); if (!Field("Serial Number", ref draft.Serial)) return; }
            else
            {
                if (!Field("Step 1 of 5 – Make", ref draft.Make) || !Field("Step 2 of 5 – Model", ref draft.Model) ||
                    !Field("Step 3 of 5 – Serial Number", ref draft.Serial) || !Capacity(ref draft) || !Type(ref draft)) return;
            }
            while (true)
            {
                DriveRecord record;
                try
                {
                    var match = Regex.Match(draft.Capacity, @"\A([0-9]+(?:\.[0-9]+)?)\s+([A-Za-z]+)\z");
                    var unit = match.Success ? match.Groups[2].Value : "";
                    var listed = new[] { "B", "KB", "MB", "GB", "TB", "PB" }.Contains(unit);
                    record = ManualValidation.Create(draft.Make, draft.Model, draft.Serial, match.Success ? match.Groups[1].Value : draft.Capacity,
                        listed ? unit : match.Success ? "Other" : "", listed ? "" : unit,
                        DriveTypes.Options.Contains(draft.Type, StringComparer.OrdinalIgnoreCase) ? draft.Type : "Other",
                        DriveTypes.Options.Contains(draft.Type, StringComparer.OrdinalIgnoreCase) ? "" : draft.Type);
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); Console.Write("Edit a field [E] or cancel [C]: "); if (Console.ReadLine()?.Trim().Equals("E", StringComparison.OrdinalIgnoreCase) != true) return; if (!EditDraft(ref draft)) return; continue; }
                Console.Clear(); Console.WriteLine("REVIEW MANUAL DRIVE\n==================="); Print(record);
                var duplicate = _book.HasSerial(record["SerialNumber"]);
                if (duplicate)
                {
                    Console.WriteLine("This serial number is already in the workbook.");
                    Console.Write("Change serial [S] or cancel [C]: ");
                    if (Console.ReadLine()?.Trim().Equals("S", StringComparison.OrdinalIgnoreCase) != true) return;
                    if (!Field("Serial Number", ref draft.Serial)) return;
                    continue;
                }
                Console.Write("Save and choose next [Y], save and copy [L], edit [E], or cancel [C]: ");
                var action = Console.ReadLine()?.Trim().ToUpperInvariant();
                if (action == "C" || action is null) return;
                if (action == "E") { if (!EditDraft(ref draft)) return; continue; }
                if (action != "Y" && action != "L") continue;
                try { var row = _book.Add(record); Recorded(record, row); }
                catch (Exception ex) { Write("Could not save: " + ex.Message); Log(ex.ToString()); break; }
                copy = action == "L";
                draft = copy ? Draft.From(record) : new Draft();
                if (copy) draft.Serial = "";
                Console.Write("Press [Enter] for another drive, or [R] to return to scanning: ");
                if (Console.ReadLine()?.Trim().Equals("R", StringComparison.OrdinalIgnoreCase) == true) return;
                break;
            }
        }
    }

    private static bool EditDraft(ref Draft draft)
    {
        Console.Write("Edit Make [1], Model [2], Serial [3], Capacity [4], Type [5], or back [B]: ");
        switch (Console.ReadLine()?.Trim())
        {
            case "1": return Field("Make", ref draft.Make);
            case "2": return Field("Model", ref draft.Model);
            case "3": return Field("Serial Number", ref draft.Serial);
            case "4": return Capacity(ref draft);
            case "5": return Type(ref draft);
            default: return true;
        }
    }

    private static bool Field(string label, ref string value)
    {
        Console.Write($"{label} (or :cancel): ");
        var entry = Console.ReadLine();
        if (entry is null || entry.Trim().Equals(":cancel", StringComparison.OrdinalIgnoreCase)) return false;
        value = entry.Trim(); return true;
    }
    private static bool Capacity(ref Draft draft)
    {
        Console.Write("Step 4 of 5 – Capacity (number only; unit next, [Enter] for N/A): ");
        var amount = Console.ReadLine()?.Trim();
        if (amount is null || amount.Equals(":cancel", StringComparison.OrdinalIgnoreCase)) return false;
        if (amount.Length == 0) { draft.Capacity = ""; return true; }
        var unit = Choose("Capacity Unit", ["B", "KB", "MB", "GB", "TB", "PB", "Other"]);
        if (unit is null) return false;
        var selectedUnit = unit;
        if (selectedUnit == "Other" && !Field("Custom Capacity Unit", ref selectedUnit)) return false;
        draft.Capacity = amount + " " + selectedUnit; return true;
    }
    private static bool Type(ref Draft draft)
    {
        var type = Choose("Step 5 of 5 – Drive Type", DriveTypes.Options);
        if (type is null) return false;
        var selectedType = type;
        if (selectedType == "Other" && !Field("Custom Drive Type", ref selectedType)) return false;
        draft.Type = selectedType; return true;
    }
    private static string? Choose(string label, IReadOnlyList<string> options)
    {
        Console.WriteLine(label + ":");
        for (var i = 0; i < options.Count; i++) Console.WriteLine($"  {i + 1}. {options[i]}");
        while (true)
        {
            Console.Write("Choose a number (or :cancel): "); var input = Console.ReadLine()?.Trim();
            if (input is null || input.Equals(":cancel", StringComparison.OrdinalIgnoreCase)) return null;
            if (int.TryParse(input, out var n) && n >= 1 && n <= options.Count) return options[n - 1];
            Console.WriteLine("Choose a listed number.");
        }
    }

    private void Setup()
    {
        var optional = InventoryBook.Catalog.Select(c => c.Key).Except(InventoryBook.Core).ToArray();
        var selected = _book.Columns.ToHashSet();
        var location = _book.Path;
        while (true)
        {
            Console.Clear(); Console.WriteLine("WORKBOOK SETUP\n==============");
            Console.WriteLine($"Workbook Save Location: {location}\n");
            for (var i = 0; i < optional.Length; i++) Console.WriteLine($"{i + 1,2}. [{(selected.Contains(optional[i]) ? 'X' : ' ')}] {InventoryBook.Header(optional[i])}");
            Console.Write("Toggle column number, location [L], defaults [D], apply [A], or cancel [C]: ");
            var input = Console.ReadLine()?.Trim().ToUpperInvariant();
            if (input == "C" || input is null) return;
            if (input == "D") { selected = InventoryBook.Core.ToHashSet(); continue; }
            if (input == "L")
            {
                Console.Write("Enter full .xlsx path (or :cancel): "); var next = Console.ReadLine()?.Trim().Trim('"');
                if (string.IsNullOrWhiteSpace(next) || next.Equals(":cancel", StringComparison.OrdinalIgnoreCase)) continue;
                if (!next.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) { Write("Choose an .xlsx file."); continue; }
                if (File.Exists(next) && !string.Equals(Path.GetFullPath(next), Path.GetFullPath(_book.Path), StringComparison.OrdinalIgnoreCase) && !Confirm("That workbook exists. Switch to its records?")) continue;
                location = next; continue;
            }
            if (input == "A")
            {
                try
                {
                    var next = _book.AtLocation(location);
                    var columns = InventoryBook.Core.Concat(optional.Where(selected.Contains)).ToList();
                    if (!columns.SequenceEqual(next.Columns)) next.ChangeColumns(columns);
                    CollectorSettings.SaveWorkbookPath(next.Path); _book = next;
                    Write($"Workbook Setup saved: {_book.Path}"); return;
                }
                catch (Exception ex) { Write("Setup failed: " + ex.Message); Log(ex.ToString()); Console.ReadLine(); }
            }
            else if (int.TryParse(input, out var n) && n >= 1 && n <= optional.Length)
            { if (!selected.Add(optional[n - 1])) selected.Remove(optional[n - 1]); }
        }
    }

    private void Details()
    {
        Console.Clear(); Console.WriteLine($"VERSION INFORMATION\n===================\nUSB Drive Inventory Collector v{Application.ProductVersion}\nMaintainer: Shannon Wetnight\nWorkbook: {_book.Path}\nLog: {_logPath}\nsmartctl: {_smartVersion}\nColumns: {string.Join(", ", _book.Columns.Select(InventoryBook.Header))}\n");
        Console.Write("Press [Enter] to return."); Console.ReadLine();
    }
    private static bool Confirm(string message) { Console.Write($"{message} [Y/N]: "); return Console.ReadLine()?.Trim().Equals("Y", StringComparison.OrdinalIgnoreCase) == true; }
    private static void Print(DriveRecord record)
    { foreach (var key in InventoryBook.Core) Console.WriteLine($"{InventoryBook.Header(key),-20} {record[key]}"); }
    private void Recorded(DriveRecord record, int row) { Write($"Drive recorded as row {row}."); Print(record); }
    private static void Help() => Console.WriteLine("Usage: [M] Manual Drive Entry  [L] Copy Last Drive  [S] Workbook Setup  [P] Pause Scanning  [D] Version Information  [H] Help  [Q] Finish");
    private void Write(string text) { Console.WriteLine(text); Log(text); }
    private void Log(string text) { try { File.AppendAllText(_logPath, $"{DateTime.Now:O} {text}\n"); } catch { } }

    private sealed class Draft
    {
        public string Make = "", Model = "", Serial = "", Capacity = "", Type = "N/A";
        public static Draft From(DriveRecord record) => new() { Make = record["Make"], Model = record["Model"], Serial = "",
            Capacity = record["Capacity"] == "N/A" ? "" : record["Capacity"], Type = record["Type"] };
    }
}

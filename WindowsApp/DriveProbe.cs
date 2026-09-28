using System.Diagnostics;
using System.Globalization;
using System.Management;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace USBDriveInventoryCollector;

internal sealed record UsbDisk(int Number, string FriendlyName, string SerialNumber);
internal sealed class DriveProbe
{
    private readonly string _smartctl;
    private readonly Action<string> _log;
    private readonly Dictionary<int, string> _preferred = [];
    private static readonly string[] Fallbacks = ["auto", "sat", "sntjmicron", "sntjmicron/sat", "sntrealtek", "sntrealtek/sat", "sntasmedia", "sntasmedia/sat", "usbjmicron", "usbprolific", "usbsunplus", "usbcypress"];
    public DriveProbe(string smartctl, Action<string> log) { _smartctl = smartctl; _log = log; }
    public static string? FindSmartctl()
    {
        var locations = new[] { Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "smartmontools", "bin", "smartctl.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "smartmontools", "bin", "smartctl.exe") };
        foreach (var path in locations) if (File.Exists(path)) return path;
        var env = Environment.GetEnvironmentVariable("PATH") ?? "";
        foreach (var folder in env.Split(';', StringSplitOptions.RemoveEmptyEntries))
        { try { var path = Path.Combine(folder.Trim('"'), "smartctl.exe"); if (File.Exists(path)) return path; } catch { } }
        return null;
    }
    public static List<UsbDisk> Disks()
    {
        var result = new List<UsbDisk>();
        using var searcher = new ManagementObjectSearcher(@"root\Microsoft\Windows\Storage", "SELECT Number,BusType,IsBoot,IsSystem,FriendlyName,SerialNumber FROM MSFT_Disk");
        using var items = searcher.Get();
        foreach (ManagementObject disk in items)
        {
            using (disk)
            {
                // MSFT_Disk.BusType 7 is USB. Never inventory a system or boot disk.
                if (Convert.ToInt32(disk["BusType"]) != 7 || Convert.ToBoolean(disk["IsBoot"]) || Convert.ToBoolean(disk["IsSystem"])) continue;
                result.Add(new UsbDisk(Convert.ToInt32(disk["Number"]), Value(disk["FriendlyName"]), Value(disk["SerialNumber"])));
            }
        }
        return result;
    }
    public async Task<string> VersionAsync(CancellationToken ct) => (await RunAsync(["--version"], ct)).Output.Split('\n')[0].Trim();
    public async Task<DriveRecord> IdentifyAsync(UsbDisk disk, CancellationToken ct)
    {
        var device = $"/dev/pd{disk.Number}";
        var transports = new List<string>();
        void Add(string? item) { if (!string.IsNullOrWhiteSpace(item) && !transports.Contains(item)) transports.Add(item); }
        if (_preferred.TryGetValue(disk.Number, out var preferred)) Add(preferred);
        // As in the legacy collector, a stalled scan-open is a timeout for the entire attempt.
        try
        {
            using var scan = JsonDocument.Parse((await RunAsync(["--scan-open", "-j"], ct)).Output);
            if (scan.RootElement.TryGetProperty("devices", out var devices))
                foreach (var item in devices.EnumerateArray())
                {
                    var name = Prop(item, "name"); var info = Prop(item, "info_name");
                    if (name == device || Regex.IsMatch(info, $@"PhysicalDrive{disk.Number}(\D|$)", RegexOptions.IgnoreCase))
                    {
                        var type = Prop(item, "type"); if (type != "N/A" && type != "auto" && type != "scsi") Add(type);
                    }
                }
        }
        catch (TimeoutException) { throw; }
        catch (Exception ex) { _log("scan-open discovery failed: " + ex.Message); }
        foreach (var item in Fallbacks) Add(item);
        DriveRecord? fallback = null;
        var lastError = "No supported transport returned usable media identity.";
        for (int cycle = 0; cycle < 2; cycle++)
        {
            foreach (var transport in transports)
            {
                ct.ThrowIfCancellationRequested();
                var args = new List<string> { "-i", "-jo" };
                if (transport != "auto") { args.Add("-d"); args.Add(transport); }
                args.Add(device);
                try
                {
                    _log($"identity disk={disk.Number} cycle={cycle + 1} transport={transport}");
                    var result = await RunAsync(args, ct);
                    if (string.IsNullOrWhiteSpace(result.Output)) { lastError = "smartctl returned no output."; continue; }
                    using var doc = JsonDocument.Parse(result.Output);
                    var root = doc.RootElement;
                    var model = Prop(root, "model_name"); var serial = Prop(root, "serial_number");
                    var capacityBytes = Prop(root, "user_capacity", "bytes");
                    var capacity = Capacity(capacityBytes);
                    if ((model == "N/A" && serial == "N/A") || capacity == "N/A") { lastError = $"No usable identity via {transport}."; continue; }
                    var protocol = Prop(root, "device", "protocol");
                    var record = new DriveRecord {
                        ["Make"] = Make(model), ["Model"] = model, ["SerialNumber"] = serial,
                        ["Capacity"] = capacity, ["Type"] = Type(root, model), ["Interface"] = Interface(root, model),
                        ["Protocol"] = protocol, ["Transport"] = transport,
                        ["FirmwareVersion"] = Prop(root, "firmware_version"), ["ModelFamily"] = Prop(root, "model_family"),
                        ["FormFactor"] = Prop(root, "form_factor", "name"), ["RotationRate"] = Prop(root, "rotation_rate"),
                        ["CapacityBytes"] = capacityBytes, ["LogicalBlockSize"] = Prop(root, "logical_block_size"),
                        ["PhysicalBlockSize"] = Prop(root, "physical_block_size"), ["AtaVersion"] = Prop(root, "ata_version", "string"),
                        ["SataVersion"] = Prop(root, "sata_version", "string")
                    };
                    if (Regex.IsMatch(protocol, "ATA|SATA|NVMe", RegexOptions.IgnoreCase)) { _preferred[disk.Number] = transport; return record; }
                    var generic = Regex.IsMatch(model, "^(SSK|USB|USB Device|External|Generic|Mass Storage|JMicron|ASMedia|Realtek|SCSI Disk Device)$", RegexOptions.IgnoreCase);
                    var bridge = generic || (model == disk.FriendlyName && serial == disk.SerialNumber && Regex.IsMatch(model, "SSK|USB|External|Generic|JMicron|ASMedia|Realtek", RegexOptions.IgnoreCase));
                    if (!bridge) fallback ??= record;
                }
                catch (TimeoutException) { throw; }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex) { lastError = ex.Message; _log($"identity disk={disk.Number} transport={transport}: {ex}"); }
            }
            if (fallback is not null) return fallback;
            if (cycle == 0) await Task.Delay(1250, ct);
        }
        throw new IOException($"Unable to read media identity from {device}. {lastError}");
    }
    private async Task<(int Code, string Output, string Error)> RunAsync(IEnumerable<string> arguments, CancellationToken ct)
    {
        using var process = new Process { StartInfo = new ProcessStartInfo(_smartctl) { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true } };
        foreach (var argument in arguments) process.StartInfo.ArgumentList.Add(argument);
        process.Start();
        var stdout = process.StandardOutput.ReadToEndAsync(ct); var stderr = process.StandardError.ReadToEndAsync(ct);
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(TimeSpan.FromSeconds(30));
        try { await process.WaitForExitAsync(timeout.Token); }
        catch (OperationCanceledException)
        {
            try { process.Kill(true); } catch { }
            if (ct.IsCancellationRequested) throw;
            throw new TimeoutException("smartctl timed out after 30 seconds. The USB adapter may need to be unplugged and reconnected.");
        }
        var output = await stdout; var error = await stderr;
        if (!string.IsNullOrWhiteSpace(error)) _log("smartctl stderr: " + error.Trim());
        return (process.ExitCode, output, error);
    }
    private static string Value(object? x) => string.IsNullOrWhiteSpace(x?.ToString()) ? "N/A" : x!.ToString()!.Trim();
    private static string Prop(JsonElement element, params string[] path)
    {
        foreach (var key in path)
            if (element.ValueKind == JsonValueKind.Object && element.TryGetProperty(key, out var next)) element = next;
            else return "N/A";
        return element.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined ? "N/A" : Value(element.ToString());
    }
    private static string Capacity(string bytes)
    {
        if (!decimal.TryParse(bytes, NumberStyles.Any, CultureInfo.InvariantCulture, out var n) || n <= 0) return "N/A";
        return n >= 1_000_000_000_000m ? (n / 1_000_000_000_000m).ToString("0.##", CultureInfo.InvariantCulture) + " TB" :
            Math.Round(n / 1_000_000_000m, 0).ToString(CultureInfo.InvariantCulture) + " GB";
    }
    private static string Interface(JsonElement root, string model)
    {
        var protocol = Prop(root, "device", "protocol");
        var text = root.TryGetProperty("smartctl", out var smart) && smart.TryGetProperty("output", out var output) && output.ValueKind == JsonValueKind.Array ? string.Join("\n", output.EnumerateArray().Select(x => x.ToString())) : "";
        if (protocol.Contains("NVMe", StringComparison.OrdinalIgnoreCase) || root.TryGetProperty("nvme_version", out _)) return "NVMe";
        if (Regex.IsMatch(text, "^Transport Type:\\s*Parallel\\b", RegexOptions.Multiline | RegexOptions.IgnoreCase) || protocol is "PATA" or "IDE") return "PATA";
        if (root.TryGetProperty("sata_version", out _) || protocol == "SATA" || text.Contains("SATA Version is:")) return "SATA";
        if (Regex.IsMatch(model, @"^(WDC\s+)?WD800AAJB(-|$)", RegexOptions.IgnoreCase)) return "PATA";
        return protocol == "ATA" || root.TryGetProperty("ata_version", out _) ? "ATA (interface unknown)" : protocol;
    }
    private static string Type(JsonElement root, string model)
    {
        var form = Prop(root, "form_factor", "name");
        var protocol = Prop(root, "device", "protocol");
        var iface = Interface(root, model);
        var rpm = Prop(root, "rotation_rate");
        if (iface == "PATA" && Regex.IsMatch(model, @"^(WDC\s+)?WD800AAJB(-|$)", RegexOptions.IgnoreCase)) { if (form == "N/A") form = "3.5 inches"; if (rpm == "N/A") rpm = "7200"; }
        var solid = rpm == "0" || Regex.IsMatch(model, "SSD|Solid[ _-]?State", RegexOptions.IgnoreCase);
        var rotational = long.TryParse(rpm, out var n) && n > 0;
        string size = form switch { "1.8 inches" => "1.8-inch ", "2.5 inches" => "2.5-inch ", "3.5 inches" => "3.5-inch ", "< 1.8 inches" => "<1.8-inch ", _ => "" };
        if (form == "M.2") size = "M.2 ";
        if (form == "mSATA") size = "mSATA ";
        if (protocol.Contains("NVMe", StringComparison.OrdinalIgnoreCase) || root.TryGetProperty("nvme_version", out _)) return size == "" ? "NVMe SSD" : size + "NVMe SSD";
        if (iface is "PATA" or "ATA (interface unknown)") return size + (iface == "PATA" ? "PATA" : "ATA") + (solid ? " SSD" : rotational ? " HDD" : " Drive");
        if (iface == "SATA") return form == "mSATA" && solid ? "mSATA SSD" : size + "SATA " + (solid ? "SSD" : rotational ? "HDD" : "Drive");
        if (solid || rotational) return size + (solid ? "SSD" : "HDD");
        return "N/A";
    }
    private static string Make(string model)
    {
        (string Pattern, string Make)[] patterns = [
            (@"^LENSE", "Lenovo"), (@"Samsung|^MZ[A-Z0-9]", "Samsung"), (@"SK[\s_-]*hynix|^HFS|^HFM", "SK hynix"),
            (@"KIOXIA|^KBG|^KXG", "KIOXIA"), (@"TOSHIBA|^THNS", "Toshiba"), (@"Western Digital|WDC|^WDS", "Western Digital"),
            (@"SanDisk", "SanDisk"), (@"Micron|^MTFD", "Micron"), (@"Crucial", "Crucial"), (@"KINGSTON", "Kingston"),
            (@"Intel|^SSDPE|^SSDSC", "Intel"), (@"Solidigm", "Solidigm"), (@"LITEON|LITE-ON", "Lite-On"),
            (@"ADATA", "ADATA"), (@"Seagate|^ST[0-9]", "Seagate"), (@"Hitachi|HGST", "HGST"), (@"PNY", "PNY"), (@"Transcend", "Transcend")];
        return patterns.FirstOrDefault(p => Regex.IsMatch(model, p.Pattern, RegexOptions.IgnoreCase)).Make ?? "N/A";
    }
}

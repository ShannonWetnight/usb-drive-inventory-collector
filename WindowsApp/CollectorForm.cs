using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;

namespace USBDriveInventoryCollector;

internal sealed class CollectorForm : Form
{
    private readonly InventoryBook _book;
    private readonly AutoPlayGuard _autoPlay = new();
    private readonly string _logPath;
    private DriveProbe? _probe;
    private readonly HashSet<int> _connected = [];
    private readonly CancellationTokenSource _closing = new();
    private readonly System.Windows.Forms.Timer _timer = new() { Interval = 1000 };
    private bool _busy, _paused, _modal;
    private string _version = "N/A";
    private readonly Label _status = new() { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 12, FontStyle.Bold), Text = "Initializing..." };
    private readonly Label _guidance = new() { Dock = DockStyle.Fill, Text = "Insert one drive at a time." };
    private readonly Label _footer = new() { Dock = DockStyle.Fill, AutoEllipsis = true };
    private readonly DataGridView _grid = new() { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, SelectionMode = DataGridViewSelectionMode.FullRowSelect };
    private readonly ListBox _activity = new() { Dock = DockStyle.Fill, HorizontalScrollbar = true, Font = new Font("Consolas", 9) };
    private readonly TextBox _details = new() { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, WordWrap = false, ScrollBars = ScrollBars.Both, Font = new Font("Consolas", 9) };
    private readonly TabControl _tabs = new() { Dock = DockStyle.Fill };
    private readonly Button _pause = new() { Text = "Pause scanning", Width = 130 };
    private readonly Button _copy = new() { Text = "Copy last", Width = 110 };
    private readonly Button _manual = new() { Text = "Manual entry", Width = 125 };
    private readonly Button _setup = new() { Text = "Workbook setup", Width = 140 };
    private readonly Button _technical = new() { Text = "Technical details", Width = 145 };
    private readonly Button _finish = new() { Text = "Finish", Width = 90 };
    public CollectorForm()
    {
        Text = "USB Drive Inventory Collector v4.0.0";
        MinimumSize = new Size(900, 590); Size = new Size(1030, 720); StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 10); BackColor = Color.FromArgb(246, 248, 251);
        var output = Path.Combine(AppContext.BaseDirectory, "Output");
        Directory.CreateDirectory(Path.Combine(output, "Logs"));
        _book = new InventoryBook(Path.Combine(output, "Inventory.xlsx"));
        _logPath = Path.Combine(output, "Logs", $"USB-Drive-Inventory-Collector-{DateTime.Now:yyyyMMdd-HHmmss}.log");
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 5, ColumnCount = 1 };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        foreach (var height in new[] { 82f, 56f, 72f }) layout.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
        Controls.Add(layout);
        var header = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(32, 55, 78) };
        var title = new Label { Text = "USB Drive Inventory Collector", ForeColor = Color.White, Font = new Font("Segoe UI", 18, FontStyle.Bold), Dock = DockStyle.Top, Height = 44, Padding = new Padding(20, 8, 0, 0) };
        var subtitle = new Label { Text = "One drive at a time. Every record is saved immediately.", ForeColor = Color.FromArgb(215, 229, 238), Dock = DockStyle.Fill, Padding = new Padding(22, 3, 0, 0) };
        header.Controls.Add(subtitle); header.Controls.Add(title); layout.Controls.Add(header, 0, 0);
        var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, Padding = new Padding(12, 10, 0, 0) };
        foreach (var b in new[] { _pause, _manual, _copy, _setup, _technical, _finish }) { b.Height = 34; b.Margin = new Padding(0, 0, 8, 0); actions.Controls.Add(b); }
        layout.Controls.Add(actions, 0, 1);
        var status = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, Padding = new Padding(22, 0, 0, 0) };
        status.RowStyles.Add(new RowStyle(SizeType.Percent, 55)); status.RowStyles.Add(new RowStyle(SizeType.Percent, 45)); status.Controls.Add(_status, 0, 0); status.Controls.Add(_guidance, 0, 1);
        layout.Controls.Add(status, 0, 2);
        _tabs.TabPages.Add(new TabPage("Recorded drives") { Controls = { _grid } });
        _tabs.TabPages.Add(new TabPage("Activity") { Controls = { _activity } });
        _tabs.TabPages.Add(new TabPage("Details") { Controls = { _details } });
        layout.Controls.Add(_tabs, 0, 3);
        _footer.Padding = new Padding(18, 9, 0, 0); layout.Controls.Add(_footer, 0, 4);
        _copy.Enabled = false;
        _pause.Click += (_, _) => { _paused = !_paused; _pause.Text = _paused ? "Resume scanning" : "Pause scanning"; Activity(_paused ? "Scanning paused." : "Waiting for a USB drive..."); };
        _manual.Click += (_, _) => ManualEntry(false);
        _copy.Click += (_, _) => ManualEntry(true);
        _setup.Click += (_, _) => Setup();
        _technical.Click += (_, _) => { RefreshDetails(); _tabs.SelectedIndex = 2; };
        _finish.Click += (_, _) => Close();
        _timer.Tick += async (_, _) => await PollAsync();
        Shown += async (_, _) => await InitializeAsync();
        FormClosing += (_, _) => { _timer.Stop(); _closing.Cancel(); try { _autoPlay.Dispose(); } catch (Exception ex) { Log(ex.ToString()); MessageBox.Show(this, "AutoPlay could not be restored. Check Windows AutoPlay settings."); } };
    }
    private void Log(string text) { try { File.AppendAllText(_logPath, $"{DateTime.Now:O} {text}\n"); } catch { } }
    private void Activity(string text, string level = "INFO")
    {
        _status.Text = text;
        _activity.Items.Insert(0, $"{DateTime.Now:HH:mm:ss}  {level,-5}  {text}");
        if (_activity.Items.Count > 500) _activity.Items.RemoveAt(500);
        Log($"{level} {text}");
    }
    private async Task InitializeAsync()
    {
        try
        {
            _book.OpenOrCreate(); RefreshGrid();
            var smart = DriveProbe.FindSmartctl();
            if (smart is null)
            {
                if (MessageBox.Show(this, "smartmontools is required for automatic drive detection. Install it with WinGet now?", "Dependency required", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    using var install = Process.Start(new ProcessStartInfo("winget.exe") { UseShellExecute = true, Arguments = "install --id smartmontools.smartmontools -e --source winget --accept-package-agreements --accept-source-agreements" });
                    if (install is null) throw new IOException("WinGet could not be started.");
                    await install.WaitForExitAsync();
                    if (install.ExitCode != 0) throw new IOException($"WinGet failed with exit code {install.ExitCode}.");
                    Environment.SetEnvironmentVariable("PATH", (Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Machine) ?? "") + ";" + (Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.User) ?? ""));
                    smart = DriveProbe.FindSmartctl();
                }
                if (smart is null) throw new FileNotFoundException("smartctl.exe is required. Install smartmontools and reopen the collector.");
            }
            _probe = new DriveProbe(smart, Log);
            try { _version = await _probe.VersionAsync(_closing.Token); } catch (Exception ex) { Log(ex.ToString()); }
            try { Activity(_autoPlay.TryOffer(this)); } catch (Exception ex) { Activity("AutoPlay setting could not be changed.", "WARN"); Log(ex.ToString()); }
            Activity("Waiting for a USB drive...");
            _guidance.Text = "Insert one drive at a time. Use Manual entry or Workbook setup at any time.";
            RefreshDetails(); _timer.Start();
        }
        catch (Exception ex) { Log(ex.ToString()); MessageBox.Show(this, ex.Message, "Startup failed", MessageBoxButtons.OK, MessageBoxIcon.Error); Close(); }
    }
    private async Task PollAsync()
    {
        if (_busy || _paused || _modal || _probe is null || _closing.IsCancellationRequested) return;
        _busy = true;
        try
        {
            var disks = await Task.Run(DriveProbe.Disks, _closing.Token);
            if (_closing.IsCancellationRequested) return;
            var numbers = disks.Select(d => d.Number).ToHashSet();
            foreach (var old in _connected.Where(n => !numbers.Contains(n)).ToList()) { _connected.Remove(old); Activity($"Disk {old} removed. Ready for another drive."); }
            var disk = disks.FirstOrDefault(d => !_connected.Contains(d.Number));
            if (disk is null) return;
            _connected.Add(disk.Number);
            Activity($"USB drive detected on Disk {disk.Number}. Reading drive identity...");
            _guidance.Text = "Reading drive identity. Slow adapters may reach the 30-second timeout.";
            await Task.Delay(2000, _closing.Token);
            try
            {
                var record = await _probe.IdentifyAsync(disk, _closing.Token);
                if (_closing.IsCancellationRequested) return;
                if (_book.HasSerial(record["SerialNumber"])) { Activity($"Disk {disk.Number} duplicate: {record["SerialNumber"]}. No row added.", "WARN"); _guidance.Text = "Remove and insert the next drive, or use Manual entry."; System.Media.SystemSounds.Exclamation.Play(); }
                else { var row = _book.Add(record); RefreshGrid(); Activity($"USB drive recorded as row {row}: {record["Model"]} / {record["SerialNumber"]}"); _guidance.Text = $"Saved as row {row}. Remove this drive and insert the next."; System.Media.SystemSounds.Asterisk.Play(); }
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            { Activity($"Disk {disk.Number} could not be read: {ex.Message}", "ERROR"); _guidance.Text = "Remove and reinsert to retry, or use Manual entry."; Log(ex.ToString()); }
        }
        catch (OperationCanceledException) { }
        catch (Exception ex) { Activity("Scan error: " + ex.Message, "ERROR"); Log(ex.ToString()); }
        finally { _busy = false; }
    }
    private void RefreshGrid()
    {
        _grid.Columns.Clear();
        foreach (var key in _book.Columns) _grid.Columns.Add(key, InventoryBook.Header(key));
        _grid.Rows.Clear();
        foreach (var row in _book.Records) _grid.Rows.Add(_book.Columns.Select(key => row[key]).Cast<object>().ToArray());
        _copy.Enabled = _book.Records.Count > 0;
        RefreshDetails();
    }
    private void RefreshDetails()
    {
        _footer.Text = $"Records: {_book.Records.Count}     Workbook: {_book.Path}";
        _details.Lines = ["USB Drive Inventory Collector v4.0.0", "Maintainer: Shannon Wetnight", "Repository: https://github.com/ShannonWetnight/usb-drive-inventory-collector", $"Workbook: {_book.Path}", $"Debug log: {_logPath}", $"smartctl: {_version}", "Scope: USB physical drives; boot and system disks excluded", "Transport: smartctl autodetection plus USB adapter fallbacks", "Workbook backend: Direct XLSX (no Excel COM)", "Timeout: 30 seconds per smartctl process", "Workbook columns: " + string.Join(", ", _book.Columns.Select(InventoryBook.Header))];
    }
    private void Setup()
    {
        _modal = true;
        try
        {
            using var dialog = new Form { Text = "Workbook setup", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, ClientSize = new Size(550, 535), MaximizeBox = false, MinimizeBox = false };
            var intro = new Label { Text = "Default columns: Make, Model, Serial Number, Reported Capacity, Type.\nSelect extra identity fields to save for each drive:", Bounds = new Rectangle(20, 12, 510, 55) };
            var list = new CheckedListBox { CheckOnClick = true, Bounds = new Rectangle(20, 75, 510, 320) };
            var optional = InventoryBook.Catalog.Select(c => c.Key).Except(InventoryBook.Core).ToList();
            foreach (var key in optional) list.Items.Add(InventoryBook.Header(key), _book.Columns.Contains(key));
            var note = new Label { Text = "Older and manual records receive N/A for extra fields. A backup keeps any fields removed from the active workbook.", Bounds = new Rectangle(20, 402, 510, 55) };
            var all = Button("Select all", 20, 478, 112); var defaults = Button("Defaults", 140, 478, 112); var apply = Button("Apply", 290, 478, 112); var cancel = Button("Cancel", 410, 478, 112);
            all.Click += (_, _) => { for (var i = 0; i < list.Items.Count; i++) list.SetItemChecked(i, true); };
            defaults.Click += (_, _) => { for (var i = 0; i < list.Items.Count; i++) list.SetItemChecked(i, false); };
            cancel.Click += (_, _) => dialog.Close();
            apply.Click += (_, _) =>
            {
                var chosen = InventoryBook.Core.Concat(optional.Where((_, i) => list.GetItemChecked(i))).ToList();
                if (chosen.SequenceEqual(_book.Columns)) { dialog.Close(); return; }
                try { _book.ChangeColumns(chosen); RefreshGrid(); Activity("Workbook columns updated."); dialog.Close(); }
                catch (Exception ex) { Log(ex.ToString()); MessageBox.Show(dialog, "Setup was not applied: " + ex.Message); }
            };
            dialog.Controls.AddRange([intro, list, note, all, defaults, apply, cancel]);
            dialog.ShowDialog(this);
        }
        finally { _modal = false; }
    }
    private static Button Button(string text, int x, int y, int width) => new() { Text = text, Bounds = new Rectangle(x, y, width, 34) };
    private void ManualEntry(bool copyLast)
    {
        _modal = true;
        try
        {
            using var dialog = new Form { Text = copyLast ? "Copy last saved drive" : "Manual drive entry", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, ClientSize = new Size(560, 410), MaximizeBox = false, MinimizeBox = false, Font = new Font("Segoe UI", 10) };
            dialog.Controls.Add(new Label { Text = "Leave a field blank for N/A. Model and serial are saved in uppercase.", Bounds = new Rectangle(20, 12, 520, 30) });
            var labels = new[] { "1. Make", "2. Model", "3. Serial number", "4. Capacity (number only)", "Capacity unit", "5. Drive type" };
            for (int i = 0; i < labels.Length; i++) dialog.Controls.Add(new Label { Text = labels[i], Bounds = new Rectangle(20, 47 + 43 * i, 195, 26) });
            var make = Box(219, 47, 320); var model = Box(219, 90, 320); var serial = Box(219, 133, 320); var amount = Box(219, 176, 175);
            model.CharacterCasing = CharacterCasing.Upper; serial.CharacterCasing = CharacterCasing.Upper;
            var unit = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Bounds = new Rectangle(219, 219, 135, 28) };
            unit.Items.AddRange(["N/A", "B", "KB", "MB", "GB", "TB", "PB", "Other"]); unit.SelectedIndex = 0;
            var customUnit = Box(365, 219, 174); customUnit.Visible = false;
            var type = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Bounds = new Rectangle(219, 262, 320, 28) };
            type.Items.Add("N/A"); foreach (var (category, values) in DriveTypes.Groups) foreach (var value in values) type.Items.Add(category + " / " + value);
            type.SelectedIndex = 0;
            var customType = Box(219, 299, 320); customType.Visible = false;
            unit.SelectedIndexChanged += (_, _) => customUnit.Visible = unit.Text == "Other";
            type.SelectedIndexChanged += (_, _) => customType.Visible = type.Text == "Other / Other";
            dialog.Controls.AddRange([make, model, serial, amount, unit, customUnit, type, customType]);
            if (copyLast && _book.Records.LastOrDefault() is { } last) Fill(last);
            var review = Button("Review drive", 20, 355, 130); var back = Button("Return to scanning", 162, 355, 165);
            back.Click += (_, _) => dialog.Close();
            review.Click += (_, _) =>
            {
                try
                {
                    var choice = type.SelectedIndex > 0 ? type.Text.Split(" / ", 2)[1] : "";
                    var record = ManualValidation.Create(make.Text, model.Text, serial.Text, amount.Text, unit.Text == "N/A" ? "" : unit.Text, customUnit.Text, choice, customType.Text);
                    var duplicate = _book.HasSerial(record["SerialNumber"]);
                    var action = Review(record, duplicate);
                    if (action == "Serial") { serial.Focus(); serial.SelectAll(); return; }
                    if (action == "Cancel") { dialog.Close(); return; }
                    if (action == "Edit") return;
                    var row = _book.Add(record); RefreshGrid(); Activity($"Manual drive recorded as row {row}: {record["Model"]} / {record["SerialNumber"]}");
                    MessageBox.Show(dialog, $"Drive saved as row {row}.", "Drive recorded", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    serial.Clear();
                    if (action == "Save") { make.Clear(); model.Clear(); amount.Clear(); unit.SelectedIndex = 0; customUnit.Clear(); type.SelectedIndex = 0; customType.Clear(); }
                    else dialog.Text = "Copy saved drive with a new serial";
                    serial.Focus();
                }
                catch (Exception ex) { Log(ex.ToString()); MessageBox.Show(dialog, ex.Message, "Manual entry", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            };
            dialog.Controls.AddRange([review, back]); dialog.ShowDialog(this);
            void Fill(DriveRecord record)
            {
                make.Text = record["Make"] == "N/A" ? "" : record["Make"];
                model.Text = record["Model"] == "N/A" ? "" : record["Model"];
                var match = Regex.Match(record["Capacity"], @"^([0-9]+(?:\.[0-9]+)?)\s+(.+)$");
                if (match.Success) { amount.Text = match.Groups[1].Value; var value = match.Groups[2].Value; if (unit.Items.Contains(value)) unit.SelectedItem = value; else { unit.SelectedItem = "Other"; customUnit.Text = value; } }
                var media = record["Type"];
                foreach (var item in type.Items) if (item.ToString()!.EndsWith(" / " + media, StringComparison.Ordinal)) { type.SelectedItem = item; break; }
                if (type.SelectedIndex == 0 && media != "N/A") { type.SelectedItem = "Other / Other"; customType.Text = media; }
            }
        }
        finally { _modal = false; }
    }
    private static TextBox Box(int x, int y, int width) => new() { Bounds = new Rectangle(x, y, width, 28) };
    private string Review(DriveRecord record, bool duplicate)
    {
        using var dialog = new Form { Text = duplicate ? "Duplicate serial number" : "Review manual drive", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, ClientSize = new Size(490, 345), MaximizeBox = false, MinimizeBox = false };
        var summary = new TextBox { Bounds = new Rectangle(20, 20, 450, 158), Multiline = true, ReadOnly = true, Font = new Font("Consolas", 11), Lines = [$"Make:     {record["Make"]}", $"Model:    {record["Model"]}", $"Serial:   {record["SerialNumber"]}", $"Capacity: {record["Capacity"]}", $"Type:     {record["Type"]}"] };
        var note = new Label { Text = duplicate ? "This serial is already in the workbook. Change it or cancel this record." : record["SerialNumber"] == "N/A" ? "Serial N/A cannot be checked for duplicates." : "Review these values before saving a new row.", Bounds = new Rectangle(20, 188, 450, 48) };
        dialog.Controls.AddRange([summary, note]); string action = "Edit";
        void Choice(string name, string text, int x, int width) { var button = Button(text, x, 272, width); button.Click += (_, _) => { action = name; dialog.Close(); }; dialog.Controls.Add(button); }
        if (duplicate) { Choice("Serial", "Change serial", 20, 140); Choice("Cancel", "Cancel record", 180, 140); }
        else { Choice("Save", "Save & next", 14, 105); Choice("Copy", "Save & copy", 125, 105); Choice("Edit", "Edit fields", 236, 105); Choice("Cancel", "Cancel record", 347, 128); }
        dialog.ShowDialog(this); return action;
    }
}

internal static class DriveTypes
{
    public static readonly (string Category, string[] Values)[] Groups = [
        ("Standard", ["1.8-inch SATA SSD", "2.5-inch IDE HDD", "2.5-inch SATA HDD", "2.5-inch SATA SSD", "3.5-inch IDE HDD", "3.5-inch SATA HDD", "IDE Drive", "IDE HDD", "IDE SSD", "M.2 NVMe SSD", "M.2 SATA SSD", "mSATA SSD", "NVMe SSD", "SATA Drive", "SATA HDD", "SATA SSD"]),
        ("Enterprise", ["2.5-inch SAS HDD", "2.5-inch SAS SSD", "3.5-inch SAS HDD", "3.5-inch SAS SSD", "SAS HDD", "SAS SSD"]),
        ("Other", ["3.5-inch Floppy Disk", "5.25-inch Floppy Disk", "CompactFlash Card", "eMMC", "HDD", "microSD Card", "SD Card", "SSD", "USB Flash Drive", "Other"])
    ];
}

internal static class ManualValidation
{
    public static DriveRecord Create(string make, string model, string serial, string amount, string unit, string customUnit, string type, string customType)
    {
        make = make.Trim(); model = model.Trim(); serial = serial.Trim(); amount = amount.Trim(); unit = unit.Trim(); customUnit = customUnit.Trim(); type = type.Trim(); customType = customType.Trim();
        Check(make, @"\A[A-Za-z0-9][A-Za-z0-9 .&()+'/_-]{0,79}\z", 80, "Make");
        Check(model, @"\A[A-Za-z0-9][A-Za-z0-9 .+/_-]{0,99}\z", 100, "Model");
        Check(serial, @"\A[A-Za-z0-9][A-Za-z0-9./_-]{0,99}\z", 100, "Serial");
        string capacity = "N/A";
        if (amount.Length > 0)
        {
            if (!Regex.IsMatch(amount, @"\A[0-9]{1,15}(?:\.[0-9]{1,6})?\z") || !decimal.TryParse(amount, NumberStyles.Number, CultureInfo.InvariantCulture, out var n) || n <= 0) throw new ArgumentException("Capacity: enter a positive number, such as 1 or 0.005, without the unit.");
            if (unit == "Other") { if (!Regex.IsMatch(customUnit, @"\A[A-Za-z]{1,12}\z")) throw new ArgumentException("Custom capacity unit must be 1–12 letters."); unit = customUnit; }
            if (unit.Length > 0 && !(new[] { "B", "KB", "MB", "GB", "TB", "PB" }.Contains(unit) || unit == customUnit)) throw new ArgumentException("Choose a listed capacity unit or enter a custom unit.");
            if (unit == "B" && n != decimal.Truncate(n)) throw new ArgumentException("Bytes must be a whole number.");
            if (unit.Length > 0) capacity = n.ToString("0.######", CultureInfo.InvariantCulture) + " " + unit;
        }
        if (type == "Other") { Check(customType, @"\A[A-Za-z0-9][A-Za-z0-9 .()+/_-]{0,59}\z", 60, "Custom drive type"); type = customType; }
        return new DriveRecord { ["Make"] = make, ["Model"] = model.ToUpperInvariant(), ["SerialNumber"] = serial.ToUpperInvariant(), ["Capacity"] = capacity, ["Type"] = type };
    }
    private static void Check(string value, string pattern, int max, string label) { if (value.Length > 0 && (value.Length > max || !Regex.IsMatch(value, pattern))) throw new ArgumentException($"{label}: use only plain letters, digits, spaces, or standard punctuation (up to {max} characters)."); }
}

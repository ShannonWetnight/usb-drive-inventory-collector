using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;

namespace USBDriveInventoryCollector;

internal sealed class CollectorForm : Form
{
    private InventoryBook _book;
    private readonly AutoPlayGuard _autoPlay = new();
    private readonly string _logPath;
    private DriveProbe? _probe;
    private readonly HashSet<int> _connected = [];
    private readonly CancellationTokenSource _closing = new();
    private readonly System.Windows.Forms.Timer _timer = new() { Interval = 1000 };
    private bool _busy, _paused, _modal;
    private string _version = "N/A";
    private readonly RichTextBox _status = new() { Dock = DockStyle.Fill, ReadOnly = true, TabStop = false, BorderStyle = BorderStyle.None, ScrollBars = RichTextBoxScrollBars.None, BackColor = Color.FromArgb(246, 248, 251), Font = new Font("Segoe UI", 12, FontStyle.Bold), Text = "Initializing..." };
    private readonly Label _guidance = new() { Dock = DockStyle.Fill, Text = "Insert one drive at a time." };
    private readonly Label _footer = new() { AutoEllipsis = true, AutoSize = false, UseMnemonic = false, TextAlign = ContentAlignment.MiddleLeft };
    private readonly Button _copyPath = new() { Text = "Copy Path", Size = new Size(95, 28) };
    private readonly Button _openPath = new() { Text = "Open Folder", Size = new Size(110, 28) };
    private readonly DataGridView _grid = new() { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None, SelectionMode = DataGridViewSelectionMode.FullRowSelect, EnableHeadersVisualStyles = false, ColumnHeadersHeight = 34 };
    private readonly ListBox _activity = new() { Dock = DockStyle.Fill, HorizontalScrollbar = true, Font = new Font("Consolas", 9) };
    private readonly TabControl _tabs = new() { Dock = DockStyle.Fill, SizeMode = TabSizeMode.Fixed, ItemSize = new Size(130, 36) };
    private readonly RichTextBox _terminalOutput = new() { Dock = DockStyle.Fill, ReadOnly = true, TabStop = false, BackColor = Color.FromArgb(18, 22, 28), ForeColor = Color.Gainsboro, Font = new Font("Consolas", 10), BorderStyle = BorderStyle.None };
    private readonly TextBox _terminalInput = new() { Font = new Font("Consolas", 10), Text = "Terminal disabled", AutoSize = false, Enabled = false };
    private readonly Button _terminalSend = new() { Text = "Send", Size = new Size(78, 28), Enabled = false };
    private readonly TabPage _terminalPage = new("Terminal");
    private EmbeddedTerminalIO? _terminalIO;
    private TerminalCollector? _terminalSession;
    private Task? _terminalTask;
    private int _terminalGeneration;
    private readonly ToolTip _toolTip = new();
    private readonly Button _pause = new() { Text = "Pause Scanning", Width = 140 };
    private readonly Button _manual = new() { Text = "Manual Drive Entry", Width = 175 };
    private readonly Button _setup = new() { Text = "Workbook Setup", Width = 150 };
    private readonly Button _terminalToggle = new() { Text = "Enable Terminal", Width = 140, Height = 24, Visible = false };
    private readonly Button _finish = new() { Text = "Finish", Width = 90 };
    public CollectorForm()
    {
        Text = "USB Drive Inventory Collector";
        KeyPreview = true;
        MinimumSize = new Size(900, 590); Size = new Size(1030, 720); StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 10); BackColor = Color.FromArgb(246, 248, 251);
        var output = Path.Combine(AppContext.BaseDirectory, "Output");
        Directory.CreateDirectory(Path.Combine(output, "Logs"));
        _book = new InventoryBook(CollectorSettings.WorkbookPath());
        _logPath = Path.Combine(output, "Logs", $"USB-Drive-Inventory-Collector-{DateTime.Now:yyyyMMdd-HHmmss}.log");
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 5, ColumnCount = 1 };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        foreach (var height in new[] { 82f, 56f, 72f }) layout.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
        Controls.Add(layout);
        var header = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(32, 55, 78) };
        var title = new Label { Text = "USB Drive Inventory Collector", ForeColor = Color.White, Font = new Font("Segoe UI", 18, FontStyle.Bold), Dock = DockStyle.Top, Height = 44, Padding = new Padding(20, 8, 60, 0) };
        var subtitle = new Label { Text = "One drive at a time. Every record is saved immediately.", ForeColor = Color.FromArgb(215, 229, 238), Dock = DockStyle.Fill, Padding = new Padding(22, 3, 0, 0) };
        var info = new Button { Text = "Version Information", AccessibleName = "Version Information", Bounds = new Rectangle(0, 12, 160, 34), FlatStyle = FlatStyle.Flat, ForeColor = Color.White, BackColor = Color.FromArgb(48, 76, 102), Font = new Font("Segoe UI", 9), TabStop = true };
        info.Click += (_, _) => ShowVersionInformation();
        _toolTip.SetToolTip(info, "Version Information");
        var headerActions = new Panel { Dock = DockStyle.Right, Width = 178, BackColor = header.BackColor };
        headerActions.Controls.Add(info);
        headerActions.Resize += (_, _) => info.Location = new Point((headerActions.ClientSize.Width - info.Width) / 2, (headerActions.ClientSize.Height - info.Height) / 2);
        info.Location = new Point((headerActions.ClientSize.Width - info.Width) / 2, (headerActions.ClientSize.Height - info.Height) / 2);
        header.Controls.Add(subtitle); header.Controls.Add(title); header.Controls.Add(headerActions); layout.Controls.Add(header, 0, 0);
        var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, Padding = new Padding(12, 10, 0, 0) };
        foreach (var b in new[] { _pause, _manual, _setup, _finish }) { b.Height = 34; b.Margin = new Padding(0, 0, 8, 0); actions.Controls.Add(b); }
        layout.Controls.Add(actions, 0, 1);
        var status = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, Padding = new Padding(22, 0, 0, 0) };
        status.RowStyles.Add(new RowStyle(SizeType.Percent, 55)); status.RowStyles.Add(new RowStyle(SizeType.Percent, 45)); status.Controls.Add(_status, 0, 0); status.Controls.Add(_guidance, 0, 1);
        layout.Controls.Add(status, 0, 2);
        _tabs.TabPages.Add(new TabPage("Recorded Drives") { Controls = { _grid } });
        _tabs.TabPages.Add(new TabPage("Activity") { Controls = { _activity } });
        var terminalEntry = new Panel { Dock = DockStyle.Bottom, Height = 44, BackColor = Color.FromArgb(230, 234, 240) };
        terminalEntry.Controls.Add(_terminalInput); terminalEntry.Controls.Add(_terminalSend);
        terminalEntry.Resize += (_, _) =>
        {
            var top = (terminalEntry.ClientSize.Height - 28) / 2;
            _terminalSend.Location = new Point(terminalEntry.ClientSize.Width - _terminalSend.Width - 8, top);
            _terminalInput.SetBounds(8, top, Math.Max(80, _terminalSend.Left - 16), 28);
        };
        _terminalPage.Controls.Add(_terminalOutput); _terminalPage.Controls.Add(terminalEntry); _tabs.TabPages.Add(_terminalPage);
        var tabHost = new Panel { Dock = DockStyle.Fill };
        tabHost.Controls.Add(_tabs); tabHost.Controls.Add(_terminalToggle);
        void PositionTerminalToggle()
        {
            if (!_tabs.IsHandleCreated) return;
            var tab = _tabs.GetTabRect(_tabs.TabPages.IndexOf(_terminalPage));
            _terminalToggle.Height = 28;
            _terminalToggle.Location = new Point(tab.Right + 8, tab.Top + (tab.Height - _terminalToggle.Height) / 2);
        }
        tabHost.Resize += (_, _) => PositionTerminalToggle();
        _tabs.HandleCreated += (_, _) => PositionTerminalToggle();
        layout.Controls.Add(tabHost, 0, 3);
        var footer = new Panel { Dock = DockStyle.Fill };
        footer.Controls.Add(_footer); footer.Controls.Add(_copyPath); footer.Controls.Add(_openPath);
        footer.Resize += (_, _) =>
        {
            const int gap = 8;
            _openPath.Location = new Point(footer.ClientSize.Width - _openPath.Width - 14, (footer.ClientSize.Height - _openPath.Height) / 2);
            _copyPath.Location = new Point(_openPath.Left - _copyPath.Width - gap, (footer.ClientSize.Height - _copyPath.Height) / 2);
            _footer.SetBounds(18, 0, Math.Max(0, _copyPath.Left - gap - 18), footer.ClientSize.Height);
        };
        _toolTip.SetToolTip(_footer, "Workbook Save Location"); _toolTip.SetToolTip(_copyPath, "Copy workbook location to clipboard"); _toolTip.SetToolTip(_openPath, "Open the workbook folder in File Explorer");
        layout.Controls.Add(footer, 0, 4);
        _pause.Click += (_, _) => SetScanningPaused(!_paused);
        _manual.Click += (_, _) => ManualEntry(false);
        _setup.Click += (_, _) => Setup();
        _terminalToggle.Click += (_, _) => { if (_terminalSession is null) _terminalTask = OpenTerminalAsync(); else _terminalSession.Stop(); };
        _tabs.SelectedIndexChanged += (_, _) => { _terminalToggle.Visible = ReferenceEquals(_tabs.SelectedTab, _terminalPage) || _terminalSession is not null; PositionTerminalToggle(); if (_terminalToggle.Visible) _terminalToggle.BringToFront(); };
        _terminalInput.KeyDown += (_, e) => { if (e.KeyCode != Keys.Enter || _terminalIO is null) return; e.SuppressKeyPress = true; SubmitTerminalInput(); };
        _terminalSend.Click += (_, _) => SubmitTerminalInput();
        KeyPress += TerminalKeyPress;
        _copyPath.Click += (_, _) => CopyWorkbookPath(_book.Path, this);
        _openPath.Click += (_, _) => OpenWorkbookFolder(_book.Path, this);
        _grid.CellDoubleClick += (_, e) => { if (e.RowIndex >= 0 && _grid.Rows[e.RowIndex].Tag is int recordIndex) EditRecord(recordIndex); };
        _finish.Click += async (_, _) =>
        {
            if (_busy) { MessageBox.Show(this, "Wait for the current drive read to finish before closing.", "Finish"); return; }
            if (_terminalSession is not null) { _terminalSession.Stop(); if (_terminalTask is not null) await _terminalTask; }
            if (!IsDisposed)
            {
                _modal = true;
                try { ShowFinishDialog(); }
                finally { _modal = false; }
            }
        };
        _timer.Tick += async (_, _) => await PollAsync();
        Shown += async (_, _) => await InitializeAsync();
        FormClosing += (_, _) => { _timer.Stop(); _closing.Cancel(); _terminalSession?.Stop(); try { _autoPlay.Dispose(); } catch (Exception ex) { Log(ex.ToString()); MessageBox.Show(this, "AutoPlay could not be restored. Check Windows AutoPlay settings."); } };
    }
    private void Log(string text) { try { File.AppendAllText(_logPath, $"{DateTime.Now:O} {text}\n"); } catch { } }
    private void SetScanningPaused(bool paused)
    {
        _paused = paused;
        _pause.Text = paused ? "Resume Scanning" : "Pause Scanning";
        _terminalSession?.SetPaused(paused);
        if (_terminalSession is not null) WriteTerminalOutput(_terminalGeneration, paused ? "Scanning paused from the main window. Press [P] or Resume Scanning to continue.\n" : "Scanning resumed from the main window.\n");
        Activity(paused ? (_terminalSession is null ? "Scanning is paused." : "Terminal scanning paused.") : "Scanning resumed.");
    }
    private void Activity(string text, string level = "INFO")
    {
        var display = _terminalSession is not null && !text.StartsWith("Terminal closed.", StringComparison.Ordinal) &&
            !text.Contains("Scanning is paused", StringComparison.Ordinal)
            ? text + (_paused ? " Scanning is paused in Terminal and this window." : " Scanning is paused in this window.") : text;
        if (_paused && _terminalSession is null && !display.Contains("Scanning is paused", StringComparison.Ordinal))
            display += " Scanning is paused.";
        _status.Text = display;
        _status.SelectAll(); _status.SelectionColor = Color.FromArgb(15, 20, 25);
        const string paused = "Scanning is paused";
        var pausedAt = display.IndexOf(paused, StringComparison.Ordinal);
        if (pausedAt >= 0)
        {
            var sentenceEnd = display.IndexOf('.', pausedAt);
            _status.Select(pausedAt, sentenceEnd >= 0 ? sentenceEnd - pausedAt + 1 : paused.Length);
            _status.SelectionColor = Color.Firebrick;
        }
        _status.Select(0, 0);
        _activity.Items.Insert(0, $"{DateTime.Now:HH:mm:ss}  {level,-5}  {text}");
        if (_activity.Items.Count > 500) _activity.Items.RemoveAt(500);
        Log($"{level} {text}");
    }
    private void CopyWorkbookPath(string path, IWin32Window owner)
    {
        try { Clipboard.SetText(path); Activity("Workbook location copied to clipboard."); }
        catch (Exception ex) { MessageBox.Show(owner, ex.Message, "Copy Workbook Location"); }
    }
    private static void OpenWorkbookFolder(string path, IWin32Window owner)
    {
        try { Process.Start(new ProcessStartInfo("explorer.exe") { UseShellExecute = true, Arguments = File.Exists(path) ? $"/select,\"{path}\"" : $"\"{Path.GetDirectoryName(path)}\"" }); }
        catch (Exception ex) { MessageBox.Show(owner, ex.Message, "Open Workbook Folder"); }
    }
    private void ShowFinishDialog()
    {
        using var dialog = new Form { Text = "Inventory Saved", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, ClientSize = new Size(560, 160), MaximizeBox = false, MinimizeBox = false, Font = new Font("Segoe UI", 10) };
        var path = _book.Path;
        var message = new Label { Text = "Workbook saved at:", Bounds = new Rectangle(20, 16, 520, 25) };
        var location = new TextBox { Text = path, ReadOnly = true, Bounds = new Rectangle(20, 46, 520, 28), TabStop = false };
        var copy = Button("Copy Path", 20, 102, 160);
        var open = Button("Open Folder", 200, 102, 160);
        var close = Button("Close", 380, 102, 160);
        _toolTip.SetToolTip(location, path);
        copy.Click += (_, _) => CopyWorkbookPath(path, dialog);
        open.Click += (_, _) => OpenWorkbookFolder(path, dialog);
        close.Click += (_, _) => { dialog.DialogResult = DialogResult.OK; dialog.Close(); };
        dialog.AcceptButton = close;
        dialog.Controls.AddRange([message, location, copy, open, close]);
        if (dialog.ShowDialog(this) == DialogResult.OK) Close();
    }
    private void SubmitTerminalInput()
    {
        if (_terminalIO is null) return;
        var answer = _terminalInput.Text;
        _terminalInput.Clear();
        _terminalOutput.AppendText(answer + Environment.NewLine);
        _terminalIO.Submit(answer);
        _terminalInput.Focus();
    }
    private void TerminalKeyPress(object? sender, KeyPressEventArgs e)
    {
        var io = _terminalIO;
        if (io is null || !ReferenceEquals(_tabs.SelectedTab, _terminalPage)) return;
        if ((ModifierKeys & (Keys.Control | Keys.Alt)) != Keys.None) return;
        if (io.WaitingForLine)
        {
            if (_terminalInput.Focused) return;
            e.Handled = true;
            _terminalInput.Focus();
            if (e.KeyChar == '\r') SubmitTerminalInput();
            else if (!char.IsControl(e.KeyChar)) _terminalInput.AppendText(e.KeyChar.ToString());
            return;
        }
        e.Handled = true;
        var command = char.ToUpperInvariant(e.KeyChar);
        if ("MLSPDHQ".Contains(command))
        {
            _terminalOutput.AppendText($"[{command}]{Environment.NewLine}");
            io.Submit(command.ToString());
        }
        _terminalInput.Focus();
    }
    private async Task InitializeAsync()
    {
        try
        {
            _book.OpenOrCreate(); RefreshGrid();
            var smart = DriveProbe.FindSmartctl();
            if (smart is null)
            {
                if (MessageBox.Show(this, "smartmontools is required for automatic drive detection. Install it with WinGet now?", "Dependency Required", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    using var install = Process.Start(new ProcessStartInfo("winget.exe") { UseShellExecute = true, Arguments = "install --id smartmontools.smartmontools -e --source winget --accept-package-agreements --accept-source-agreements" });
                    if (install is null) throw new IOException("WinGet could not be started.");
                    await install.WaitForExitAsync();
                    if (install.ExitCode != 0) throw new IOException($"WinGet failed with exit code {install.ExitCode}.");
                    Environment.SetEnvironmentVariable("PATH", (Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Machine) ?? "") + ";" + (Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.User) ?? ""));
                    smart = DriveProbe.FindSmartctl();
                }
                if (smart is null) Activity("Automatic scanning is unavailable until smartmontools is installed. Manual Drive Entry is available.", "WARN");
            }
            if (smart is not null)
            {
                _probe = new DriveProbe(smart, Log);
                try { _version = await _probe.VersionAsync(_closing.Token); } catch (Exception ex) { Log(ex.ToString()); }
            }
            try { Activity(_autoPlay.TryOffer(this)); } catch (Exception ex) { Activity("AutoPlay setting could not be changed.", "WARN"); Log(ex.ToString()); }
            Activity(_probe is null ? "Manual Drive Entry is ready. Install smartmontools for automatic scanning." : "Waiting for a USB drive...");
            _guidance.Text = _probe is null ? "Use Manual Drive Entry or Workbook Setup." : "Insert one drive at a time. Use Manual Drive Entry or Workbook Setup at any time.";
            RefreshFooter(); _timer.Start();
        }
        catch (Exception ex) { Log(ex.ToString()); MessageBox.Show(this, ex.Message, "Startup Failed", MessageBoxButtons.OK, MessageBoxIcon.Error); Close(); }
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
                if (_book.HasSerial(record["SerialNumber"])) { Activity($"Disk {disk.Number} duplicate: {record["SerialNumber"]}. No row added.", "WARN"); _guidance.Text = "Remove and insert the next drive, or use Manual Drive Entry."; System.Media.SystemSounds.Exclamation.Play(); }
                else { var row = _book.Add(record); RefreshGrid(); Activity($"USB drive recorded as row {row}: {record["Model"]} / {record["SerialNumber"]}"); _guidance.Text = $"Saved as row {row}. Remove this drive and insert the next."; System.Media.SystemSounds.Asterisk.Play(); }
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            { Activity($"Disk {disk.Number} could not be read: {ex.Message}", "ERROR"); _guidance.Text = "Remove and reinsert to retry, or use Manual Drive Entry."; Log(ex.ToString()); }
        }
        catch (OperationCanceledException) { }
        catch (Exception ex) { Activity("Scan error: " + ex.Message, "ERROR"); Log(ex.ToString()); }
        finally { _busy = false; }
    }
    private void RefreshGrid()
    {
        var sortKey = _grid.SortedColumn?.Name;
        var sortOrder = _grid.SortOrder;
        _grid.Columns.Clear();
        _grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(48, 76, 102);
        _grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
        _grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(48, 76, 102);
        _grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = Color.White;
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "RecordNumber", HeaderText = "#", Width = 54, MinimumWidth = 54, Frozen = true, SortMode = DataGridViewColumnSortMode.NotSortable });
        foreach (var key in _book.Columns)
        {
            var width = key switch { "Make" => 160, "Model" => 260, "SerialNumber" => 210, "Capacity" => 150, "Type" => 200, _ => 175 };
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = key, HeaderText = InventoryBook.Header(key), Width = width, MinimumWidth = 100, SortMode = DataGridViewColumnSortMode.Automatic });
        }
        _grid.Rows.Clear();
        for (var index = 0; index < _book.Records.Count; index++)
        {
            var row = _book.Records[index];
            var displayIndex = _grid.Rows.Add(new object[] { index + 1 }.Concat(_book.Columns.Select(key => (object)row[key])).ToArray());
            _grid.Rows[displayIndex].Tag = index;
        }
        if (sortKey is not null && _grid.Columns.Contains(sortKey))
            _grid.Sort(_grid.Columns[sortKey]!, sortOrder == SortOrder.Descending ? System.ComponentModel.ListSortDirection.Descending : System.ComponentModel.ListSortDirection.Ascending);
        _grid.ClearSelection();
        RefreshFooter();
    }
    private void RefreshFooter()
    {
        _footer.Text = $"Records: {_book.Records.Count}     Workbook: {_book.Path}";
    }
    private void ShowVersionInformation()
    {
        using var dialog = new Form { Text = "Version Information", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.Sizable, ClientSize = new Size(700, 350), MinimumSize = new Size(520, 300) };
        var info = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, TabStop = false, WordWrap = false, ScrollBars = ScrollBars.Both, Font = new Font("Consolas", 10), Lines = [
            $"USB Drive Inventory Collector v{Application.ProductVersion}", "Maintainer: Shannon Wetnight",
            "Repository: https://github.com/ShannonWetnight/usb-drive-inventory-collector",
            $"Workbook: {_book.Path}", $"Debug log: {_logPath}", $"smartctl: {_version}",
            "Scope: USB physical drives; boot and system disks excluded",
            "Transport: smartctl autodetection plus USB adapter fallbacks",
            "Workbook backend: Direct XLSX (no Excel COM)", "Timeout: 30 seconds per smartctl process",
            "Workbook columns: " + string.Join(", ", _book.Columns.Select(InventoryBook.Header))] };
        dialog.Controls.Add(info);
        dialog.Shown += (_, _) => info.Select(0, 0);
        dialog.ShowDialog(this);
    }
    private async Task OpenTerminalAsync()
    {
        if (_terminalSession is not null) { _terminalSession.Stop(); return; }
        if (_busy) { MessageBox.Show(this, "Wait for the current drive read to finish before opening Terminal.", "Terminal", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
        _modal = true;
        _timer.Stop();
        using var io = new EmbeddedTerminalIO();
        _terminalIO = io;
        _terminalSession = TerminalCollector.ForEmbedded(io, _paused, paused =>
        {
            if (IsDisposed || !IsHandleCreated) return;
            BeginInvoke((Action)(() => { if (!IsDisposed) { _paused = paused; _pause.Text = paused ? "Resume Scanning" : "Pause Scanning"; Activity(paused ? "Terminal scanning paused." : "Terminal scanning resumed."); } }));
        });
        var generation = ++_terminalGeneration;
        _terminalOutput.Clear(); _terminalInput.Clear(); _terminalInput.PlaceholderText = "Command or response"; _terminalInput.Enabled = true; _terminalSend.Enabled = true;
        _manual.Enabled = false; _setup.Enabled = false; _grid.Enabled = false;
        _terminalToggle.Text = "Disable Terminal";
        _terminalToggle.FlatStyle = FlatStyle.Flat; _terminalToggle.UseVisualStyleBackColor = false;
        _terminalToggle.BackColor = Color.Firebrick; _terminalToggle.ForeColor = Color.White;
        _tabs.SelectedTab = _terminalPage;
        BeginInvoke((Action)(() => { if (!IsDisposed && ReferenceEquals(io, _terminalIO)) _terminalInput.Focus(); }));
        io.Output += value => WriteTerminalOutput(generation, value);
        io.Cleared += () => ClearTerminalOutput(generation);
        try
        {
            Activity("Terminal opened.");
            await Task.Run(_terminalSession.RunEmbedded);
            if (_closing.IsCancellationRequested) return;
            var updated = new InventoryBook(CollectorSettings.WorkbookPath());
            updated.OpenOrCreate();
            _book = updated;
            RefreshGrid();
            _connected.Clear();
            try { foreach (var disk in DriveProbe.Disks()) _connected.Add(disk.Number); } catch (Exception ex) { Log(ex.ToString()); }
            Activity(_paused ? "Terminal closed. Workbook reloaded. Scanning is paused." : "Terminal closed. Workbook reloaded.");
        }
        catch (OperationCanceledException) { }
        catch (Exception ex) { Log(ex.ToString()); WriteTerminalOutput(generation, $"Terminal error: {ex.Message}{Environment.NewLine}"); MessageBox.Show(this, ex.Message, "Terminal", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        finally
        {
            _terminalSession = null; _terminalIO = null;
            if (!IsDisposed)
            {
                _terminalInput.PlaceholderText = ""; _terminalInput.Text = "Terminal disabled"; _terminalInput.Enabled = false; _terminalSend.Enabled = false; _terminalToggle.Text = "Enable Terminal";
                _manual.Enabled = true; _setup.Enabled = true; _grid.Enabled = true;
                _terminalToggle.FlatStyle = FlatStyle.Standard; _terminalToggle.UseVisualStyleBackColor = true;
                _terminalToggle.BackColor = SystemColors.Control; _terminalToggle.ForeColor = SystemColors.ControlText;
                _terminalToggle.Visible = ReferenceEquals(_tabs.SelectedTab, _terminalPage);
            }
            _modal = false;
            if (!_closing.IsCancellationRequested) _timer.Start();
        }
    }
    private void WriteTerminalOutput(int generation, string value)
    {
        if (IsDisposed || !IsHandleCreated) return;
        try
        {
            if (InvokeRequired) Invoke((Action)Append); else Append();
            void Append() { if (!IsDisposed && generation == _terminalGeneration) { _terminalOutput.AppendText(value); _terminalOutput.ScrollToCaret(); } }
        }
        catch (InvalidOperationException) { }
    }
    private void ClearTerminalOutput(int generation)
    {
        if (IsDisposed || !IsHandleCreated) return;
        try { if (InvokeRequired) Invoke((Action)Clear); else Clear(); void Clear() { if (!IsDisposed && generation == _terminalGeneration) _terminalOutput.Clear(); } }
        catch (InvalidOperationException) { }
    }
    private void Setup()
    {
        if (_busy) { MessageBox.Show(this, "Wait for the current drive read to finish before changing Workbook Setup.", "Workbook Setup"); return; }
        _modal = true;
        try
        {
            using var dialog = new Form { Text = "Workbook Setup", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, ClientSize = new Size(550, 620), MaximizeBox = false, MinimizeBox = false };
            var intro = new Label { Text = "Default columns: Make, Model, Serial Number, Reported Capacity, Type.\nSelect extra identity fields to save for each drive:", Bounds = new Rectangle(20, 12, 510, 55) };
            var list = new CheckedListBox { CheckOnClick = true, Bounds = new Rectangle(20, 75, 510, 320) };
            var optional = InventoryBook.Catalog.Select(c => c.Key).Except(InventoryBook.Core).ToList();
            foreach (var key in optional) list.Items.Add(InventoryBook.Header(key), _book.Columns.Contains(key));
            var note = new Label { Text = "Older and manual records receive N/A for extra fields. A backup keeps any fields removed from the active workbook.", Bounds = new Rectangle(20, 402, 510, 45) };
            var locationLabel = new Label { Text = "Workbook Save Location", Bounds = new Rectangle(20, 456, 270, 24), Font = new Font(Font, FontStyle.Bold) };
            var location = new TextBox { Bounds = new Rectangle(20, 484, 397, 28), Text = _book.Path, ReadOnly = true };
            var browse = Button("Browse...", 425, 482, 105);
            browse.Click += (_, _) =>
            {
                using var pick = new SaveFileDialog { Title = "Workbook Save Location", Filter = "Excel Workbook (*.xlsx)|*.xlsx", DefaultExt = "xlsx", AddExtension = true, OverwritePrompt = false, FileName = Path.GetFileName(location.Text), InitialDirectory = Path.GetDirectoryName(location.Text) };
                if (pick.ShowDialog(dialog) != DialogResult.OK) return;
                if (File.Exists(pick.FileName) && !string.Equals(Path.GetFullPath(pick.FileName), Path.GetFullPath(_book.Path), StringComparison.OrdinalIgnoreCase) &&
                    MessageBox.Show(dialog, "That workbook already exists. Switch to its records instead of copying the current workbook?", "Existing Workbook", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                try
                {
                    if (File.Exists(pick.FileName))
                    {
                        var target = new InventoryBook(pick.FileName);
                        target.OpenOrCreate();
                        for (var i = 0; i < optional.Count; i++) list.SetItemChecked(i, target.Columns.Contains(optional[i]));
                    }
                    location.Text = pick.FileName;
                }
                catch (Exception ex) { MessageBox.Show(dialog, "That workbook could not be opened: " + ex.Message, "Workbook Setup", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            };
            var all = Button("Select All", 20, 558, 112); var defaults = Button("Defaults", 153, 558, 112); var apply = Button("Apply", 286, 558, 112); var cancel = Button("Cancel", 419, 558, 112);
            all.Click += (_, _) => { for (var i = 0; i < list.Items.Count; i++) list.SetItemChecked(i, true); };
            defaults.Click += (_, _) => { for (var i = 0; i < list.Items.Count; i++) list.SetItemChecked(i, false); };
            cancel.Click += (_, _) => dialog.Close();
            apply.Click += (_, _) =>
            {
                var chosen = InventoryBook.Core.Concat(optional.Where((_, i) => list.GetItemChecked(i))).ToList();
                try
                {
                    var selected = _book.AtLocation(location.Text);
                    if (!chosen.SequenceEqual(selected.Columns)) selected.ChangeColumns(chosen);
                    CollectorSettings.SaveWorkbookPath(selected.Path);
                    _book = selected;
                    RefreshGrid(); Activity("Workbook Setup applied."); dialog.Close();
                }
                catch (Exception ex) { Log(ex.ToString()); MessageBox.Show(dialog, "Setup was not applied: " + ex.Message); }
            };
            dialog.Controls.AddRange([intro, list, note, locationLabel, location, browse, all, defaults, apply, cancel]);
            dialog.ShowDialog(this);
        }
        finally { _modal = false; }
    }
    private static Button Button(string text, int x, int y, int width) => new() { Text = text, Bounds = new Rectangle(x, y, width, 34) };
    private void EditRecord(int index)
    {
        if (_busy) { MessageBox.Show(this, "Wait for the current drive read to finish before editing a record.", "Edit Recorded Drive"); return; }
        _modal = true;
        try
        {
            using var dialog = new Form { Text = $"Edit Recorded Drive – Row {index + 2}", StartPosition = FormStartPosition.CenterParent, ClientSize = new Size(650, 570), MinimumSize = new Size(500, 380), Font = new Font("Segoe UI", 10) };
            var intro = new Label { Text = "Edit the saved values below, then save the row. Blank values become N/A.", Dock = DockStyle.Top, Height = 42, Padding = new Padding(18, 11, 0, 0) };
            var keys = _book.Columns.ToArray();
            var fields = new TableLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, ColumnCount = 2, RowCount = keys.Length + 1, Padding = new Padding(18, 4, 18, 4) };
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            var inputs = new Dictionary<string, TextBox>();
            for (var row = 0; row < keys.Length; row++)
            {
                var key = keys[row];
                fields.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
                fields.Controls.Add(new Label { Text = InventoryBook.Header(key), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, row);
                var input = new TextBox { Text = _book.Records[index][key], Dock = DockStyle.Fill, Margin = new Padding(0, 4, 2, 4) };
                fields.Controls.Add(input, 1, row);
                inputs[key] = input;
            }
            fields.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var actions = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 55, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(12, 7, 15, 0) };
            var cancel = new Button { Text = "Cancel", Width = 95, Height = 34 };
            var save = new Button { Text = "Save Changes", Width = 140, Height = 34 };
            cancel.Click += (_, _) => dialog.Close();
            save.Click += (_, _) =>
            {
                try
                {
                    var candidate = _book.Records[index].Copy();
                    foreach (var (key, input) in inputs) candidate[key] = input.Text;
                    candidate = ManualValidation.NormalizeEdited(candidate, _book.Columns);
                    if (_book.HasSerial(candidate["SerialNumber"], index)) throw new InvalidOperationException("This serial number is already in the workbook.");
                    if (MessageBox.Show(dialog, $"Save changes to row {index + 2}?", "Review Changes", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                    _book.Update(index, candidate);
                    RefreshGrid();
                    _grid.ClearSelection();
                    foreach (DataGridViewRow row in _grid.Rows) if (row.Tag is int recordIndex && recordIndex == index) { row.Selected = true; _grid.FirstDisplayedScrollingRowIndex = row.Index; break; }
                    Activity($"Row {index + 2} updated: {candidate["Model"]} / {candidate["SerialNumber"]}");
                    dialog.Close();
                }
                catch (Exception ex) { Log(ex.ToString()); MessageBox.Show(dialog, ex.Message, "Edit Recorded Drive", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            };
            actions.Controls.AddRange([cancel, save]);
            dialog.Controls.Add(fields); dialog.Controls.Add(actions); dialog.Controls.Add(intro);
            dialog.ShowDialog(this);
        }
        finally { _modal = false; }
    }
    private void ManualEntry(bool copyLast)
    {
        if (_busy) { MessageBox.Show(this, "Wait for the current drive read to finish before entering a manual record.", "Manual Drive Entry"); return; }
        _modal = true;
        try
        {
            using var dialog = new Form { Text = "Manual Drive Entry", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, ClientSize = new Size(560, 410), MaximizeBox = false, MinimizeBox = false, Font = new Font("Segoe UI", 10) };
            dialog.Controls.Add(new Label { Text = "Leave a field blank for N/A. Model and serial are saved in uppercase.", Bounds = new Rectangle(20, 12, 520, 30) });
            var labels = new[] { "1. Make", "2. Model", "3. Serial Number", "4. Capacity (number only)" };
            for (int i = 0; i < labels.Length; i++) dialog.Controls.Add(new Label { Text = labels[i], Bounds = new Rectangle(20, 47 + 43 * i, 195, 26) });
            var make = Box(219, 47, 320); var model = Box(219, 90, 320); var serial = Box(219, 133, 320); var amount = Box(219, 176, 175);
            model.CharacterCasing = CharacterCasing.Upper; serial.CharacterCasing = CharacterCasing.Upper;
            var unit = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Bounds = new Rectangle(405, 176, 134, 28), AccessibleName = "Capacity Unit" };
            _toolTip.SetToolTip(unit, "Capacity Unit");
            unit.Items.AddRange(["N/A", "B", "KB", "MB", "GB", "TB", "PB", "Other"]); unit.SelectedIndex = 0;
            var customUnitLabel = new Label { Text = "Custom Capacity Unit", Bounds = new Rectangle(20, 219, 195, 26) };
            var customUnit = Box(219, 219, 320);
            var typeLabel = new Label { Text = "5. Drive Type", Bounds = new Rectangle(20, 219, 195, 26) };
            var type = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown, Bounds = new Rectangle(219, 219, 320, 28), MaxDropDownItems = 12 };
            var types = DriveTypes.Options;
            type.Items.AddRange(types.Cast<object>().ToArray());
            type.SelectedIndex = 0;
            var customTypeLabel = new Label { Text = "Custom Drive Type", Bounds = new Rectangle(20, 262, 195, 26) };
            var customType = Box(219, 262, 320);
            void UpdateManualLayout()
            {
                var otherUnit = unit.Text == "Other";
                customUnit.Visible = customUnitLabel.Visible = otherUnit;
                type.Top = typeLabel.Top = otherUnit ? 262 : 219;
                customType.Top = customTypeLabel.Top = type.Top + 43;
                customType.Visible = customTypeLabel.Visible = type.Text.Equals("Other", StringComparison.OrdinalIgnoreCase);
            }
            unit.SelectedIndexChanged += (_, _) => UpdateManualLayout();
            type.SelectedIndexChanged += (_, _) => UpdateManualLayout();
            type.TextUpdate += (_, _) =>
            {
                var query = type.Text; var caret = type.SelectionStart;
                type.BeginUpdate(); type.Items.Clear(); type.Items.AddRange(DriveTypes.Matches(query).Cast<object>().ToArray()); type.EndUpdate();
                type.Text = query; type.SelectionStart = caret; type.DroppedDown = true;
                UpdateManualLayout();
            };
            dialog.Controls.AddRange([make, model, serial, amount, unit, customUnitLabel, customUnit, typeLabel, type, customTypeLabel, customType]);
            UpdateManualLayout();
            if (copyLast && _book.Records.LastOrDefault() is { } last) { Fill(last); serial.Clear(); }
            var review = Button("Review Drive", 20, 355, 130); var copy = Button("Copy Last Drive", 162, 355, 160); var back = Button("Return to Scanning", 334, 355, 205);
            copy.Enabled = _book.Records.Count > 0;
            copy.Click += (_, _) => { if (_book.Records.LastOrDefault() is { } saved) { Fill(saved); serial.Clear(); serial.Focus(); } };
            back.Click += (_, _) => dialog.Close();
            review.Click += (_, _) =>
            {
                try
                {
                    var choice = type.Text.Trim();
                    if (!types.Contains(choice, StringComparer.OrdinalIgnoreCase)) throw new ArgumentException("Select a Drive Type from the list, or choose Other.");
                    var record = ManualValidation.Create(make.Text, model.Text, serial.Text, amount.Text, unit.Text == "N/A" ? "" : unit.Text, customUnit.Text, choice, customType.Text);
                    var duplicate = _book.HasSerial(record["SerialNumber"]);
                    var action = Review(record, duplicate);
                    if (action == "Serial") { serial.Focus(); serial.SelectAll(); return; }
                    if (action == "Cancel") { dialog.Close(); return; }
                    if (action == "Edit") return;
                    var row = _book.Add(record); RefreshGrid(); Activity($"Manual drive recorded as row {row}: {record["Model"]} / {record["SerialNumber"]}");
                    MessageBox.Show(dialog, $"Drive saved as row {row}.", "Drive Recorded", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    copy.Enabled = true;
                    serial.Clear();
                    if (action == "Save") { make.Clear(); model.Clear(); amount.Clear(); unit.SelectedIndex = 0; customUnit.Clear(); type.Items.Clear(); type.Items.AddRange(types.Cast<object>().ToArray()); type.SelectedIndex = 0; customType.Clear(); dialog.Text = "Manual Drive Entry"; }
                    else dialog.Text = "Manual Drive Entry – Copy Saved Drive";
                    serial.Focus();
                }
                catch (Exception ex) { Log(ex.ToString()); MessageBox.Show(dialog, ex.Message, "Manual Drive Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            };
            dialog.Controls.AddRange([review, copy, back]); dialog.ShowDialog(this);
            void Fill(DriveRecord record)
            {
                make.Text = record["Make"] == "N/A" ? "" : record["Make"];
                model.Text = record["Model"] == "N/A" ? "" : record["Model"];
                amount.Clear(); unit.SelectedIndex = 0; customUnit.Clear(); customType.Clear();
                var match = Regex.Match(record["Capacity"], @"^([0-9]+(?:\.[0-9]+)?)\s+(.+)$");
                if (match.Success) { amount.Text = match.Groups[1].Value; var value = match.Groups[2].Value; if (unit.Items.Contains(value)) unit.SelectedItem = value; else { unit.SelectedItem = "Other"; customUnit.Text = value; } }
                var media = record["Type"];
                type.Items.Clear(); type.Items.AddRange(types.Cast<object>().ToArray());
                if (types.Contains(media, StringComparer.OrdinalIgnoreCase)) type.Text = media;
                else { type.Text = "Other"; customType.Text = media; }
                UpdateManualLayout();
            }
        }
        finally { _modal = false; }
    }
    private static TextBox Box(int x, int y, int width) => new() { Bounds = new Rectangle(x, y, width, 28) };
    private string Review(DriveRecord record, bool duplicate)
    {
        using var dialog = new Form { Text = duplicate ? "Duplicate Serial Number" : "Review Manual Drive", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, ClientSize = new Size(490, 345), MaximizeBox = false, MinimizeBox = false };
        var summary = new TextBox { Bounds = new Rectangle(20, 20, 450, 158), Multiline = true, ReadOnly = true, Font = new Font("Consolas", 11), Lines = [$"Make:     {record["Make"]}", $"Model:    {record["Model"]}", $"Serial:   {record["SerialNumber"]}", $"Capacity: {record["Capacity"]}", $"Type:     {record["Type"]}"] };
        var note = new Label { Text = duplicate ? "This serial is already in the workbook. Change it or cancel this record." : record["SerialNumber"] == "N/A" ? "Serial N/A cannot be checked for duplicates." : "Review these values before saving a new row.", Bounds = new Rectangle(20, 188, 450, 48) };
        dialog.Controls.AddRange([summary, note]); string action = "Edit";
        void Choice(string name, string text, int x, int width) { var button = Button(text, x, 272, width); button.Click += (_, _) => { action = name; dialog.Close(); }; dialog.Controls.Add(button); }
        if (duplicate) { Choice("Serial", "Change Serial", 20, 140); Choice("Cancel", "Cancel Record", 180, 140); }
        else { Choice("Save", "Save & Next", 14, 105); Choice("Copy", "Save & Copy", 125, 105); Choice("Edit", "Edit Fields", 236, 105); Choice("Cancel", "Cancel Record", 347, 128); }
        dialog.ShowDialog(this); return action;
    }
}

internal static class DriveTypes
{
    private static readonly string[] Known = [
        "1.8-inch SATA SSD", "2.5-inch IDE HDD", "2.5-inch SAS HDD", "2.5-inch SAS SSD", "2.5-inch SATA HDD", "2.5-inch SATA SSD",
        "3.5-inch Floppy Disk", "3.5-inch IDE HDD", "3.5-inch SAS HDD", "3.5-inch SAS SSD", "3.5-inch SATA HDD",
        "5.25-inch Floppy Disk", "CompactFlash Card", "eMMC", "HDD", "IDE Drive", "IDE HDD", "IDE SSD",
        "M.2 NVMe SSD", "M.2 SATA SSD", "microSD Card", "mSATA SSD", "NVMe SSD", "PATA Drive", "PATA HDD", "PATA SSD",
        "SAS HDD", "SAS SSD", "SATA Drive", "SATA HDD", "SATA SSD", "SD Card", "SSD", "USB Flash Drive"
    ];
    public static readonly string[] Options = ["N/A", ..Known.OrderBy(s => s, StringComparer.OrdinalIgnoreCase), "Other"];

    public static IEnumerable<string> Matches(string query)
    {
        query = query.Trim();
        if (query.Length == 0) return Options;
        return Options.Where(option => option != "Other" &&
            (option.Contains(query, StringComparison.OrdinalIgnoreCase) || IsSubsequence(option, query)))
            .Append("Other");
    }

    private static bool IsSubsequence(string value, string query)
    {
        var index = 0;
        foreach (var c in value) if (index < query.Length && char.ToUpperInvariant(c) == char.ToUpperInvariant(query[index])) index++;
        return index == query.Length;
    }
}

internal static class ManualValidation
{
    public static DriveRecord NormalizeEdited(DriveRecord edited, IEnumerable<string> columns)
    {
        var capacity = edited["Capacity"];
        var match = Regex.Match(capacity, @"\A([0-9]{1,15}(?:\.[0-9]{1,6})?)\s+([A-Za-z]{1,12})\z");
        if (capacity != "N/A" && !match.Success) throw new ArgumentException("Reported Capacity must be a number followed by a unit, such as 0.005 GB.");
        var unit = match.Success ? match.Groups[2].Value : "";
        var listed = new[] { "B", "KB", "MB", "GB", "TB", "PB" }.Contains(unit);
        var type = edited["Type"];
        var known = DriveTypes.Options.Contains(type, StringComparer.OrdinalIgnoreCase);
        var normalized = Create(edited["Make"] == "N/A" ? "" : edited["Make"], edited["Model"] == "N/A" ? "" : edited["Model"],
            edited["SerialNumber"] == "N/A" ? "" : edited["SerialNumber"], match.Success ? match.Groups[1].Value : "",
            listed ? unit : match.Success ? "Other" : "", listed ? "" : unit, known ? type : "Other", known ? "" : type);
        foreach (var key in columns.Except(InventoryBook.Core))
        {
            var value = edited[key];
            if (value.Length > 200 || value.Any(c => char.IsControl(c) || c < 32)) throw new ArgumentException($"{InventoryBook.Header(key)} contains unsupported characters or is too long.");
            normalized[key] = value;
        }
        return normalized;
    }
    public static DriveRecord Create(string make, string model, string serial, string amount, string unit, string customUnit, string type, string customType)
    {
        make = make.Trim(); model = model.Trim(); serial = serial.Trim(); amount = amount.Trim(); unit = unit.Trim(); customUnit = customUnit.Trim(); type = type.Trim(); customType = customType.Trim();
        Check(make, @"\A[A-Za-z0-9][A-Za-z0-9 .&()+'/_-]{0,79}\z", 80, "Make");
        Check(model, "\\A[A-Za-z0-9][A-Za-z0-9 .+/_\"-]{0,99}\\z", 100, "Model");
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

using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace USBDriveInventoryCollector;

internal sealed class CollectorForm : Form
{
    private enum StatusTone { Default, Reading, Success, Warning, Error, Paused }
    private enum DriveNotification { Saved, Duplicate, Error, TerminalEnabled, TerminalDisabled }
    private static readonly Color NeutralStatusBackground = Color.FromArgb(246, 248, 251);
    private InventoryBook _book;
    private readonly AutoPlayGuard _autoPlay = new();
    private string _logPath;
    private DriveProbe? _probe;
    private readonly HashSet<int> _connected = [];
    private readonly CancellationTokenSource _closing = new();
    private CancellationTokenSource? _activeScan;
    private readonly System.Windows.Forms.Timer _timer = new() { Interval = 1000 };
    private readonly System.Windows.Forms.Timer _statusPulse = new() { Interval = 700 };
    private readonly System.Windows.Forms.Timer _statusDisplay = new() { Interval = 4000 };
    private readonly System.Windows.Forms.Timer _statusFade = new() { Interval = 25 };
    private readonly object _soundGate = new();
    private readonly HashSet<CancellationTokenSource> _pendingSounds = [];
    private DateTime _nextSoundAt;
    private Panel? _statusPanel;
    private Color _fadeFrom = NeutralStatusBackground, _fadeTo = NeutralStatusBackground;
    private long _fadeStart;
    private bool _pulseBright, _soundsEnabled = CollectorSettings.SoundsEnabled();
    private StatusTone _statusTone;
    private DateTime _scanHoldUntil;
    private bool _busy, _identifying, _paused, _modal, _terminalStarting, _updatingGrid, _viewChanged;
    private string _version = "N/A";
    private readonly RichTextBox _status = new() { Dock = DockStyle.Fill, ReadOnly = true, TabStop = false, BorderStyle = BorderStyle.None, ScrollBars = RichTextBoxScrollBars.None, BackColor = NeutralStatusBackground, Font = new Font("Segoe UI", 12, FontStyle.Bold), Text = "Initializing..." };
    private readonly Label _guidance = new() { Dock = DockStyle.Fill, Text = "Insert one drive at a time." };
    private readonly Label _recordCount = new() { AutoSize = false, UseMnemonic = false, TextAlign = ContentAlignment.MiddleLeft };
    private readonly Label _workbookLocation = new() { AutoEllipsis = true, AutoSize = false, UseMnemonic = false, TextAlign = ContentAlignment.MiddleLeft };
    private readonly Button _copyPath = new() { Text = "Copy Path", Size = new Size(95, 28) };
    private readonly Button _openPath = new() { Text = "Open Folder", Size = new Size(110, 28) };
    private readonly DataGridView _grid = new() { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None, SelectionMode = DataGridViewSelectionMode.FullRowSelect, EnableHeadersVisualStyles = false, ColumnHeadersHeight = 34 };
    private readonly ListBox _activity = new() { Dock = DockStyle.Fill, HorizontalScrollbar = true, Font = new Font("Consolas", 9) };
    private readonly CollectorTabControl _tabs = new() { Dock = DockStyle.Fill, SizeMode = TabSizeMode.Fixed, ItemSize = new Size(130, 36) };
    private readonly RichTextBox _terminalOutput = new() { Dock = DockStyle.Fill, ReadOnly = true, TabStop = false, BackColor = Color.FromArgb(18, 22, 28), ForeColor = Color.Gainsboro, Font = new Font("Consolas", 10), BorderStyle = BorderStyle.None };
    private readonly TextBox _terminalInput = new() { Font = new Font("Consolas", 10), Text = "Terminal disabled", BorderStyle = BorderStyle.None, Enabled = false };
    private readonly Panel _terminalInputFrame = new() { BorderStyle = BorderStyle.FixedSingle, BackColor = SystemColors.Control };
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
    private readonly Button _sound = new() { Width = 36, Height = 34, Image = CreateSoundIcon(true), ImageAlign = ContentAlignment.MiddleCenter, Padding = Padding.Empty, FlatStyle = FlatStyle.Flat, ForeColor = Color.White, BackColor = Color.FromArgb(48, 76, 102), TextAlign = ContentAlignment.MiddleCenter };
    private readonly Button _theme = new() { Width = 36, Height = 34, ImageAlign = ContentAlignment.MiddleCenter, Padding = Padding.Empty, FlatStyle = FlatStyle.Flat, ForeColor = Color.White, BackColor = Color.FromArgb(48, 76, 102), AccessibleName = "Toggle Theme" };
    private readonly Panel _tabActions = new();
    private ThemePreference _themePreference = CollectorSettings.Theme();
    private bool _dark;
    private readonly Button _terminalToggle = new() { Text = "Enable Terminal", Width = 140, Height = 24, Visible = false };
    private readonly Button _resetView = new() { Text = "Reset View", Width = 110, Height = 28, Visible = false };
    private readonly Button _refreshWorkbook = new() { Image = CreateRefreshIcon(), ImageAlign = ContentAlignment.MiddleCenter, AccessibleName = "Refresh Workbook", Width = 34, Height = 28, Visible = true };
    private readonly Button _finish = new() { Text = "Finish", Width = 90 };
    public CollectorForm(bool initialize = true)
    {
        Text = "USB Drive Inventory Collector";
        KeyPreview = true;
        MinimumSize = new Size(900, 590); Size = new Size(1030, 720); StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 10); BackColor = Color.FromArgb(246, 248, 251);
        Directory.CreateDirectory(CollectorSettings.LogsDirectory());
        _book = new InventoryBook(CollectorSettings.WorkbookPath());
        _logPath = Path.Combine(CollectorSettings.LogsDirectory(), $"USB-Drive-Inventory-Collector-{DateTime.Now:yyyyMMdd-HHmmss}.log");
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 5, ColumnCount = 1 };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        foreach (var height in new[] { 82f, 56f, 112f }) layout.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
        Controls.Add(layout);
        var header = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty, BackColor = Color.FromArgb(32, 55, 78), Tag = CollectorTheme.PreserveColors };
        var title = new Label { Text = "USB Drive Inventory Collector", ForeColor = Color.White, Font = new Font("Segoe UI", 18, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft };
        var subtitle = new Label { Text = "One drive at a time. Every record is saved immediately.", ForeColor = Color.FromArgb(215, 229, 238), TextAlign = ContentAlignment.MiddleLeft };
        var info = new Button { Text = "Version Information", AccessibleName = "Version Information", Size = new Size(160, 34), FlatStyle = FlatStyle.Flat, ForeColor = Color.White, BackColor = Color.FromArgb(48, 76, 102), Font = new Font("Segoe UI", 9), TextAlign = ContentAlignment.MiddleCenter, TabStop = true };
        info.Click += (_, _) => ShowVersionInformation();
        _toolTip.SetToolTip(info, "Version Information");
        _setup.Height = 34; _setup.FlatStyle = FlatStyle.Flat; _setup.ForeColor = Color.White; _setup.BackColor = Color.FromArgb(48, 76, 102); _setup.Font = new Font("Segoe UI", 9); _setup.TextAlign = ContentAlignment.MiddleCenter;
        var headerActions = new Panel { Dock = DockStyle.Right, Width = 432, BackColor = header.BackColor };
        headerActions.Controls.AddRange([_setup, info, _sound, _theme]);
        void PositionHeaderActions()
        {
            var top = (headerActions.ClientSize.Height - 34) / 2;
            _theme.Location = new Point(headerActions.ClientSize.Width - _theme.Width - 14, top);
            _sound.Location = new Point(_theme.Left - _sound.Width - 8, top);
            info.Location = new Point(_sound.Left - info.Width - 8, top);
            _setup.Location = new Point(info.Left - _setup.Width - 8, top);
        }
        headerActions.Resize += (_, _) => PositionHeaderActions(); PositionHeaderActions();
        void UpdateSoundButton()
        {
            var oldIcon = _sound.Image;
            _sound.Image = CreateSoundIcon(_soundsEnabled);
            oldIcon?.Dispose();
            _sound.Text = string.Empty;
            _sound.AccessibleName = _soundsEnabled ? "Mute Sounds" : "Enable Sounds";
            _toolTip.SetToolTip(_sound, _soundsEnabled ? "Mute drive notification sounds" : "Enable drive notification sounds");
        }
        UpdateSoundButton();
        _sound.Click += (_, _) =>
        {
            try { CollectorSettings.SaveSoundsEnabled(!_soundsEnabled); _soundsEnabled = !_soundsEnabled; if (!_soundsEnabled) CancelSounds(); UpdateSoundButton(); }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "Sound Preference", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        };
        _theme.Click += (_, _) =>
        {
            try
            {
                var preference = OppositeTheme(_dark);
                CollectorSettings.SaveTheme(preference);
                _themePreference = preference;
                ApplyTheme();
            }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "Theme Preference", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        };
        header.Controls.Add(subtitle); header.Controls.Add(title); header.Controls.Add(headerActions);
        header.Resize += (_, _) =>
        {
            var width = Math.Max(1, header.ClientSize.Width - headerActions.Width - 28);
            title.SetBounds(20, 7, width, 37);
            subtitle.SetBounds(22, 40, width - 2, 25);
        };
        layout.Controls.Add(header, 0, 0);
        var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, Margin = Padding.Empty, Padding = new Padding(12, 13, 0, 0) };
        _finish.Width = _setup.Width;
        foreach (var b in new[] { _pause, _manual, _finish }) { b.Height = 34; b.Margin = new Padding(0, 0, 8, 0); actions.Controls.Add(b); }
        layout.Controls.Add(actions, 0, 1);
        var status = new Panel { Dock = DockStyle.Fill, Margin = new Padding(12, 4, 12, 4), BorderStyle = BorderStyle.Fixed3D, BackColor = NeutralStatusBackground };
        _statusPanel = status;
        _status.Dock = DockStyle.None; _guidance.Dock = DockStyle.None;
        _guidance.Padding = Padding.Empty; _guidance.TextAlign = ContentAlignment.MiddleLeft;
        status.Controls.Add(_status); status.Controls.Add(_guidance);
        void PositionStatus()
        {
            var width = Math.Max(1, status.ClientSize.Width - 44);
            var measured = TextRenderer.MeasureText(_status.Text, _status.Font, new Size(width, 1000), TextFormatFlags.WordBreak);
            var statusHeight = Math.Clamp(measured.Height, 24, 70);
            const int guidanceHeight = 24, gap = 2;
            var top = Math.Max(0, (status.ClientSize.Height - statusHeight - gap - guidanceHeight) / 2);
            _status.SetBounds(22, top, width, statusHeight);
            _guidance.SetBounds(20, top + statusHeight + gap, width + 2, guidanceHeight);
        }
        status.Resize += (_, _) => PositionStatus();
        _status.TextChanged += (_, _) => PositionStatus();
        layout.Controls.Add(status, 0, 2);
        var recordsPage = new TabPage("Recorded Drives") { Controls = { _grid } };
        _tabs.TabPages.Add(recordsPage);
        _tabs.TabPages.Add(new TabPage("Activity") { Controls = { _activity } });
        _tabs.DrawMode = TabDrawMode.OwnerDrawFixed;
        _tabs.DrawItem += (_, e) =>
        {
            var selected = e.Index == _tabs.SelectedIndex;
            using var brush = new SolidBrush(selected ? CollectorTheme.Field(_dark) : CollectorTheme.Surface(_dark));
            e.Graphics.FillRectangle(brush, e.Bounds);
            TextRenderer.DrawText(e.Graphics, _tabs.TabPages[e.Index].Text, _tabs.Font, e.Bounds, CollectorTheme.Text(_dark), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            if ((e.State & DrawItemState.Focus) != 0) ControlPaint.DrawFocusRectangle(e.Graphics, e.Bounds, CollectorTheme.Text(_dark), brush.Color);
        };
        var terminalEntry = new Panel { Dock = DockStyle.Bottom, Height = 44, BackColor = Color.FromArgb(230, 234, 240) };
        _terminalInputFrame.Controls.Add(_terminalInput);
        _terminalInputFrame.Resize += (_, _) => _terminalInput.SetBounds(5, (_terminalInputFrame.ClientSize.Height - _terminalInput.Height) / 2, Math.Max(1, _terminalInputFrame.ClientSize.Width - 10), _terminalInput.Height);
        terminalEntry.Controls.Add(_terminalInputFrame); terminalEntry.Controls.Add(_terminalSend);
        terminalEntry.Resize += (_, _) =>
        {
            var top = (terminalEntry.ClientSize.Height - 28) / 2;
            _terminalSend.Location = new Point(terminalEntry.ClientSize.Width - _terminalSend.Width - 8, top);
            _terminalInputFrame.SetBounds(8, top, Math.Max(80, _terminalSend.Left - 16), 28);
        };
        _terminalPage.Controls.Add(_terminalOutput); _terminalPage.Controls.Add(terminalEntry); _tabs.TabPages.Add(_terminalPage);
        var tabHost = new Panel { Dock = DockStyle.Fill, Margin = new Padding(12, 3, 12, 3) };
        // Keep an opaque strip present even when its buttons are hidden. Native
        // TabControl painting otherwise leaves pale rectangles behind the overlays.
        _tabActions.Controls.AddRange([_terminalToggle, _resetView, _refreshWorkbook]);
        _tabs.HeaderActions = _tabActions;
        tabHost.Controls.Add(_tabs); tabHost.Controls.Add(_tabActions);
        _tabActions.BringToFront();
        void PositionTerminalToggle()
        {
            if (!_tabs.IsHandleCreated) return;
            var tab = _tabs.GetTabRect(_tabs.TabPages.IndexOf(_terminalPage));
            _tabActions.SetBounds(tab.Right + 1, tab.Top, Math.Max(0, tabHost.ClientSize.Width - tab.Right - 2), tab.Height);
            _terminalToggle.Height = 28;
            var top = (tab.Height - _terminalToggle.Height) / 2;
            _terminalToggle.Location = new Point(7, top);
            _refreshWorkbook.Location = new Point(_tabActions.ClientSize.Width - _refreshWorkbook.Width - 4, top);
            _resetView.Location = new Point(_refreshWorkbook.Left - _resetView.Width - 8, top);
            _tabActions.BringToFront();
        }
        tabHost.Resize += (_, _) => PositionTerminalToggle();
        _tabs.HandleCreated += (_, _) => PositionTerminalToggle();
        Shown += (_, _) => { PositionTerminalToggle(); UpdateResetViewButton(); };
        layout.Controls.Add(tabHost, 0, 3);
        var footer = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty };
        footer.Controls.AddRange([_recordCount, _workbookLocation, _copyPath, _openPath]);
        footer.Resize += (_, _) =>
        {
            const int gap = 8;
            var top = (footer.ClientSize.Height - 28) / 2;
            _openPath.Location = new Point(footer.ClientSize.Width - _openPath.Width - 14, top);
            _copyPath.Location = new Point(_openPath.Left - _copyPath.Width - gap, top);
            _recordCount.SetBounds(18, top, 110, 28);
            var workbookLeft = _recordCount.Right + 4;
            _workbookLocation.SetBounds(workbookLeft, top, Math.Max(0, _copyPath.Left - gap - workbookLeft), 28);
        };
        _toolTip.SetToolTip(_workbookLocation, "Workbook Save Location"); _toolTip.SetToolTip(_copyPath, "Copy workbook location to clipboard"); _toolTip.SetToolTip(_openPath, "Open the workbook folder in File Explorer");
        layout.Controls.Add(footer, 0, 4);
        _pause.Click += (_, _) => SetScanningPaused(!_paused);
        _manual.Click += (_, _) => ManualEntry(false);
        _setup.Click += (_, _) => Setup();
        _terminalToggle.Click += (_, _) => { if (_terminalSession is null) _terminalTask = OpenTerminalAsync(); else _terminalSession.Stop(); };
        _tabs.SelectedIndexChanged += (_, _) =>
        {
            UpdateTabActions(_terminalSession is not null);
            PositionTerminalToggle();
            if (_terminalToggle.Visible) _terminalToggle.BringToFront();
            if (ReferenceEquals(_tabs.SelectedTab, recordsPage)) BeginInvoke((Action)(() => { if (!IsDisposed && ReferenceEquals(_tabs.SelectedTab, recordsPage)) _grid.Focus(); }));
        };
        _resetView.Click += (_, _) => ResetRecordView();
        _refreshWorkbook.Click += (_, _) => ReloadWorkbook();
        _toolTip.SetToolTip(_refreshWorkbook, "Refresh Workbook from disk");
        _grid.ColumnWidthChanged += (_, _) => { EnsureHeaderHeight(); MarkViewChanged(); };
        _grid.ColumnHeadersHeightChanged += (_, _) => { EnsureHeaderHeight(); MarkViewChanged(); };
        _grid.ColumnDisplayIndexChanged += (_, _) => MarkViewChanged();
        _grid.RowHeightChanged += (_, _) => MarkViewChanged();
        _grid.Sorted += (_, _) => MarkViewChanged();
        _grid.Scroll += (_, _) => MarkViewChanged();
        _terminalInput.KeyDown += (_, e) => { if (e.KeyCode != Keys.Enter || _terminalIO is null) return; e.SuppressKeyPress = true; SubmitTerminalInput(); };
        _terminalSend.Click += (_, _) => SubmitTerminalInput();
        KeyPress += TerminalKeyPress;
        _copyPath.Click += async (_, _) => await CopyWorkbookPathAsync(_book.Path, _copyPath, this);
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
        _statusPulse.Tick += (_, _) =>
        {
            _pulseBright = !_pulseBright;
            ColorizeStatus();
        };
        _statusFade.Tick += (_, _) => AdvanceStatusFade();
        _statusDisplay.Tick += (_, _) =>
        {
            _statusDisplay.Stop();
            if (_busy || _modal) { _statusDisplay.Start(); return; }
            Activity(_paused ? "Scanning is paused." : _terminalSession is null ? "Waiting for a USB drive..." : "Terminal opened.");
        };
        if (initialize) Shown += async (_, _) => await InitializeAsync();
        FormClosing += (_, _) => { _timer.Stop(); _statusPulse.Stop(); _statusFade.Stop(); _statusDisplay.Stop(); _closing.Cancel(); CancelSounds(); _terminalSession?.Stop(); try { _autoPlay.Dispose(); } catch (Exception ex) { Log(ex.ToString()); MessageBox.Show(this, "AutoPlay could not be restored. Check Windows AutoPlay settings."); } };
        HandleCreated += (_, _) => ApplyTheme();
        SystemEvents.UserPreferenceChanged += SystemThemeChanged;
        Disposed += (_, _) => { SystemEvents.UserPreferenceChanged -= SystemThemeChanged; _theme.Image?.Dispose(); };
        ApplyTheme();
    }
    private void SystemThemeChanged(object? sender, UserPreferenceChangedEventArgs args)
    {
        if (IsDisposed || !IsHandleCreated) return;
        try { BeginInvoke((Action)(() => { if (!IsDisposed) ApplyTheme(); })); }
        catch (InvalidOperationException) { }
    }
    private void ApplyTheme()
    {
        _dark = CollectorTheme.IsDark(_themePreference);
        CollectorTheme.Apply(this, _dark);
        CollectorTheme.ApplyTitleBar(this, _dark);
        foreach (Form dialog in OwnedForms) { CollectorTheme.Apply(dialog, _dark); CollectorTheme.ApplyTitleBar(dialog, _dark); }
        _terminalOutput.BackColor = Color.FromArgb(18, 22, 28); _terminalOutput.ForeColor = Color.Gainsboro;
        SetTerminalIndicator(_terminalSession is not null);
        var oldIcon = _theme.Image; _theme.Image = CreateThemeIcon(_dark); oldIcon?.Dispose();
        oldIcon = _refreshWorkbook.Image; _refreshWorkbook.Image = CreateRefreshIcon(_dark); oldIcon?.Dispose();
        var next = OppositeTheme(_dark);
        _theme.AccessibleName = $"Switch to {next} Mode";
        _toolTip.SetToolTip(_theme, $"Switch to {next} mode" + (_themePreference == ThemePreference.System ? " (currently following Windows)" : ""));
        _statusFade.Stop(); _fadeTo = _fadeFrom = CollectorTheme.Surface(_dark);
        if (_statusPanel is not null) _statusPanel.BackColor = _fadeTo;
        _status.BackColor = _guidance.BackColor = _fadeTo;
        FadeStatusBackground(_statusTone); ColorizeStatus();
    }
    private static ThemePreference OppositeTheme(bool dark) => dark ? ThemePreference.Light : ThemePreference.Dark;
    private DialogResult ShowThemedDialog(Form dialog)
    {
        CollectorTheme.Apply(dialog, _dark);
        EventHandler updateTitle = (_, _) => CollectorTheme.ApplyTitleBar(dialog, _dark);
        dialog.HandleCreated += updateTitle;
        try
        {
            return dialog.ShowDialog(this);
        }
        finally { dialog.HandleCreated -= updateTitle; }
    }
    private static Bitmap CreateThemeIcon(bool dark)
    {
        var icon = new Bitmap(18, 18);
        using var graphics = Graphics.FromImage(icon);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        using var pen = new Pen(Color.White, 1.5f);
        if (dark)
        {
            using var light = new SolidBrush(Color.White);
            using var cutout = new SolidBrush(Color.FromArgb(48, 76, 102));
            graphics.FillEllipse(light, 2, 2, 13, 13); graphics.FillEllipse(cutout, 7, 0, 12, 12);
        }
        else
        {
            graphics.DrawEllipse(pen, 5, 5, 8, 8);
            for (var i = 0; i < 8; i++)
            {
                var angle = i * Math.PI / 4;
                graphics.DrawLine(pen, 9 + (float)Math.Cos(angle) * 6, 9 + (float)Math.Sin(angle) * 6, 9 + (float)Math.Cos(angle) * 8, 9 + (float)Math.Sin(angle) * 8);
            }
        }
        return icon;
    }
    internal static void VerifyThemes(string directory)
    {
        using var form = new CollectorForm(initialize: false) { ShowInTaskbar = false };
        form.Show();
        Application.DoEvents();
        form._book.Records.Add(new DriveRecord { ["Manufacturer"] = "Example", ["Model"] = "THEME-CHECK", ["SerialNumber"] = "THEME-SERIAL" });
        form.RefreshGrid();
        var savedPreference = CollectorSettings.Theme();
        foreach (var preference in new[] { ThemePreference.Dark, ThemePreference.Light, ThemePreference.System })
        {
            form._themePreference = preference; form.ApplyTheme();
            form.PerformLayout();
            Application.DoEvents();
            if (form._theme.Size != form._sound.Size || form._theme.Left - form._sound.Right != 8 ||
                form._grid.DefaultCellStyle.BackColor != CollectorTheme.Field(form._dark) ||
                form._grid.Rows[0].Cells["SerialNumber"].Value?.ToString() != "THEME-SERIAL")
                throw new InvalidDataException("Theme changed header spacing, grid colors, or recorded values.");
            using var dialog = new Form();
            var input = new TextBox { Text = "Unchanged", Dock = DockStyle.Top };
            var link = new LinkLabel { Text = "Attributions", Dock = DockStyle.Top };
            var summary = new RichTextBox { Text = "Review entry", Dock = DockStyle.Fill };
            dialog.Controls.AddRange([input, link, summary]);
            CollectorTheme.Apply(dialog, form._dark);
            if (input.Text != "Unchanged" || input.BackColor != CollectorTheme.Field(form._dark) ||
                summary.BackColor != input.BackColor || link.LinkColor == input.BackColor)
                throw new InvalidDataException("Theme changed dialog values or made links unreadable.");
            var statusPanel = form._statusPanel!;
            var workbookHost = form._tabs.Parent!;
            var banner = form._theme.Parent!.Parent!;
            var statusLeft = statusPanel.Parent!.PointToScreen(statusPanel.Location).X;
            if (form._pause.PointToScreen(Point.Empty).X != statusLeft ||
                workbookHost.PointToScreen(Point.Empty).X != statusLeft ||
                workbookHost.Width != statusPanel.Width ||
                form._manual.Left - form._pause.Right != form._pause.Margin.Right || form._finish.Left - form._manual.Right != form._manual.Margin.Right ||
                banner.Margin != Padding.Empty)
                throw new InvalidDataException($"Inconsistent margins: pause={form._pause.PointToScreen(Point.Empty).X}, status={statusLeft}, tabs={workbookHost.PointToScreen(Point.Empty).X}, widths={workbookHost.Width}/{statusPanel.Width}.");
            SavePreview(form, directory, preference.ToString());
            form._viewChanged = true; form.UpdateResetViewButton();
            form._tabs.SelectedIndex = 1;
            VerifyHiddenActions(form, terminalEnabled: false);
            SavePreview(form, directory, $"{preference}-Activity");
            form._tabs.SelectedTab = form._terminalPage;
            VerifyHiddenActions(form, terminalEnabled: false);
            SavePreview(form, directory, $"{preference}-Terminal");
            // Reproduce switching away while Terminal is active, then disabling it.
            form.UpdateTabActions(terminalEnabled: true);
            form._tabs.SelectedIndex = 1;
            form.UpdateTabActions(terminalEnabled: true);
            form.UpdateTabActions(terminalEnabled: false);
            VerifyHiddenActions(form, terminalEnabled: false);
            SavePreview(form, directory, $"{preference}-Terminal-Disabled");
            form._tabs.SelectedIndex = 0;
            form.UpdateTabActions(terminalEnabled: true);
            form.UpdateTabActions(terminalEnabled: false);
            if (!form._refreshWorkbook.Visible || !form._resetView.Visible || form._terminalToggle.Visible)
                throw new InvalidDataException("Recorded Drives lost its view actions after Terminal stopped.");
            form._viewChanged = false; form.UpdateResetViewButton();
            form.ManualEntry(false, manual =>
            {
                var unit = manual.Controls.OfType<ThemedComboBox>().Single(c => c.AccessibleName == "Capacity Unit");
                unit.SelectedItem = "GB";
                // Keep unattended verification independent of native popup focus
                // and desktop input, which can close the list or select another item.
                if (unit.Text != "GB" || unit.SelectedItem?.ToString() != "GB")
                    throw new InvalidDataException($"Capacity selection changed: text='{unit.Text}', selected='{unit.SelectedItem}', theme={preference}.");
                if (unit.ForeColor == unit.BackColor)
                    throw new InvalidDataException($"Capacity dropdown text has no contrast in {preference} theme.");
                SavePreview(manual, directory, $"{preference}-Manual");
                unit.SelectedItem = "Other";
                if (!manual.Controls.OfType<Label>().Single(c => c.Text == "Custom Capacity Unit").Visible)
                    throw new InvalidDataException("Custom capacity unit field did not appear.");
            });
            foreach (var tone in new[] { StatusTone.Paused, StatusTone.Success })
            {
                form._statusTone = tone;
                form._status.Text = tone == StatusTone.Paused ? "Scanning is paused." : "Manual drive recorded.";
                form.FadeStatusBackground(tone);
                form._fadeStart = Stopwatch.GetTimestamp() - Stopwatch.Frequency;
                form.AdvanceStatusFade(); form.ColorizeStatus();
                SavePreview(form, directory, $"{preference}-Status-{tone}");
            }
            form._statusTone = StatusTone.Default; form._status.Text = "Initializing..."; form.ApplyTheme();
        }
        if (CollectorSettings.Theme() != savedPreference) throw new InvalidDataException("Theme preview changed the saved preference.");
        using var io = new EmbeddedTerminalIO();
        form._terminalIO = io; form._tabs.SelectedTab = form._terminalPage;
        form.TerminalKeyPress(form, new KeyPressEventArgs('r'));
        if (!io.KeyAvailable || io.ReadKey() != 'R') throw new InvalidDataException("Embedded Terminal did not accept the removal shortcut.");
        form._terminalIO = null;
    }
    private static void SavePreview(Form form, string directory, string name)
    {
        Application.DoEvents();
        using var bitmap = new Bitmap(form.Width, form.Height);
        form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, form.Size));
        bitmap.Save(Path.Combine(directory, $"Collector-Theme-{name}.png"));
    }
    private static void VerifyHiddenActions(CollectorForm form, bool terminalEnabled)
    {
        form.UpdateTabActions(terminalEnabled);
        Application.DoEvents();
        if (form._refreshWorkbook.Visible || form._resetView.Visible ||
            form._terminalToggle.Visible != (terminalEnabled || form._tabs.SelectedTab == form._terminalPage))
            throw new InvalidDataException("A tab showed actions that do not belong on it.");
        using var bitmap = new Bitmap(form._tabActions.Width, form._tabActions.Height);
        form._tabActions.DrawToBitmap(bitmap, form._tabActions.ClientRectangle);
        foreach (var button in new[] { form._refreshWorkbook, form._resetView, form._terminalToggle }.Where(b => !b.Visible))
        {
            var bounds = Rectangle.Intersect(button.Bounds, form._tabActions.ClientRectangle);
            for (var x = bounds.Left + 2; x < bounds.Right - 2; x++)
                for (var y = bounds.Top + 2; y < bounds.Bottom - 2; y++)
                    if (bitmap.GetPixel(x, y).ToArgb() != form._tabActions.BackColor.ToArgb())
                        throw new InvalidDataException("Hidden tab action left an unpainted rectangle.");
        }
    }
    private void Log(string text) { try { File.AppendAllText(_logPath, $"{DateTime.Now:O} {text}\n"); } catch { } }
    private static Bitmap CreateRefreshIcon(bool dark = false)
    {
        var icon = new Bitmap(16, 16);
        using var graphics = Graphics.FromImage(icon);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        using var pen = new Pen(CollectorTheme.Text(dark), 1.8f);
        graphics.DrawArc(pen, 2.5f, 2.5f, 11f, 11f, 45f, 285f);
        using var arrow = new SolidBrush(CollectorTheme.Text(dark));
        graphics.FillPolygon(arrow, [new PointF(14, 5), new PointF(10, 3), new PointF(11, 8)]);
        return icon;
    }
    private static Bitmap CreateSoundIcon(bool enabled)
    {
        var icon = new Bitmap(20, 20);
        using var graphics = Graphics.FromImage(icon);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        graphics.Clear(Color.Transparent);
        using var fill = new SolidBrush(Color.White);
        using var pen = new Pen(Color.White, 1.7f) { StartCap = LineCap.Round, EndCap = LineCap.Round };
        using var speaker = new GraphicsPath();
        speaker.AddPolygon([new PointF(2, 7), new PointF(6, 7), new PointF(11, 3), new PointF(11, 17), new PointF(6, 13), new PointF(2, 13)]);
        graphics.FillPath(fill, speaker);
        if (enabled)
        {
            graphics.DrawArc(pen, 8, 5, 8, 10, -65, 130);
            graphics.DrawArc(pen, 8, 2, 10, 16, -60, 120);
        }
        else
        {
            graphics.DrawLine(pen, 13, 7, 18, 13);
            graphics.DrawLine(pen, 18, 7, 13, 13);
        }
        return icon;
    }
    private void SetScanningPaused(bool paused)
    {
        _paused = paused;
        if (paused) CancelSounds();
        _pause.Text = paused ? "Resume Scanning" : "Pause Scanning";
        _terminalSession?.SetPaused(paused);
        if (_terminalSession is not null) WriteTerminalOutput(_terminalGeneration, paused ? "Scanning paused from the main window. Press [P] or Resume Scanning to continue.\n" : "Scanning resumed from the main window.\n");
        Activity(paused ? (_terminalSession is null ? "Scanning is paused." : "Terminal scanning paused.") : "Scanning resumed.");
    }
    private void Activity(string text, string level = "INFO", StatusTone tone = StatusTone.Default)
    {
        _statusPulse.Stop(); _statusDisplay.Stop();
        _statusTone = tone == StatusTone.Default ? level == "ERROR" ? StatusTone.Error : level == "WARN" && text.Contains("duplicate", StringComparison.OrdinalIgnoreCase) ? StatusTone.Warning : _paused ? StatusTone.Paused : StatusTone.Default : tone;
        FadeStatusBackground(_statusTone);
        var display = _terminalSession is not null && !text.StartsWith("Terminal closed.", StringComparison.Ordinal) &&
            !text.Contains("Scanning is paused", StringComparison.Ordinal)
            ? text + (_paused ? " Scanning is paused in Terminal and this window." : " Scanning is paused in this window.") : text;
        if (_paused && _terminalSession is null && !display.Contains("Scanning is paused", StringComparison.Ordinal))
            display += " Scanning is paused.";
        _status.Text = display;
        _pulseBright = false;
        ColorizeStatus();
        if (_statusTone is StatusTone.Reading or StatusTone.Warning or StatusTone.Error) _statusPulse.Start();
        if (_statusTone is StatusTone.Success or StatusTone.Warning or StatusTone.Error) _statusDisplay.Start();
        _activity.Items.Insert(0, $"{DateTime.Now:HH:mm:ss}  {level,-5}  {text}");
        if (_activity.Items.Count > 500) _activity.Items.RemoveAt(500);
        Log($"{level} {text}");
    }
    private void ColorizeStatus()
    {
        _status.SelectAll();
        _status.SelectionColor = _dark ? _statusTone switch
        {
            StatusTone.Reading => _pulseBright ? Color.FromArgb(174, 214, 255) : Color.FromArgb(132, 193, 255),
            StatusTone.Error or StatusTone.Paused => Color.FromArgb(255, 161, 161),
            StatusTone.Warning => _pulseBright ? Color.FromArgb(255, 222, 144) : Color.FromArgb(245, 205, 112),
            StatusTone.Success => Color.FromArgb(146, 224, 182),
            _ => CollectorTheme.Text(true)
        } : _statusTone switch
        {
            StatusTone.Reading => _pulseBright ? Color.FromArgb(66, 111, 151) : Color.FromArgb(30, 69, 110),
            StatusTone.Error => _pulseBright ? Color.FromArgb(195, 73, 73) : Color.Firebrick,
            StatusTone.Paused => Color.Firebrick,
            StatusTone.Warning => _pulseBright ? Color.FromArgb(178, 128, 14) : Color.FromArgb(123, 86, 0),
            StatusTone.Success => Color.FromArgb(24, 111, 68),
            _ => Color.FromArgb(15, 20, 25)
        };
        const string paused = "Scanning is paused";
        var pausedAt = _status.Text.IndexOf(paused, StringComparison.Ordinal);
        if (pausedAt >= 0)
        {
            var sentenceEnd = _status.Text.IndexOf('.', pausedAt);
            _status.Select(pausedAt, sentenceEnd >= 0 ? sentenceEnd - pausedAt + 1 : paused.Length);
            _status.SelectionColor = _dark ? Color.FromArgb(255, 161, 161) : Color.Firebrick;
        }
        _status.Select(0, 0);
    }
    private void FadeStatusBackground(StatusTone tone)
    {
        var target = _dark ? tone switch
        {
            StatusTone.Reading => Color.FromArgb(35, 53, 72),
            StatusTone.Warning => Color.FromArgb(66, 53, 29),
            StatusTone.Error or StatusTone.Paused => Color.FromArgb(70, 35, 42),
            StatusTone.Success => Color.FromArgb(31, 58, 46),
            _ => CollectorTheme.Surface(true)
        } : tone switch
        {
            StatusTone.Reading => Color.FromArgb(222, 234, 246),
            StatusTone.Warning => Color.FromArgb(250, 240, 209),
            StatusTone.Error => Color.FromArgb(250, 226, 227),
            StatusTone.Paused => Color.FromArgb(250, 226, 227),
            StatusTone.Success => Color.FromArgb(224, 243, 231),
            _ => NeutralStatusBackground
        };
        if (target == _fadeTo) return; // Text pulsing must not restart the background transition.
        _fadeFrom = _statusPanel?.BackColor ?? NeutralStatusBackground;
        _fadeTo = target;
        _fadeStart = Stopwatch.GetTimestamp();
        _statusFade.Start();
    }
    private void AdvanceStatusFade()
    {
        if (_statusPanel is null) return;
        var progress = Math.Clamp(Stopwatch.GetElapsedTime(_fadeStart).TotalMilliseconds / 300.0, 0, 1);
        var eased = progress < 0.5 ? 4 * progress * progress * progress : 1 - Math.Pow(-2 * progress + 2, 3) / 2;
        byte Channel(byte a, byte b) => (byte)Math.Round(a + (b - a) * eased);
        var color = Color.FromArgb(Channel(_fadeFrom.R, _fadeTo.R), Channel(_fadeFrom.G, _fadeTo.G), Channel(_fadeFrom.B, _fadeTo.B));
        _statusPanel.BackColor = color;
        _status.BackColor = _guidance.BackColor = color;
        if (progress >= 1) _statusFade.Stop();
    }
    private void CancelSounds()
    {
        lock (_soundGate) foreach (var source in _pendingSounds) source.Cancel();
    }
    private void PlayDriveNotification(DriveNotification notification)
    {
        var terminalConfirmation = notification is DriveNotification.TerminalEnabled or DriveNotification.TerminalDisabled;
        if (!_soundsEnabled || _closing.IsCancellationRequested || (_paused && !terminalConfirmation)) return;
        // One sound sequence at a time. Drop repeat events within one second rather
        // than queueing beeps that might play after a pause or a burst of errors.
        var tones = notification switch
        {
            DriveNotification.Saved => new[] { (1000, 150), (1200, 150) },
            DriveNotification.Duplicate => new[] { (500, 400) },
            DriveNotification.Error => new[] { (400, 220), (300, 320) },
            DriveNotification.TerminalEnabled => new[] { (750, 140), (1050, 170) },
            DriveNotification.TerminalDisabled => new[] { (1050, 140), (750, 170) },
            _ => new[] { (800, 150) }
        };
        CancellationTokenSource source;
        TimeSpan wait;
        lock (_soundGate)
        {
            var now = DateTime.UtcNow;
            if (now < _nextSoundAt && !terminalConfirmation) return;
            var startAt = now < _nextSoundAt ? _nextSoundAt : now;
            wait = startAt - now;
            source = new CancellationTokenSource();
            _pendingSounds.Add(source);
            _nextSoundAt = startAt.AddSeconds(1);
        }
        _ = Task.Run(async () =>
        {
            try
            {
                if (wait > TimeSpan.Zero) await Task.Delay(wait, source.Token);
                foreach (var (frequency, duration) in tones)
                {
                    source.Token.ThrowIfCancellationRequested();
                    Console.Beep(frequency, duration);
                    await Task.Delay(120, source.Token);
                }
            }
            catch (OperationCanceledException) { }
            catch (Exception ex) { Log("Sound playback failed: " + ex.Message); }
            finally
            {
                lock (_soundGate) _pendingSounds.Remove(source);
                source.Dispose();
            }
        });
    }
    private static async Task CopyWorkbookPathAsync(string path, Button button, IWin32Window owner)
    {
        try
        {
            Clipboard.SetText(path);
            var marker = new object(); button.Tag = marker; button.Text = "Copied";
            await Task.Delay(1000);
            if (!button.IsDisposed && ReferenceEquals(button.Tag, marker)) button.Text = "Copy Path";
        }
        catch (Exception ex) { MessageBox.Show(owner, ex.Message, "Copy Workbook Location"); }
    }
    private static void OpenWorkbookFolder(string path, IWin32Window owner)
    {
        try { Process.Start(new ProcessStartInfo("explorer.exe") { UseShellExecute = true, Arguments = File.Exists(path) ? $"/select,\"{path}\"" : $"\"{Path.GetDirectoryName(path)}\"" }); }
        catch (Exception ex) { MessageBox.Show(owner, ex.Message, "Open Workbook Folder"); }
    }
    private static Panel CenteredField(TextBox input, Rectangle bounds)
    {
        var frame = new Panel { Bounds = bounds, BorderStyle = BorderStyle.FixedSingle, BackColor = SystemColors.Window };
        input.BorderStyle = BorderStyle.None;
        frame.Controls.Add(input);
        frame.Resize += (_, _) => input.SetBounds(5, (frame.ClientSize.Height - input.Height) / 2, Math.Max(1, frame.ClientSize.Width - 10), input.Height);
        input.SetBounds(5, (frame.ClientSize.Height - input.Height) / 2, Math.Max(1, frame.ClientSize.Width - 10), input.Height);
        return frame;
    }
    private void ShowFinishDialog()
    {
        using var dialog = new Form { Text = "Inventory Saved", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, ClientSize = new Size(560, 160), MaximizeBox = false, MinimizeBox = false, Font = new Font("Segoe UI", 10) };
        var path = _book.Path;
        var message = new Label { Text = "Workbook saved at:", Bounds = new Rectangle(20, 16, 520, 25) };
        var location = new TextBox { Text = path, ReadOnly = true, TabStop = false };
        var locationFrame = CenteredField(location, new Rectangle(20, 46, 520, 28));
        var copy = Button("Copy Path", 20, 102, 160);
        var open = Button("Open Folder", 200, 102, 160);
        var close = Button("Close", 380, 102, 160);
        _toolTip.SetToolTip(location, path);
        copy.Click += async (_, _) => await CopyWorkbookPathAsync(path, copy, dialog);
        open.Click += (_, _) => OpenWorkbookFolder(path, dialog);
        close.Click += (_, _) => { dialog.DialogResult = DialogResult.OK; dialog.Close(); };
        dialog.AcceptButton = close;
        dialog.Controls.AddRange([message, locationFrame, copy, open, close]);
        if (ShowThemedDialog(dialog) == DialogResult.OK) Close();
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
        if ("MLSRPDHQ".Contains(command))
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
            await Task.Delay(3000, _closing.Token);
            Activity(_probe is null ? "Manual Drive Entry is ready. Install smartmontools for automatic scanning." : "Waiting for a USB drive...");
            _guidance.Text = _probe is null ? "Use Manual Drive Entry or Workbook Setup." : "Insert one drive at a time. Use Manual Drive Entry or Workbook Setup at any time.";
            RefreshFooter(); _timer.Start();
        }
        catch (OperationCanceledException) when (_closing.IsCancellationRequested) { }
        catch (Exception ex) { Log(ex.ToString()); MessageBox.Show(this, ex.Message, "Startup Failed", MessageBoxButtons.OK, MessageBoxIcon.Error); Close(); }
    }
    private async Task PollAsync()
    {
        if (_busy || _paused || _modal || _probe is null || _closing.IsCancellationRequested || DateTime.UtcNow < _scanHoldUntil) return;
        _busy = true;
        using var scan = CancellationTokenSource.CreateLinkedTokenSource(_closing.Token);
        _activeScan = scan;
        try
        {
            var disks = await Task.Run(DriveProbe.Disks, scan.Token).WaitAsync(scan.Token);
            scan.Token.ThrowIfCancellationRequested();
            var numbers = disks.Select(d => d.Number).ToHashSet();
            foreach (var old in _connected.Where(n => !numbers.Contains(n)).ToList()) { _connected.Remove(old); _probe.Forget(old); Activity($"Disk {old} removed. Ready for another drive."); }
            var disk = disks.FirstOrDefault(d => !_connected.Contains(d.Number));
            if (disk is null) return;
            _connected.Add(disk.Number);
            _identifying = true;
            Activity($"USB drive detected on Disk {disk.Number}. Reading drive identity...", tone: StatusTone.Reading);
            _guidance.Text = "Reading drive identity. Slow adapters may reach the 30-second timeout.";
            await Task.Delay(2000, scan.Token);
            DriveRecord record;
            try
            {
                record = await _probe.IdentifyAsync(disk, scan.Token);
                scan.Token.ThrowIfCancellationRequested();
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                Activity($"Disk {disk.Number} identity could not be read: {ex.Message}", "ERROR");
                _guidance.Text = "Remove and reinsert to retry, or use Manual Drive Entry.";
                PlayDriveNotification(DriveNotification.Error);
                _scanHoldUntil = DateTime.UtcNow.AddSeconds(4);
                Log(ex.ToString());
                return;
            }
            try
            {
                if (_book.HasSerial(record["SerialNumber"])) { Activity($"Disk {disk.Number} duplicate: {record["SerialNumber"]}. No row added.", "WARN"); _guidance.Text = "Remove and insert the next drive, or use Manual Drive Entry."; PlayDriveNotification(DriveNotification.Duplicate); }
                else { var row = _book.Add(record); RefreshGrid(latestNumberAtTop: true); Activity($"USB drive recorded as row {row}: {record["Model"]} / {record["SerialNumber"]}", tone: StatusTone.Success); _guidance.Text = $"Saved as row {row}. Remove this drive and insert the next."; PlayDriveNotification(DriveNotification.Saved); }
                _scanHoldUntil = DateTime.UtcNow.AddSeconds(4);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                Activity($"Disk {disk.Number} was identified, but the workbook could not be updated ({ex.GetType().Name}): {ex.Message}", "ERROR");
                _guidance.Text = "Check the workbook save location and file access before continuing.";
                PlayDriveNotification(DriveNotification.Error);
                _scanHoldUntil = DateTime.UtcNow.AddSeconds(4);
                Log(ex.ToString());
            }
        }
        catch (OperationCanceledException) { }
        catch (Exception ex) { Activity("Scan error: " + ex.Message, "ERROR"); _scanHoldUntil = DateTime.UtcNow.AddSeconds(4); PlayDriveNotification(DriveNotification.Error); Log(ex.ToString()); }
        finally { _identifying = false; _activeScan = null; _busy = false; }
    }
    private void RefreshGrid(bool preserveSort = true, bool latestNumberAtTop = false)
    {
        var wasUpdating = _updatingGrid;
        _updatingGrid = true;
        try
        {
            var sortKey = preserveSort ? _grid.SortedColumn?.Name : null;
            var sortOrder = _grid.SortOrder;
            if (latestNumberAtTop && sortKey == "RecordNumber") sortOrder = SortOrder.Descending;
            _grid.Columns.Clear();
            _grid.ColumnHeadersHeight = 34;
            _grid.RowTemplate.Height = MinimumRowHeight();
            _grid.RowTemplate.MinimumHeight = _grid.RowTemplate.Height;
            _grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(48, 76, 102);
            _grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(48, 76, 102);
            _grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = Color.White;
            _grid.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "RecordNumber", HeaderText = "#", Width = 54, MinimumWidth = 54, Frozen = true, ValueType = typeof(int), SortMode = DataGridViewColumnSortMode.Automatic });
            foreach (var key in _book.Columns)
            {
                var width = key switch { "Manufacturer" => 175, "Model" => 260, "SerialNumber" => 210, "Capacity" => 150, "Type" => 200, _ => 175 };
                _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = key, HeaderText = InventoryBook.Header(key), Width = width, MinimumWidth = 100, SortMode = DataGridViewColumnSortMode.Automatic });
            }
            EnsureHeaderHeight();
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
            CollectorTheme.Apply(_grid, _dark);
            RefreshFooter();
            _viewChanged = sortKey is not null;
        }
        finally { _updatingGrid = wasUpdating; if (!wasUpdating) UpdateResetViewButton(); }
    }
    private int MinimumRowHeight()
        => Math.Max(26, TextRenderer.MeasureText("Ag", _grid.DefaultCellStyle.Font ?? _grid.Font).Height + 8);
    private void EnsureHeaderHeight()
    {
        var font = _grid.ColumnHeadersDefaultCellStyle.Font ?? _grid.Font;
        var minimum = 34;
        foreach (DataGridViewColumn column in _grid.Columns)
        {
            var availableWidth = Math.Max(24, column.Width - 28);
            var textHeight = TextRenderer.MeasureText(column.HeaderText, font, new Size(availableWidth, 1000), TextFormatFlags.WordBreak).Height;
            minimum = Math.Max(minimum, textHeight + 10);
        }
        if (_grid.ColumnHeadersHeight < minimum) _grid.ColumnHeadersHeight = minimum;
    }
    private void ResetRecordView()
    {
        _updatingGrid = true;
        try
        {
            RefreshGrid(preserveSort: false);
            _grid.HorizontalScrollingOffset = 0;
            if (_grid.Rows.Count > 0) _grid.FirstDisplayedScrollingRowIndex = 0;
        }
        finally { _updatingGrid = false; _viewChanged = false; UpdateResetViewButton(); }
        _grid.Focus();
    }
    private void MarkViewChanged()
    {
        if (_updatingGrid || !_grid.IsHandleCreated || _tabs.SelectedIndex != 0) return;
        _viewChanged = true;
        UpdateResetViewButton();
    }
    private void UpdateTabActions(bool terminalEnabled)
    {
        _terminalToggle.Visible = ReferenceEquals(_tabs.SelectedTab, _terminalPage) || terminalEnabled;
        UpdateResetViewButton();
    }
    private void UpdateResetViewButton()
    {
        _resetView.Visible = _viewChanged && _tabs.SelectedIndex == 0;
        _refreshWorkbook.Visible = _tabs.SelectedIndex == 0;
        if (_resetView.Visible) _resetView.BringToFront();
        if (_refreshWorkbook.Visible) _refreshWorkbook.BringToFront();
        _tabActions.Invalidate(true);
    }
    private void ReloadWorkbook()
    {
        if (_busy || _modal || _terminalSession is not null)
        {
            MessageBox.Show(this, "Wait for the current operation to finish before refreshing the workbook.", "Refresh Workbook", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        try
        {
            if (!File.Exists(_book.Path)) throw new FileNotFoundException("The workbook was not found at its current location.", _book.Path);
            var updated = new InventoryBook(_book.Path);
            updated.OpenOrCreate();
            _book = updated;
            RefreshGrid();
            Activity("Workbook refreshed from disk.");
        }
        catch (Exception ex)
        {
            Log(ex.ToString());
            MessageBox.Show(this, ex.Message, "Refresh Workbook", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
    private void RefreshFooter()
    {
        _recordCount.Text = $"Records: {_book.Records.Count}";
        _workbookLocation.Text = $"Workbook: {_book.Path}";
    }
    private void ShowVersionInformation()
    {
        using var dialog = new Form { Text = "Version Information", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.Sizable, ClientSize = new Size(700, 380), MinimumSize = new Size(520, 320) };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(16, 12, 16, 12) };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
        var displayVersion = Application.ProductVersion.Split('+', 2)[0];
        var title = new Label { Text = $"USB Drive Inventory Collector v{displayVersion}", Dock = DockStyle.Fill, Margin = Padding.Empty, Padding = new Padding(2, 0, 0, 0), TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Consolas", 10, FontStyle.Bold) };
        var maintainer = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, Margin = Padding.Empty, Padding = new Padding(2, 0, 0, 0) };
        maintainer.Controls.Add(new Label { Text = "Maintainer: Shannon Wetnight |", AutoSize = true, Font = new Font("Consolas", 10), Margin = new Padding(0, 4, 7, 0) });
        var website = new LinkLabel { Text = "shannonwetnight.com", AutoSize = true, Font = new Font("Consolas", 10), Margin = new Padding(0, 4, 0, 0) };
        maintainer.Controls.Add(website);
        var executableName = Path.GetFileName(Environment.ProcessPath ?? "USB-Drive-Inventory-Collector.exe");
        var info = new RichTextBox { Dock = DockStyle.Fill, Margin = Padding.Empty, BorderStyle = BorderStyle.Fixed3D, ReadOnly = true, TabStop = false, DetectUrls = true, WordWrap = false, ScrollBars = RichTextBoxScrollBars.Both, Font = new Font("Consolas", 10), Lines = [
            $"Executable: {executableName}",
            "License: MIT",
            "Copyright (c) 2026 Shannon Wetnight",
            "License details: https://github.com/ShannonWetnight/usb-drive-inventory-collector/blob/main/LICENSE",
            "Repository: https://github.com/ShannonWetnight/usb-drive-inventory-collector",
            $"Workbook: {_book.Path}", $"Debug log: {_logPath}", $"smartctl: {_version}",
            "Scope: USB physical drives with media; boot and system disks excluded",
            "Transport: smartctl autodetection plus USB adapter fallbacks",
            "Workbook backend: Direct XLSX (no Excel COM)", "Timeout: 30 seconds per smartctl process",
            "Workbook columns: " + string.Join(", ", _book.Columns.Select(InventoryBook.Header)),
            "", "Attributions:", "",
            "Open XML SDK 3.3.0 — .NET Foundation and Contributors (MIT)",
            "https://github.com/dotnet/Open-XML-SDK", "",
            ".NET 8 / Windows Desktop — .NET Foundation and Contributors",
            "MIT source license; Microsoft Windows binary terms also apply",
            "Includes System.Management, System.Text.Encoding.CodePages, System.IO.Packaging, and System.CodeDom",
            "https://github.com/dotnet/runtime",
            "https://github.com/dotnet/winforms",
            "https://github.com/dotnet/wpf",
            "Windows binary terms: https://github.com/dotnet/core/blob/main/license-information-windows.md", "",
            "smartmontools / smartctl — smartmontools developers (GPL-2.0-or-later)",
            "https://www.smartmontools.org/",
            "smartctl is installed separately and is called as an external program.", "",
            "Full license text and attributions: LICENSE beside the executable",
            "https://github.com/ShannonWetnight/usb-drive-inventory-collector/blob/main/LICENSE"] };
        var disclaimer = new Label { Text = "AI Workflow Notice: This project was written through AI prompting and reviewed by its maintainer. Check collected data against the drive label when accuracy matters.", Dock = DockStyle.Fill, Margin = Padding.Empty, ForeColor = Color.DimGray, Font = new Font("Segoe UI", 9), Padding = new Padding(2, 7, 0, 0) };
        void OpenLink(string? url)
        {
            if (string.IsNullOrWhiteSpace(url) || !Uri.TryCreate(url, UriKind.Absolute, out var uri) || uri.Scheme != Uri.UriSchemeHttps) return;
            try { Process.Start(new ProcessStartInfo(uri.AbsoluteUri) { UseShellExecute = true }); }
            catch (Exception ex) { MessageBox.Show(dialog, ex.Message, "Open Link", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }
        website.LinkClicked += (_, _) => OpenLink("https://shannonwetnight.com");
        info.LinkClicked += (_, e) => OpenLink(e.LinkText);
        layout.Controls.Add(title, 0, 0); layout.Controls.Add(maintainer, 0, 1); layout.Controls.Add(info, 0, 2); layout.Controls.Add(disclaimer, 0, 3);
        dialog.Controls.Add(layout);
        dialog.Shown += (_, _) => info.Select(0, 0);
        ShowThemedDialog(dialog);
    }
    private void SetTerminalIndicator(bool enabled)
    {
        _terminalToggle.Text = enabled ? "Disable Terminal" : "Enable Terminal";
        _terminalToggle.AccessibleName = enabled ? "Disable Terminal" : "Enable Terminal";
        _terminalToggle.FlatStyle = enabled || _dark ? FlatStyle.Flat : FlatStyle.Standard;
        _terminalToggle.UseVisualStyleBackColor = !enabled && !_dark;
        _terminalToggle.BackColor = enabled ? Color.Firebrick : _dark ? Color.FromArgb(48, 55, 65) : SystemColors.Control;
        _terminalToggle.ForeColor = enabled ? Color.White : CollectorTheme.Text(_dark);
        _toolTip.SetToolTip(_terminalToggle, enabled ? "Terminal is enabled. Click to close it." : "Terminal is disabled. Click to open it.");
    }
    private async Task OpenTerminalAsync()
    {
        if (_terminalSession is not null) { _terminalSession.Stop(); return; }
        if (_terminalStarting) return;
        _terminalStarting = true;
        _modal = true;
        _timer.Stop();
        _terminalToggle.Enabled = false;
        try
        {
            // A routine disk check can be abandoned; an active identity read must finish
            // so its record is saved before the Terminal takes over scanning.
            if (_busy && !_identifying) _activeScan?.Cancel();
            while (_busy) await Task.Delay(50, _closing.Token);
            if (_closing.IsCancellationRequested) { _modal = false; return; }
        }
        catch (OperationCanceledException) when (_closing.IsCancellationRequested)
        {
            _modal = false;
            return;
        }
        finally
        {
            _terminalStarting = false;
            if (!IsDisposed) _terminalToggle.Enabled = true;
        }
        using var io = new EmbeddedTerminalIO();
        _terminalIO = io;
        _terminalSession = TerminalCollector.ForEmbedded(io, _paused, paused =>
        {
            if (IsDisposed || !IsHandleCreated) return;
            BeginInvoke((Action)(() => { if (!IsDisposed) { _paused = paused; if (paused) CancelSounds(); _pause.Text = paused ? "Resume Scanning" : "Pause Scanning"; Activity(paused ? "Terminal scanning paused." : "Terminal scanning resumed."); } }));
        });
        var generation = ++_terminalGeneration;
        _terminalOutput.Clear(); _terminalInput.Clear(); _terminalInput.PlaceholderText = "Command or response"; _terminalInput.Enabled = true; _terminalInputFrame.BackColor = CollectorTheme.Field(_dark); _terminalSend.Enabled = true;
        _manual.Enabled = false; _setup.Enabled = false;
        SetTerminalIndicator(true);
        _tabs.SelectedTab = _terminalPage;
        BeginInvoke((Action)(() => { if (!IsDisposed && ReferenceEquals(io, _terminalIO)) _terminalInput.Focus(); }));
        io.Output += value => WriteTerminalOutput(generation, value);
        io.Cleared += () => ClearTerminalOutput(generation);
        try
        {
            PlayDriveNotification(DriveNotification.TerminalEnabled);
            Activity("Terminal opened.");
            await Task.Run(_terminalSession.RunEmbedded);
            if (_closing.IsCancellationRequested) return;
            var updated = new InventoryBook(CollectorSettings.WorkbookPath());
            updated.OpenOrCreate();
            _book = updated;
            RefreshGrid();
            foreach (var old in _connected) _probe?.Forget(old);
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
                _terminalInput.PlaceholderText = ""; _terminalInput.Text = "Terminal disabled"; _terminalInput.Enabled = false; _terminalInputFrame.BackColor = CollectorTheme.Surface(_dark); _terminalSend.Enabled = false;
                _manual.Enabled = true; _setup.Enabled = true;
                SetTerminalIndicator(false);
                UpdateTabActions(terminalEnabled: false);
            }
            _modal = false;
            if (!_closing.IsCancellationRequested) PlayDriveNotification(DriveNotification.TerminalDisabled);
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
        if (_terminalSession is not null) { MessageBox.Show(this, "Disable Terminal before changing Workbook Setup.", "Workbook Setup"); return; }
        _modal = true;
        try
        {
            using var dialog = new Form { Text = "Workbook Setup", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, ClientSize = new Size(550, 620), MaximizeBox = false, MinimizeBox = false };
            var intro = new Label { Text = "Default columns: Manufacturer, Model, Serial Number, Reported Capacity, Type.\nSelect extra identity fields to save for each drive:", Bounds = new Rectangle(20, 12, 510, 55) };
            var list = new CheckedListBox { CheckOnClick = true, Bounds = new Rectangle(20, 70, 510, 265) };
            var optional = InventoryBook.Catalog.Select(c => c.Key).Except(InventoryBook.Core).ToList();
            foreach (var key in optional) list.Items.Add(InventoryBook.Header(key), _book.Columns.Contains(key));
            var note = new Label { Text = "Older and manual records receive N/A for extra fields. A backup keeps any fields removed from the active workbook.", Bounds = new Rectangle(20, 342, 510, 43) };
            var locationLabel = new Label { Text = "Workbook Save Location", Bounds = new Rectangle(20, 392, 270, 24), Font = new Font(Font, FontStyle.Bold) };
            var location = new TextBox { Text = _book.Path, ReadOnly = true };
            var locationFrame = CenteredField(location, new Rectangle(20, 418, 397, 28));
            var browse = Button("Browse...", 425, 418, 105); browse.Height = 28;
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
            var logsLabel = new Label { Text = "Logs Save Location", Bounds = new Rectangle(20, 462, 270, 24), Font = new Font(Font, FontStyle.Bold) };
            var logsLocation = new TextBox { Text = CollectorSettings.LogsDirectory(), ReadOnly = true };
            var logsFrame = CenteredField(logsLocation, new Rectangle(20, 488, 397, 28));
            var browseLogs = Button("Browse...", 425, 488, 105); browseLogs.Height = 28;
            browseLogs.Click += (_, _) =>
            {
                using var pick = new FolderBrowserDialog { Description = "Choose where collector logs are saved", SelectedPath = logsLocation.Text, ShowNewFolderButton = true };
                if (pick.ShowDialog(dialog) == DialogResult.OK) logsLocation.Text = pick.SelectedPath;
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
                    var targetLogs = Path.GetFullPath(logsLocation.Text);
                    Directory.CreateDirectory(targetLogs);
                    var newLogPath = Path.Combine(targetLogs, Path.GetFileName(_logPath));
                    if (!string.Equals(newLogPath, _logPath, StringComparison.OrdinalIgnoreCase))
                    {
                        if (File.Exists(newLogPath)) throw new IOException("A log with this name already exists in the selected folder.");
                        if (File.Exists(_logPath)) File.Move(_logPath, newLogPath);
                        _logPath = newLogPath;
                    }
                    CollectorSettings.SavePaths(selected.Path, targetLogs);
                    _book = selected;
                    RefreshGrid(); Activity("Workbook Setup applied."); dialog.Close();
                }
                catch (Exception ex) { Log(ex.ToString()); MessageBox.Show(dialog, "Setup was not applied: " + ex.Message); }
            };
            dialog.Controls.AddRange([intro, list, note, locationLabel, locationFrame, browse, logsLabel, logsFrame, browseLogs, all, defaults, apply, cancel]);
            ShowThemedDialog(dialog);
        }
        finally { _modal = false; }
    }
    private static Button Button(string text, int x, int y, int width) => new() { Text = text, Bounds = new Rectangle(x, y, width, 34) };
    private void EditRecord(int index)
    {
        if (_terminalSession is not null) { MessageBox.Show(this, "Disable Terminal before editing a recorded drive.", "Edit Recorded Drive"); return; }
        if (_busy) { MessageBox.Show(this, "Wait for the current drive read to finish before editing a record.", "Edit Recorded Drive"); return; }
        if (_modal || index < 0 || index >= _book.Records.Count) return;
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
            var remove = new Button { Text = "Remove Entry", Width = 140, Height = 34 };
            dialog.CancelButton = cancel;
            cancel.Click += (_, _) => dialog.Close();
            remove.Click += (_, _) =>
            {
                var saved = _book.Records[index];
                if (MessageBox.Show(dialog,
                    $"Remove the saved entry in workbook row {index + 2}?\n\nModel: {saved["Model"]}\nSerial: {saved["SerialNumber"]}\n\nThis removes the row from the workbook. Unsaved edits in this dialog will be discarded.",
                    "Confirm Remove Entry", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
                try
                {
                    _book.Remove(index);
                    RefreshGrid();
                    Activity($"Row {index + 2} removed: {saved["Model"]} / {saved["SerialNumber"]}", tone: StatusTone.Success);
                    dialog.Close();
                }
                catch (Exception ex) { Log(ex.ToString()); MessageBox.Show(dialog, "Entry was not removed: " + ex.Message, "Remove Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            };
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
                    Activity($"Row {index + 2} updated: {candidate["Model"]} / {candidate["SerialNumber"]}", tone: StatusTone.Success);
                    PlayDriveNotification(DriveNotification.Saved);
                    dialog.Close();
                }
                catch (Exception ex) { Log(ex.ToString()); MessageBox.Show(dialog, ex.Message, "Edit Recorded Drive", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            };
            actions.Controls.AddRange([cancel, save, remove]);
            dialog.Controls.Add(fields); dialog.Controls.Add(actions); dialog.Controls.Add(intro);
            ShowThemedDialog(dialog);
        }
        finally { _modal = false; }
    }
    private void ManualEntry(bool copyLast, Action<Form>? preview = null)
    {
        if (_busy) { MessageBox.Show(this, "Wait for the current drive read to finish before entering a manual record.", "Manual Drive Entry"); return; }
        _modal = true;
        try
        {
            using var dialog = new Form { Text = "Manual Drive Entry", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, ClientSize = new Size(560, 410), MaximizeBox = false, MinimizeBox = false, Font = new Font("Segoe UI", 10) };
            dialog.Controls.Add(new Label { Text = "Leave a field blank for N/A. Model and serial are saved in uppercase.", Bounds = new Rectangle(20, 12, 520, 30) });
            var labels = new[] { "1. Manufacturer", "2. Model", "3. Serial Number", "4. Capacity (number only)" };
            for (int i = 0; i < labels.Length; i++) dialog.Controls.Add(new Label { Text = labels[i], Bounds = new Rectangle(20, 47 + 43 * i, 195, 26) });
            var manufacturer = Box(219, 47, 320); var model = Box(219, 90, 320); var serial = Box(219, 133, 320); var amount = Box(219, 176, 175);
            model.CharacterCasing = CharacterCasing.Upper; serial.CharacterCasing = CharacterCasing.Upper;
            var unit = new ThemedComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Bounds = new Rectangle(405, 176, 134, 28), AccessibleName = "Capacity Unit" };
            _toolTip.SetToolTip(unit, "Capacity Unit");
            unit.Items.AddRange(["N/A", "B", "KB", "MB", "GB", "TB", "PB", "Other"]); unit.SelectedIndex = 0;
            var customUnitLabel = new Label { Text = "Custom Capacity Unit", Bounds = new Rectangle(20, 219, 195, 26) };
            var customUnit = Box(219, 219, 320);
            var typeLabel = new Label { Text = "5. Drive Type", Bounds = new Rectangle(20, 219, 195, 26) };
            var type = new ThemedComboBox { DropDownStyle = ComboBoxStyle.DropDown, AutoCompleteMode = AutoCompleteMode.None, Bounds = new Rectangle(219, 219, 320, 28), MaxDropDownItems = 15, IntegralHeight = true };
            var types = DriveTypes.Options;
            type.Items.AddRange(types.Cast<object>().ToArray());
            type.SelectedIndex = 0;
            void SizeTypeDropdown() => type.DropDownHeight = Math.Max(type.ItemHeight, type.ItemHeight * Math.Min(15, type.Items.Count));
            SizeTypeDropdown();
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
            var updatingTypeOptions = false;
            var latestTypeQuery = type.Text;
            var dropdownMaySelectText = false;
            type.SelectedIndexChanged += (_, _) =>
            {
                if (!updatingTypeOptions) { latestTypeQuery = type.Text; dropdownMaySelectText = false; }
                UpdateManualLayout();
            };
            void FilterTypeOptions(string query)
            {
                latestTypeQuery = query;
                dropdownMaySelectText = true;
                type.DroppedDown = false;
                updatingTypeOptions = true;
                try
                {
                    type.BeginUpdate(); type.Items.Clear(); type.Items.AddRange(DriveTypes.Matches(query).Cast<object>().ToArray()); type.EndUpdate();
                    SizeTypeDropdown();
                    type.DroppedDown = true;
                    // Reopening the native ComboBox can select the matching text. Clear that
                    // selection after opening so the next key extends the user's query.
                    type.Text = query; type.SelectionStart = query.Length; type.SelectionLength = 0;
                }
                finally { updatingTypeOptions = false; }
                type.BeginInvoke((Action)(() =>
                {
                    if (type.IsDisposed || !type.Focused || latestTypeQuery != query || type.SelectionLength == 0) return;
                    type.Text = query; type.SelectionStart = query.Length; type.SelectionLength = 0;
                }));
                UpdateManualLayout();
            }
            // The native ComboBox may select a matching item when its list reopens.
            // Handle typed characters against the user's query, rather than that
            // temporary selection (which could turn "SATA" into "ATA").
            type.KeyPress += (_, e) =>
            {
                if (char.IsControl(e.KeyChar)) return;
                var query = dropdownMaySelectText ? latestTypeQuery : type.Text;
                var start = dropdownMaySelectText ? query.Length : type.SelectionStart;
                var length = dropdownMaySelectText ? 0 : type.SelectionLength;
                e.Handled = true;
                FilterTypeOptions(query.Remove(start, length).Insert(start, e.KeyChar.ToString()));
                type.SelectionStart = start + 1; type.SelectionLength = 0;
            };
            type.KeyDown += (_, e) =>
            {
                if (e.KeyCode is not (Keys.Back or Keys.Delete)) return;
                var query = dropdownMaySelectText ? latestTypeQuery : type.Text;
                var start = dropdownMaySelectText ? query.Length : type.SelectionStart;
                var length = dropdownMaySelectText ? 0 : type.SelectionLength;
                if (length == 0 && e.KeyCode == Keys.Back && start > 0) { start--; length = 1; }
                if (length == 0 && e.KeyCode == Keys.Delete && start < query.Length) length = 1;
                e.SuppressKeyPress = true;
                if (length > 0) FilterTypeOptions(query.Remove(start, length));
                type.SelectionStart = start; type.SelectionLength = 0;
            };
            type.TextUpdate += (_, _) =>
            {
                if (!updatingTypeOptions) FilterTypeOptions(type.Text);
            };
            type.Enter += (_, _) => { dropdownMaySelectText = false; type.SelectAll(); };
            type.MouseDown += (_, _) => dropdownMaySelectText = false;
            dialog.Controls.AddRange([manufacturer, model, serial, amount, unit, customUnitLabel, customUnit, typeLabel, type, customTypeLabel, customType]);
            UpdateManualLayout();
            if (copyLast && _book.Records.LastOrDefault() is { } last) { Fill(last); serial.Clear(); }
            var review = Button("Next", 20, 355, 130); var copy = Button("Copy Last Drive", 162, 355, 160); var back = Button("Return to Scanning", 334, 355, 205);
            copy.Enabled = _book.Records.Count > 0;
            copy.Click += (_, _) => { if (_book.Records.LastOrDefault() is { } saved) { Fill(saved); serial.Clear(); serial.Focus(); } };
            back.Click += (_, _) => dialog.Close();
            review.Click += (_, _) =>
            {
                try
                {
                    var choice = type.Text.Trim();
                    if (!types.Contains(choice, StringComparer.OrdinalIgnoreCase)) throw new ArgumentException("Select a Drive Type from the list, or choose Other.");
                    var record = ManualValidation.Create(manufacturer.Text, model.Text, serial.Text, amount.Text, unit.Text == "N/A" ? "" : unit.Text, customUnit.Text, choice, customType.Text);
                    var duplicate = _book.HasSerial(record["SerialNumber"]);
                    if (duplicate) PlayDriveNotification(DriveNotification.Duplicate);
                    var action = Review(record, duplicate);
                    if (action == "Serial") { serial.Focus(); serial.SelectAll(); return; }
                    if (action == "Cancel") return;
                    var row = _book.Add(record); RefreshGrid(latestNumberAtTop: true); Activity($"Manual drive recorded as row {row}: {record["Model"]} / {record["SerialNumber"]}", tone: StatusTone.Success); PlayDriveNotification(DriveNotification.Saved);
                    MessageBox.Show(dialog, $"Drive saved as row {row}.", "Drive Recorded", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    copy.Enabled = true;
                    serial.Clear();
                    if (action == "Save") { manufacturer.Clear(); model.Clear(); amount.Clear(); unit.SelectedIndex = 0; customUnit.Clear(); type.Items.Clear(); type.Items.AddRange(types.Cast<object>().ToArray()); type.SelectedIndex = 0; SizeTypeDropdown(); customType.Clear(); dialog.Text = "Manual Drive Entry"; }
                    else dialog.Text = "Manual Drive Entry – Copy Saved Drive";
                    serial.Focus();
                }
                catch (Exception ex) { Log(ex.ToString()); MessageBox.Show(dialog, ex.Message, "Manual Drive Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            };
            dialog.Controls.AddRange([review, copy, back]);
            if (preview is null) ShowThemedDialog(dialog);
            else
            {
                CollectorTheme.Apply(dialog, _dark);
                dialog.Show(this); Application.DoEvents();
                preview(dialog);
                dialog.Close();
            }
            void Fill(DriveRecord record)
            {
                manufacturer.Text = record["Manufacturer"] == "N/A" ? "" : record["Manufacturer"];
                model.Text = record["Model"] == "N/A" ? "" : record["Model"];
                amount.Clear(); unit.SelectedIndex = 0; customUnit.Clear(); customType.Clear();
                var match = Regex.Match(record["Capacity"], @"^([0-9]+(?:\.[0-9]+)?)\s+(.+)$");
                if (match.Success) { amount.Text = match.Groups[1].Value; var value = match.Groups[2].Value; if (unit.Items.Contains(value)) unit.SelectedItem = value; else { unit.SelectedItem = "Other"; customUnit.Text = value; } }
                var media = record["Type"];
                type.Items.Clear(); type.Items.AddRange(types.Cast<object>().ToArray()); SizeTypeDropdown();
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
        var summary = new RichTextBox { Bounds = new Rectangle(20, 20, 450, 158), BorderStyle = BorderStyle.Fixed3D, BackColor = SystemColors.Window, ReadOnly = true, TabStop = false, DetectUrls = false, ScrollBars = RichTextBoxScrollBars.Vertical, Font = new Font("Consolas", 11), Lines = [$"Manufacturer: {record["Manufacturer"]}", $"Model:        {record["Model"]}", $"Serial:       {record["SerialNumber"]}", $"Capacity:     {record["Capacity"]}", $"Type:         {record["Type"]}"] };
        var note = new Label { Text = duplicate ? "This serial is already in the workbook. Change it or cancel this record." : record["SerialNumber"] == "N/A" ? "Serial N/A cannot be checked for duplicates." : "Review these values before saving a new row.", Bounds = new Rectangle(20, 188, 450, 48) };
        dialog.Controls.AddRange([summary, note]); string action = "Cancel";
        dialog.Shown += (_, _) => summary.Select(0, 0);
        Button Choice(string name, string text, int x, int width) { var button = Button(text, x, 272, width); button.UseMnemonic = false; button.Click += (_, _) => { action = name; dialog.Close(); }; dialog.Controls.Add(button); return button; }
        if (duplicate) { Choice("Serial", "Change Serial", 118, 120); Choice("Cancel", "Cancel", 252, 120); }
        else
        {
            Choice("Save", "Save", 20, 142);
            var saveCopy = Choice("Copy", "Save & Copy", 174, 142);
            _toolTip.SetToolTip(saveCopy, "Save this drive and copy its details with a new serial number.");
            dialog.CancelButton = Choice("Cancel", "Cancel", 328, 142);
        }
        ShowThemedDialog(dialog); return action;
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
        var normalized = Create(edited["Manufacturer"] == "N/A" ? "" : edited["Manufacturer"], edited["Model"] == "N/A" ? "" : edited["Model"],
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
    public static DriveRecord Create(string manufacturer, string model, string serial, string amount, string unit, string customUnit, string type, string customType)
    {
        manufacturer = manufacturer.Trim(); model = model.Trim(); serial = serial.Trim(); amount = amount.Trim(); unit = unit.Trim(); customUnit = customUnit.Trim(); type = type.Trim(); customType = customType.Trim();
        Check(manufacturer, @"\A[A-Za-z0-9][A-Za-z0-9 .&()+'/_-]{0,79}\z", 80, "Manufacturer");
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
        return new DriveRecord { ["Manufacturer"] = manufacturer, ["Model"] = model.ToUpperInvariant(), ["SerialNumber"] = serial.ToUpperInvariant(), ["Capacity"] = capacity, ["Type"] = type };
    }
    private static void Check(string value, string pattern, int max, string label) { if (value.Length > 0 && (value.Length > max || !Regex.IsMatch(value, pattern))) throw new ArgumentException($"{label}: use only plain letters, digits, spaces, or standard punctuation (up to {max} characters)."); }
}

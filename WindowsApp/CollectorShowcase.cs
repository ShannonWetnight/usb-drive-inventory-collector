#if SHOWCASE_CAPTURE
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace USBDriveInventoryCollector;

internal sealed partial class CollectorForm
{
    // Only the explicit capture command sets this callback. Normal dialogs remain modal.
    private Action<Form>? _showcasePreview;

    internal static void CaptureShowcase(string directory)
    {
        Directory.CreateDirectory(directory);
        var savedPreference = CollectorSettings.Theme();
        using var form = new CollectorForm(initialize: false);
        // In-memory demonstration data: no drive probing, workbook writes, or saved preferences.
        form._book = new InventoryBook(@"C:\USB-Drive-Inventory-Collector\Output\Inventory.xlsx");
        form._logPath = @"C:\USB-Drive-Inventory-Collector\Output\Logs\USB-Drive-Inventory-Collector.log";
        var sample = new DriveRecord { ["Manufacturer"] = "Example", ["Model"] = "DEMO-SATA-SSD", ["SerialNumber"] = "DEMO-0001", ["Capacity"] = "500 GB", ["Type"] = "2.5-inch SATA SSD" };
        form._book.Records.Add(sample);
        form._book.Records.Add(new DriveRecord { ["Manufacturer"] = "Example", ["Model"] = "DEMO-NVME", ["SerialNumber"] = "DEMO-0002", ["Capacity"] = "1 TB", ["Type"] = "M.2 NVMe SSD" });
        form._book.Records.Add(new DriveRecord { ["Manufacturer"] = "Example", ["Model"] = "DEMO-USB-FLASH", ["SerialNumber"] = "DEMO-0003", ["Capacity"] = "64 GB", ["Type"] = "USB Flash Drive" });
        form.RefreshGrid();
        form.Show();
        Application.DoEvents();
        var captures = new List<object>();
        foreach (var preference in new[] { ThemePreference.Light, ThemePreference.Dark })
        {
            var suffix = preference == ThemePreference.Light ? "light_mode" : "dark_mode";
            form._themePreference = preference;
            form._tabs.SelectedIndex = 0;
            form._paused = false;
            form._pause.Text = "Pause Scanning";
            form._statusTone = StatusTone.Default;
            form._status.Text = "Waiting for a USB drive...";
            form._guidance.Text = "Insert one drive at a time. Use Manual Drive Entry or Workbook Setup at any time.";
            form.ApplyTheme();
            form._viewChanged = false;
            form.UpdateResetViewButton();
            Save(form, "recorded_drives");

            form._activity.Items.Clear();
            form._activity.Items.AddRange([
                "10:02:14  INFO   Manual drive recorded as row 4: DEMO-USB-FLASH / DEMO-0003",
                "10:01:42  INFO   Manual drive recorded as row 3: DEMO-NVME / DEMO-0002",
                "10:01:05  INFO   Manual drive recorded as row 2: DEMO-SATA-SSD / DEMO-0001",
                "10:00:00  INFO   Waiting for a USB drive..."
            ]);
            form._tabs.SelectedIndex = 1;
            Save(form, "activity");
            form._tabs.SelectedTab = form._terminalPage;
            Save(form, "terminal");
            form._tabs.SelectedIndex = 0;

            Preview("manual_entry", () => form.ManualEntry(true));
            Preview("review_entry", () => form.Review(sample, duplicate: false));
            Preview("duplicate_serial", () => form.Review(sample, duplicate: true));
            form._showcasePreview = dialog =>
            {
                Save(dialog, "edit_entry");
                CaptureRemovalConfirmation(dialog, () =>
                    ControlsIn(dialog).OfType<Button>().Single(b => b.Text == "Remove Entry").PerformClick(),
                    (handle, title) => SaveHandle(handle, title, "remove_entry"));
            };
            form.EditRecord(0);
            form._showcasePreview = null;
            Preview("workbook_setup", form.Setup);
            Preview("version_info", form.ShowVersionInformation);

            form._paused = true;
            form._pause.Text = "Resume Scanning";
            form._statusTone = StatusTone.Paused;
            form._status.Text = "Scanning is paused.";
            form.FadeStatusBackground(StatusTone.Paused);
            form._fadeStart = System.Diagnostics.Stopwatch.GetTimestamp() - System.Diagnostics.Stopwatch.Frequency;
            form.AdvanceStatusFade();
            form.ColorizeStatus();
            Save(form, "paused_scanning");

            void Preview(string name, Action open)
            {
                form._showcasePreview = dialog =>
                {
                    if (name == "workbook_setup")
                    {
                        var logs = ControlsIn(dialog).OfType<TextBox>().Single(input => input.Text == CollectorSettings.LogsDirectory());
                        logs.Text = Path.GetDirectoryName(form._logPath)!;
                    }
                    Save(dialog, name);
                };
                try { open(); }
                finally { form._showcasePreview = null; }
            }
            void Save(Form window, string name)
            {
                window.Refresh();
                Application.DoEvents();
                SaveHandle(window.Handle, window.Text, name);
            }
            void SaveHandle(IntPtr handle, string title, string name)
            {
                if (!GetWindowRect(handle, out var rect)) throw new InvalidOperationException("Cannot read window bounds.");
                var width = rect.Right - rect.Left;
                var height = rect.Bottom - rect.Top;
                using var bitmap = new Bitmap(width, height);
                using (var graphics = Graphics.FromImage(bitmap))
                {
                    var dc = graphics.GetHdc();
                    try
                    {
                        if (!PrintWindow(handle, dc, 2)) throw new InvalidOperationException($"Cannot capture {title}.");
                    }
                    finally { graphics.ReleaseHdc(dc); }
                }
                var samples = 0;
                var black = 0;
                for (var x = 20; x < width - 20; x += 20)
                    for (var y = 45; y < height - 20; y += 20)
                    {
                        samples++;
                        if (bitmap.GetPixel(x, y).ToArgb() == Color.Black.ToArgb()) black++;
                    }
                if (samples == 0 || black > samples * 0.9)
                    throw new InvalidDataException($"Windows returned a blank capture for {title} ({preference}).");
                // Win32 bounds include invisible resize margins. DWM reports
                // the visible frame in the same physical pixels for this DPI-aware app.
                if (DwmGetWindowAttribute(handle, 9, out var visible, Marshal.SizeOf<NativeRect>()) != 0)
                    throw new InvalidOperationException("Cannot read visible window bounds.");
                var crop = new Rectangle(visible.Left - rect.Left, visible.Top - rect.Top,
                    visible.Right - visible.Left, visible.Bottom - visible.Top);
                if (crop.Width <= 0 || crop.Height <= 0 || !new Rectangle(0, 0, width, height).Contains(crop))
                    throw new InvalidDataException("Visible window bounds do not fit the native capture.");
                using var framed = bitmap.Clone(crop, PixelFormat.Format32bppArgb);
                var file = $"showcase_{name}_{suffix}.png";
                framed.Save(Path.Combine(directory, file), ImageFormat.Png);
                captures.Add(new { file, theme = preference.ToString(), title, width = crop.Width, height = crop.Height,
                    windowWidth = width, windowHeight = height, cropLeft = crop.Left, cropTop = crop.Top });
            }
        }
        if (CollectorSettings.Theme() != savedPreference || form._book.Records.Count != 3)
            throw new InvalidDataException("Capture changed preferences or sample records.");
        File.WriteAllText(Path.Combine(directory, "capture.json"), JsonSerializer.Serialize(captures, new JsonSerializerOptions { WriteIndented = true }));
    }

    private static IEnumerable<Control> ControlsIn(Control parent)
    {
        foreach (Control child in parent.Controls)
        {
            yield return child;
            foreach (var descendant in ControlsIn(child)) yield return descendant;
        }
    }

    private static void CaptureRemovalConfirmation(Form owner, Action open, Action<IntPtr, string> capture)
    {
        const string title = "Confirm Remove Entry";
        Exception? failure = null;
        var captured = false;
        var ticks = 0;
        using var timer = new System.Windows.Forms.Timer { Interval = 100 };
        timer.Tick += (_, _) =>
        {
            var handle = FindWindow(null, title);
            if (handle == IntPtr.Zero || GetWindowThreadProcessId(handle, out var process) == 0 || process != Environment.ProcessId)
            {
                if (++ticks < 100) return;
                failure = new TimeoutException("Removal confirmation did not open.");
                timer.Stop();
                owner.Close();
                return;
            }
            timer.Stop();
            try { capture(handle, title); captured = true; }
            catch (Exception ex) { failure = ex; }
            finally { PostMessage(handle, 0x0111, new IntPtr(7), IntPtr.Zero); } // IDNO: do not remove a row.
        };
        timer.Start();
        open();
        if (failure is not null) throw failure;
        if (!captured) throw new InvalidOperationException("Removal confirmation was not captured.");
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect { public int Left, Top, Right, Bottom; }
    [DllImport("dwmapi.dll")] private static extern int DwmGetWindowAttribute(IntPtr window, uint attribute, out NativeRect rect, int size);
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out NativeRect rect);
    [DllImport("user32.dll")] private static extern bool PrintWindow(IntPtr window, IntPtr dc, uint flags);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr FindWindow(string? className, string title);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
}

#endif

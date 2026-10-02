using Microsoft.Win32;
using System.Runtime.InteropServices;

namespace USBDriveInventoryCollector;

internal enum ThemePreference { System, Light, Dark }

internal static class CollectorTheme
{
    // Branded header controls retain their existing colors in either theme.
    public const string PreserveColors = "PreserveThemeColors";
    public static Color Surface(bool dark) => dark ? Color.FromArgb(28, 32, 38) : Color.FromArgb(246, 248, 251);
    public static Color Field(bool dark) => dark ? Color.FromArgb(38, 43, 51) : SystemColors.Window;
    public static Color Text(bool dark) => dark ? Color.FromArgb(230, 234, 240) : SystemColors.ControlText;
    public static bool IsDark(ThemePreference preference)
    {
        if (SystemInformation.HighContrast) return false;
        if (preference != ThemePreference.System) return preference == ThemePreference.Dark;
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
            return key?.GetValue("AppsUseLightTheme") is int value && value == 0;
        }
        catch { return false; }
    }
    public static void Apply(Control control, bool dark)
    {
        if (Equals(control.Tag, PreserveColors) && !SystemInformation.HighContrast) return;
        var highContrast = SystemInformation.HighContrast;
        var surface = highContrast ? SystemColors.Control : Surface(dark);
        var field = highContrast ? SystemColors.Window : Field(dark);
        var text = highContrast ? SystemColors.ControlText : Text(dark);
        control.BackColor = surface; control.ForeColor = text;
        switch (control)
        {
            case TextBoxBase or ListBox or ComboBox:
                control.BackColor = field;
                break;
            case Button button:
                button.FlatStyle = highContrast ? FlatStyle.Standard : FlatStyle.Flat;
                button.UseVisualStyleBackColor = highContrast;
                button.BackColor = highContrast ? SystemColors.Control : dark ? Color.FromArgb(48, 55, 65) : SystemColors.Control;
                button.FlatAppearance.BorderColor = dark ? Color.FromArgb(89, 101, 115) : SystemColors.ControlDark;
                break;
            case LinkLabel link:
                link.LinkColor = highContrast ? SystemColors.HotTrack : dark ? Color.FromArgb(135, 195, 255) : Color.RoyalBlue;
                link.ActiveLinkColor = link.LinkColor;
                link.VisitedLinkColor = link.LinkColor;
                break;
            case Panel panel when panel.BorderStyle != BorderStyle.None:
                panel.BackColor = field;
                break;
            case TabPage page:
                page.UseVisualStyleBackColor = false;
                break;
            case DataGridView grid:
                grid.BackgroundColor = field;
                grid.GridColor = dark ? Color.FromArgb(70, 79, 91) : SystemColors.ControlDark;
                grid.DefaultCellStyle.BackColor = field;
                grid.DefaultCellStyle.ForeColor = text;
                grid.DefaultCellStyle.SelectionBackColor = highContrast ? SystemColors.Highlight : Color.FromArgb(48, 76, 102);
                grid.DefaultCellStyle.SelectionForeColor = highContrast ? SystemColors.HighlightText : Color.White;
                grid.AlternatingRowsDefaultCellStyle.BackColor = highContrast ? field : dark ? Color.FromArgb(32, 37, 44) : Color.FromArgb(240, 244, 249);
                break;
        }
        foreach (Control child in control.Controls) Apply(child, dark);
        control.Invalidate();
    }
    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);
    public static void ApplyTitleBar(Form form, bool dark)
    {
        if (!form.IsHandleCreated) return;
        var enabled = dark && !SystemInformation.HighContrast ? 1 : 0;
        // Unsupported Windows versions simply leave the native title bar unchanged.
        _ = DwmSetWindowAttribute(form.Handle, 20, ref enabled, sizeof(int));
    }
}

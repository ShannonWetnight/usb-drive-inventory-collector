using System.Runtime.InteropServices;

namespace USBDriveInventoryCollector;

// Retains native selection, keyboard navigation, and editable text behavior.
// Draw the list and arrow explicitly so Windows' light combo rendering cannot
// hide a dark-mode field's text or dropdown affordance.
internal sealed class ThemedComboBox : ComboBox
{
    public ThemedComboBox()
    {
        FlatStyle = FlatStyle.Flat;
        DrawMode = DrawMode.OwnerDrawFixed;
    }

    protected override void OnDrawItem(DrawItemEventArgs e)
    {
        var selected = (e.State & DrawItemState.Selected) != 0;
        var background = selected ? SystemColors.Highlight : BackColor;
        var foreground = selected ? SystemColors.HighlightText : ForeColor;
        using var brush = new SolidBrush(background);
        e.Graphics.FillRectangle(brush, e.Bounds);
        var text = e.Index >= 0 && e.Index < Items.Count ? GetItemText(Items[e.Index]) : Text;
        TextRenderer.DrawText(e.Graphics, text, Font, e.Bounds, foreground,
            TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
        e.DrawFocusRectangle();
        base.OnDrawItem(e);
    }

    protected override void WndProc(ref Message m)
    {
        base.WndProc(ref m);
        if (SystemInformation.HighContrast || m.Msg is not (0x000F or 0x0317 or 0x0318)) return;
        using var graphics = m.Msg is 0x0317 or 0x0318 && m.WParam != IntPtr.Zero
            ? Graphics.FromHdc(m.WParam) : Graphics.FromHwnd(Handle);
        DrawArrow(graphics);
    }

    private void DrawArrow(Graphics graphics)
    {
        // Get the native button bounds instead of assuming a particular DPI.
        var info = new ComboBoxInfo { Size = Marshal.SizeOf<ComboBoxInfo>() };
        if (!GetComboBoxInfo(Handle, ref info)) return;
        var bounds = Rectangle.FromLTRB(info.Button.Left, info.Button.Top, info.Button.Right, info.Button.Bottom);
        using var background = new SolidBrush(BackColor);
        graphics.FillRectangle(background, bounds);
        var color = Enabled ? ForeColor : SystemColors.GrayText;
        using var brush = new SolidBrush(color);
        var x = bounds.Left + bounds.Width / 2;
        var y = bounds.Top + bounds.Height / 2;
        var halfWidth = Math.Max(3, DeviceDpi / 32);
        graphics.FillPolygon(brush, [new Point(x - halfWidth, y - 1), new Point(x + halfWidth, y - 1), new Point(x, y + halfWidth - 1)]);
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)]
    private struct ComboBoxInfo
    {
        public int Size;
        public NativeRect Item, Button;
        public int ButtonState;
        public IntPtr Combo, Edit, List;
    }
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetComboBoxInfo(IntPtr handle, ref ComboBoxInfo info);
}

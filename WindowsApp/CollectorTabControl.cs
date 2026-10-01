using System.Runtime.InteropServices;

namespace USBDriveInventoryCollector;

internal sealed class CollectorTabControl : TabControl
{
    public Control? HeaderActions { get; set; }

    protected override CreateParams CreateParams
    {
        get
        {
            var parameters = base.CreateParams;
            parameters.Style |= 0x04000000; // WS_CLIPSIBLINGS: respect the opaque action strip.
            return parameters;
        }
    }

    protected override void WndProc(ref Message m)
    {
        // WM_PRINT/WM_PRINTCLIENT use a shared DC and do not apply the normal
        // sibling-window clipping. Preserve the action strip during captures too.
        if (m.Msg is 0x0317 or 0x0318 && m.WParam != IntPtr.Zero && HeaderActions is { Visible: true } actions)
        {
            var bounds = RectangleToClient(actions.RectangleToScreen(actions.ClientRectangle));
            var saved = SaveDC(m.WParam);
            try
            {
                ExcludeClipRect(m.WParam, bounds.Left, bounds.Top, bounds.Right, bounds.Bottom);
                base.WndProc(ref m);
            }
            finally { if (saved != 0) RestoreDC(m.WParam, saved); }
            return;
        }
        base.WndProc(ref m);
    }

    [DllImport("gdi32.dll")]
    private static extern int SaveDC(IntPtr dc);
    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool RestoreDC(IntPtr dc, int saved);
    [DllImport("gdi32.dll")]
    private static extern int ExcludeClipRect(IntPtr dc, int left, int top, int right, int bottom);
}

# Function Showcase Captures

These PNGs come from the app's actual Windows controls and dialogs at their default opening sizes. All drive records, activity entries, and displayed paths are demonstration data. Capturing skips drive scanning and does not save records, change theme preferences, or remove an entry. The removal dialog is captured and answered No.

Each function has a Light Theme and Dark Theme capture, with `_light_mode` or `_dark_mode` at the end of its PNG filename. `capture.json` records the theme, window title, visible pixel dimensions, original window dimensions, and crop offsets. Captures exclude Windows' invisible resize margins; the app window and its controls are not resized. The Terminal image shows the disabled tab and its Enable Terminal control.

The capture helper is compiled only when `ShowcaseCapture=true`. Normal release builds exclude the capture command, helper, and dialog-capture hook. The Windows workflow publishes a separate capture build and uploads a `Function-Showcase-v4.3.5` artifact with fresh PNGs. To capture locally, build and run that helper on Windows from the repository root:

```powershell
dotnet publish WindowsApp/USBDriveInventoryCollector.csproj -c Release -r win-x64 --self-contained true -p:ShowcaseCapture=true -o showcase-publish
.\showcase-publish\USB-Drive-Inventory-Collector-v4.3.5.exe --capture-showcase C:\Temp\Collector-Showcase
```

Replace the PNGs and `capture.json` in this directory with the new capture files. Keep the original visible-window pixel dimensions. From the repository root, rebuild the GIF with Python and Pillow:

```shell
python -m pip install Pillow==12.3.0
python .github/scripts/build-showcase.py
```

The GIF holds each frame for 4.5 seconds and loops, approximately 1.33 times the previous playback speed. Each function appears in Light Theme, then Dark Theme, before the animation moves to the next function. Smaller windows are centered on a fixed canvas without resizing them. GIF palette conversion can reduce color detail; the PNGs preserve the original captures.

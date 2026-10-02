# Function showcase captures

These PNGs come from the app's actual Windows controls and dialogs at their default opening sizes. All drive records, activity entries, and displayed paths are demonstration data. Capturing skips drive scanning and does not save records, change theme preferences, or remove an entry. The removal dialog is captured and answered No.

Each function has two PNGs: `showcase_function.png` uses the dark theme; `showcase_function2.png` uses the light theme. `capture.json` records the theme, window title, and pixel dimensions of each capture. The Terminal image shows the disabled tab and its Enable Terminal control.

The Windows build workflow uploads a `Function-Showcase-v4.3.5` artifact with fresh PNGs. To capture locally, run the built executable on Windows:

```powershell
.\USB-Drive-Inventory-Collector-v4.3.5.exe --capture-showcase C:\Temp\Collector-Showcase
```

Copy the PNGs and `capture.json` from the capture folder into this directory. Keep the original PNG dimensions. From the repository root, rebuild the GIF with Python and Pillow:

```shell
python -m pip install Pillow==12.3.0
python .github/scripts/build-showcase.py
```

The GIF holds each frame for six seconds and loops. It centers smaller windows on a fixed canvas without resizing them. GIF palette conversion can reduce color detail; the PNGs preserve the original captures.

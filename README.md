# USB Drive Inventory Collector

USB Drive Inventory Collector records the identity of USB-connected physical drives in an XLSX workbook. It excludes Windows boot and system disks, saves after each drive, and supports manual records when a drive cannot be read. It does not format, erase, or write to the attached drive.

> **AI workflow notice:** The project was written through AI prompting and reviewed by its maintainer. Validate collected data in your own environment before relying on it.

## Choose an interface

- **Native Windows GUI (4.0 preview):** The Windows Forms application in [WindowsApp](WindowsApp/README.md) builds into a double-clickable `USB-Drive-Inventory-Collector.exe`. It does not launch PowerShell and does not require changing PowerShell execution policy. This is an unsigned, unreleased preview; Windows and Microsoft Defender testing is still in progress.
- **Legacy console collector:** [USB-Drive-Inventory-Collector.ps1](USB-Drive-Inventory-Collector.ps1) remains available for an elevated PowerShell session where script execution is allowed. It is independent of the GUI.

The two collectors write the same five default workbook columns: Make, Model, Serial Number, Reported Capacity, and Type. Both can add optional identity columns and write directly to XLSX without Excel.

## Native GUI preview

Download the artifact from a successful **Build native Windows GUI** run in [pull request #6](https://github.com/ShannonWetnight/usb-drive-inventory-collector/pull/6), extract it, and double-click `USB-Drive-Inventory-Collector.exe`. Approve Windows' administrator prompt. The EXE is self-contained for Windows x64; automatic drive reads also require smartmontools' `smartctl.exe`. If smartmontools is missing, the application asks before installing it through WinGet.

Insert one drive at a time. **Recorded drives** shows saved records, **Activity** shows progress and errors, and **Technical details** shows the workbook and log paths. **Manual entry** records a drive by hand; **Copy last** starts a new record with the prior drive's information and asks for a new serial. **Workbook setup** selects optional identity columns and backs up the workbook before changing its layout. **Pause scanning** and **Finish** control the session. The app can offer to disable AutoPlay temporarily for the signed-in desktop user and restore its previous setting on normal exit.

The GUI saves `Output/Inventory.xlsx` and a timestamped log under `Output/Logs/` beside the EXE. Keep the workbook closed while collecting so the app can replace it on each save. This build has not yet been tested on the maintainer's USB adapters or cleared against the reported Defender detection. See [WindowsApp/README.md](WindowsApp/README.md) for the preview build command and validation checklist.

## Legacy console

On Windows, open an elevated PowerShell session in the repository directory and run:

```powershell
.\USB-Drive-Inventory-Collector.ps1
```

The script needs Windows PowerShell 5.1 or later, administrator access, and smartmontools. It asks before using WinGet if smartmontools is missing. If local policy blocks scripts, use your organization's approved process; the native GUI preview is the double-click option.

While waiting for a drive, press `[M]` for manual entry, `[S]` for workbook setup, or `[D]` for technical details. Insert one USB drive at a time, wait for the saved row, remove it, and insert the next. Press `[Ctrl+C]` when finished. After a read error, remove and reinsert the drive to retry or enter it manually.

Manual entry asks for Make, Model, Serial Number, capacity number and unit, and a categorized drive type. Skipped fields become `N/A`; model and serial are capitalized; leading and trailing spaces are removed. Review and edit before saving. Choose to save and continue or save and copy with another serial. Duplicate serials must be changed or canceled. The console also supports optional identity columns, a temporary AutoPlay change, and a timestamped diagnostic log.

The default output is `Output/Inventory.xlsx` and `Output/Logs/USB-Drive-Inventory-Collector-YYYYMMDD-HHMMSS.log` beside the script. To select a different workbook or open manual entry immediately:

```powershell
.\USB-Drive-Inventory-Collector.ps1 -OutputPath "C:\Inventory\Inventory.xlsx" -ManualEntryOnStartup
```

Run `Get-Help .\USB-Drive-Inventory-Collector.ps1 -Full` or inspect the parameter block for other console options. The native GUI currently uses the default output folder beside its EXE.

## Drive identification and limits

The collectors query USB physical disks with smartctl autodetection and transport fallbacks for common USB bridges. They classify drives as specifically as their reported identity allows, including PATA/IDE where supported. Unknown values are recorded as `N/A`. Some bridges hide the drive's identity or report their own; check unusual records against the physical label.

A smartctl process is limited to 30 seconds. If the bridge locks up, unplug and reconnect the adapter before retrying. The application does not reset a shared USB controller. A drive that is swapped entirely between scan intervals may be missed; wait for its removal message.

## License

[The Unlicense](UNLICENSE). The software is provided without warranty.

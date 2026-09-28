# Native Windows GUI preview

This folder contains the Windows Forms collector being tested in [pull request #6](https://github.com/ShannonWetnight/usb-drive-inventory-collector/pull/6). It is separate from the PowerShell console collector in the repository root. The maintainer has confirmed that the preview EXE launches on Windows; drive and workbook workflows are being checked before release.

The Windows build job publishes a self-contained `USB-Drive-Inventory-Collector.exe` in its Actions artifact. Download the artifact from the latest successful run on the PR, extract it, and double-click the EXE. Windows asks for administrator access. The app still needs `smartctl.exe` from smartmontools; if it is missing, the app asks before running WinGet. No PowerShell execution-policy change is needed for the EXE.

The GUI reads USB physical disks through Windows Storage, excludes boot and system disks, and probes smartctl's USB transports for drive identity. It saves to `Output\Inventory.xlsx` beside the EXE, using the same headers and XLSX layout as the console collector. The app also has manual entry, copy-last, duplicate serial checks, optional workbook columns, a timestamped log, a pause control, and an optional temporary AutoPlay change. It restores the AutoPlay value when the GUI closes normally.

To build from source on Windows with the .NET 8 SDK:

```powershell
dotnet publish WindowsApp/USBDriveInventoryCollector.csproj --configuration Release --runtime win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true --output publish
```

The Actions artifact is an unsigned preview build. Before release, check a known USB drive, appending to an existing workbook, duplicate serial handling, and AutoPlay restoration on exit. The PowerShell console collector remains available for users who have an approved script execution policy.

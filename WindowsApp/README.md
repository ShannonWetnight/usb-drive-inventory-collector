# Native Windows GUI

This folder contains the Windows Forms GUI source for version 4.0.0. The PowerShell console collector remains in the repository root.

Download the signed `USB-Drive-Inventory-Collector.exe` from the [v4.0.0 release](https://github.com/ShannonWetnight/usb-drive-inventory-collector/releases/tag/v4.0.0) and double-click it. Windows asks for administrator access. The app still needs `smartctl.exe` from smartmontools; if it is missing, the app asks before running WinGet. No PowerShell execution-policy change is needed for the EXE.

The GUI reads USB physical disks through Windows Storage, excludes boot and system disks, and probes smartctl's USB transports for drive identity. It saves to `Output\Inventory.xlsx` beside the EXE, using the same headers and XLSX layout as the console collector. The app also has manual entry, copy-last, duplicate serial checks, optional workbook columns, a timestamped log, a pause control, and an optional temporary AutoPlay change. It restores the AutoPlay value when the GUI closes normally.

To build from source on Windows with the .NET 8 SDK:

```powershell
dotnet publish WindowsApp/USBDriveInventoryCollector.csproj --configuration Release --runtime win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true --output publish
```

CI build artifacts are unsigned test builds. The release EXE is signed with the publisher name Shannon Wetnight. The PowerShell console collector remains available for users who have an approved script execution policy.

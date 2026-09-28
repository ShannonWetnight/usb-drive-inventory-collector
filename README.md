# USB Drive Inventory Collector

USB Drive Inventory Collector records USB-connected physical drives in an XLSX workbook. It skips Windows boot and system disks and saves each drive as soon as it is read. You can enter a record by hand when an adapter does not report useful drive information. The collector does not erase, format, or write to an attached drive.

> **AI workflow notice:** This project was written through AI prompting and reviewed by its maintainer. Check collected data against the drive label when accuracy matters.

## Download and run

The [signed v4.0.0 release](https://github.com/ShannonWetnight/usb-drive-inventory-collector/releases/tag/v4.0.0) contains a self-contained Windows x64 EXE. Download `USB-Drive-Inventory-Collector.exe`, double-click it, and approve the administrator prompt. Microsoft Excel and a PowerShell execution policy change are not required.

The current source is the v4.1.0 development build. It adds the Terminal interface and the workflow changes described below. Those changes are not in the v4.0.0 download until a later signed release is published.

Automatic drive reads require smartmontools (`smartctl.exe`). If it is missing, the app asks before installing it through WinGet. Insert one drive at a time and wait for its saved row before swapping drives. The app can offer to disable AutoPlay for the session and restore the previous setting on normal exit.

## Interfaces

The GUI shows **Recorded Drives** and **Activity**. Double-click a recorded row to edit it. **Manual Drive Entry** lets you enter a drive, review it, and save it; **Copy Last Drive** is inside that window for batches that share a model and capacity. Drive types are searchable by name, with **Other** at the end for a custom type.

Use **Workbook Setup** to select optional identity columns and choose a **Workbook Save Location**. Selecting a new path copies the current workbook there. Selecting an existing workbook switches to its records after confirmation. Column changes make a backup of the active workbook. **Pause Scanning** stops new reads, **Finish** ends the session, and the info button in the banner opens **Version Information**.

The **Terminal** button opens console mode from the same EXE. The GUI pauses scanning while Terminal is open, then reloads the workbook when it closes. Terminal also opens directly with `USB-Drive-Inventory-Collector.exe --terminal`. Its keys are `[M]` Manual Drive Entry, `[L]` Copy Last Drive, `[S]` Workbook Setup, `[P]` Pause Scanning, `[D]` Version Information, `[H]` Help, and `[Q]` Finish. Manual entries use the same validation and duplicate checks as the GUI.

The default workbook is `Output/Inventory.xlsx` beside the EXE. Session logs go in `Output/Logs/`. The chosen workbook location is remembered for the Windows account running the app. Keep the workbook closed in spreadsheet software while collecting so each save can replace the file.

## Drive identification

The collector uses smartctl autodetection and USB transport fallbacks for common bridges. It stores Make, Model, Serial Number, Reported Capacity, and Type by default. Workbook Setup can add fields such as firmware, interface, sector size, and the probe transport. Model and serial values entered by hand are saved in uppercase; blank fields become `N/A`. Duplicate serial numbers cannot be saved twice.

Some USB bridges report their own identity or hide the drive's details. A smartctl process times out after 30 seconds. If the adapter stops responding, unplug and reconnect it before retrying. The app does not reset a shared USB controller.

## Build from source

Use the .NET 8 SDK on Windows:

```powershell
dotnet publish WindowsApp/USBDriveInventoryCollector.csproj --configuration Release --runtime win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true --output publish
```

The [Build Windows GUI](.github/workflows/build-native.yml) workflow checks pull requests and produces an unsigned preview EXE. Official releases use [Publish Signed Release](.github/workflows/release.yml), which checks the version and release notes, signs the EXE through Azure Artifact Signing, verifies the publisher, and attaches the EXE and SHA-256 checksum to a GitHub release.

## License

[The Unlicense](UNLICENSE). The software is provided without warranty.

# USB Drive Inventory Collector

> **AI Workflow Notice:** This project was written through AI prompting and reviewed by its maintainer. Check collected data against the drive label when accuracy matters.

## Table of Contents

- [Features](#features)
- [Download and Run](#download-and-run)
- [Interfaces](#interfaces)
  - [Graphical Interface](#graphical-interface)
  - [Terminal Interface](#terminal-interface)
  - [Terminal Keys](#terminal-keys)
  - [Settings and Values](#settings-and-values)
- [Drive Identification](#drive-identification)
- [Files and Workbook](#files-and-workbook)
- [Build from Source](#build-from-source)
- [License](#license)

## Features

- Scans USB-connected physical drives and excludes Windows boot and system disks.
- Reads drive identity through smartmontools, with USB transport fallbacks for common adapters.
- Saves each drive to an XLSX workbook as soon as it is recorded. Excel is not required.
- Provides manual entry with input checks, review, duplicate-serial detection, and `N/A` for skipped values.
- Lets you copy the last saved drive when entering a batch that shares the same information.
- Supports optional identity columns and a configurable workbook location.
- Includes both graphical and terminal interfaces in the same executable.
- Offers to disable AutoPlay for the session and restores its previous setting when the collector exits normally.
- Shows recorded drives and activity, and lets you edit a saved row by double-clicking it.

## Download and Run

The next tagged version is **4.2.2**. Signed Windows x64 builds are published on the [GitHub Releases page](https://github.com/ShannonWetnight/usb-drive-inventory-collector/releases). The build produces `USB-Drive-Inventory-Collector-v4.2.2.exe`. Download the release asset, double-click it, and approve the administrator prompt. No PowerShell execution-policy change is required.

Automatic drive identification requires `smartctl.exe` from smartmontools. If it is missing, the graphical app offers to install smartmontools through WinGet. In Terminal, automatic scanning remains unavailable if smartmontools is missing; manual entry is still available.

Connect one drive at a time and wait for the result before removing it. Keep the workbook closed in spreadsheet software while collecting so the app can save it after each drive.

## Interfaces

The executable opens the graphical interface by default. Use the **Terminal** button to switch to the terminal workflow inside the same window. While Terminal is open, GUI scanning pauses; it resumes after Terminal closes and the workbook reloads. Switching to another tab does not close Terminal.

### Graphical Interface

The main window includes **Recorded Drives**, **Activity**, and **Terminal** tabs.

| Control | Action |
| --- | --- |
| **Pause Scanning** | Stops automatic drive checks until you resume scanning. |
| **Manual Drive Entry** | Opens the form to enter and review a drive record. |
| **Workbook Setup** | Selects optional identity columns and the workbook save location. |
| **Terminal** | Opens the terminal workflow in the main window. |
| **Version Information** | Opens the app version and technical details from the banner. |
| **Copy Path** | Copies the active workbook path to the clipboard. |
| **Open Folder** | Opens the workbook's containing folder in File Explorer. |
| **Finish** | Closes the collector. |
| Double-click a recorded row | Edits that saved record. |

Drive types can be searched in the manual-entry list. Choose **Other** to enter a custom type. **Copy Last Drive** is available inside Manual Drive Entry.

### Terminal Interface

Select the **Terminal** tab or **Open Terminal** to start the console-style workflow. Press a command key while scanning. For prompts, type a response in the input box and press Enter or select **Send**. Prompts, manual entry, workbook setup, and scan results appear in the output area. Select **Close Terminal** to return to automatic GUI scanning.

The same executable can also open a separate console window when started with `--terminal`:

```powershell
& '.\USB-Drive-Inventory-Collector-v4.2.2.exe' --terminal
```

### Terminal Keys

| Key | Action |
| --- | --- |
| `[M]` | Start Manual Drive Entry. |
| `[L]` | Copy the last saved drive and enter a new serial number. |
| `[S]` | Open Workbook Setup. |
| `[P]` | Pause or resume scanning. |
| `[D]` | Show Version Information. |
| `[H]` | Show the usage summary. |
| `[Q]` | Finish and close the collector. |
| `[Enter]` | Submit the current answer. |

### Settings and Values

| Setting or value | Behavior |
| --- | --- |
| Workbook save location | Defaults to `Output/Inventory.xlsx` beside the executable. The selected path is remembered for the current Windows account. |
| Workbook columns | Make, Model, Serial Number, Reported Capacity, and Type are included by default. Workbook Setup can add supported identity fields. |
| AutoPlay | If enabled, the app asks whether to disable it for the session. The previous setting is restored on normal exit. |
| Manual fields | Blank values are saved as `N/A`. Model and serial values are converted to uppercase. |
| Capacity | Enter a numeric value and select a unit: B, KB, MB, GB, TB, PB, or a custom unit. |
| Drive type | Select a listed type or choose Other and enter a custom value. |
| Duplicate serial | A record with a serial number already in the workbook cannot be saved again. |
| Drive read timeout | Each smartctl process has a 30-second timeout. |

## Drive Identification

The collector uses smartctl autodetection and USB transport fallbacks for supported USB adapters. Some bridges report their own identity or hide information about the drive behind them. When a read times out, check the adapter and reconnect it before trying again. The collector does not reset a USB controller.

## Files and Workbook

The default workbook is saved in `Output/Inventory.xlsx` beside the executable. You can choose another workbook location in Workbook Setup; the app remembers it for the current Windows account. Session logs are written to `Output/Logs/` beside the executable. The workbook uses the XLSX format and is written directly without Excel COM automation.

## Build from Source

Build on Windows with the .NET 8 SDK:

```powershell
dotnet publish WindowsApp/USBDriveInventoryCollector.csproj --configuration Release --runtime win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true --output publish
```

The [Build Windows GUI](.github/workflows/build-native.yml) workflow builds pull requests and uploads an unsigned preview executable. The [Publish Signed Release](.github/workflows/release.yml) workflow checks the version and release notes, signs the executable through Azure Artifact Signing, verifies its publisher, and attaches the executable and SHA-256 checksum to a GitHub release.

## License

[The Unlicense](UNLICENSE). The software is provided without warranty.

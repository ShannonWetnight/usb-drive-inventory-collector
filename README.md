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
- Displays a record number in the app without adding a column to the workbook.
- Restores the Recorded Drives layout with Reset View without changing workbook data.
- Uses distinct reading, error, and saved-record status colors in the graphical interface.

## Download and Run

The next tagged version is **4.2.9**. Signed Windows x64 builds are published on the [GitHub Releases page](https://github.com/ShannonWetnight/usb-drive-inventory-collector/releases). The build produces `USB-Drive-Inventory-Collector-v4.2.9.exe`. Download the release asset, double-click it, and approve the administrator prompt. No PowerShell execution-policy change is required.

Automatic drive identification requires `smartctl.exe` from smartmontools. If it is missing, the graphical app offers to install smartmontools through WinGet. In Terminal, automatic scanning remains unavailable if smartmontools is missing; manual entry is still available.

Connect one drive at a time and wait for the result before removing it. Keep the workbook closed in spreadsheet software while collecting so the app can save it after each drive.

## Interfaces

The executable opens the graphical interface by default. Select the **Terminal** tab, then choose **Enable Terminal** beside the tab to start the terminal workflow. The toggle reads **Disable Terminal** while active and remains available when switching tabs. GUI scanning pauses while Terminal runs; after Terminal closes, the workbook reloads and GUI scanning resumes unless scanning was paused. The **Pause Scanning** control and Terminal `[P]` command share the same pause state.

### Graphical Interface

The main window includes **Recorded Drives**, **Activity**, and **Terminal** tabs.

| Control | Action |
| --- | --- |
| **Pause Scanning** | Stops automatic drive checks until you resume scanning. |
| **Manual Drive Entry** | Opens the form to enter and review a drive record. |
| **Workbook Setup** | Selects optional identity columns, the workbook save location, and the logs save folder. |
| **Terminal** tab | Shows the embedded terminal panel, initially disabled. |
| **Enable Terminal / Disable Terminal** | Starts or stops the embedded terminal. Appears beside the tab when selected and remains visible while Terminal is running. |
| **Version Information** | Opens the app version and technical details from the banner. |
| Sound icon | Turns drive notification sounds on or off. The choice is saved for this Windows account. |
| **Copy Path** | Copies the active workbook path to the clipboard. |
| **Open Folder** | Opens the workbook's containing folder in File Explorer. |
| **Reset View** | Appears when the Recorded Drives view changes. Restores default grid widths, row heights, sort order, and scroll position without changing the workbook. |
| Refresh icon | Reloads the active workbook from disk and updates Recorded Drives. |
| **Finish** | Shows the saved workbook path with Copy Path, Open Folder, and Close controls. |
| Double-click a recorded row | Edits that saved record, including after sorting. |

Drive types can be searched in the manual-entry list. Choose **Other** to enter a custom type. **Copy Last Drive** is available inside Manual Drive Entry.

### Terminal Interface

Select the **Terminal** tab, then **Enable Terminal** to start the console-style workflow. Press a command key while scanning. For prompts, type a response in the input box and press Enter or select **Send**. Prompts, manual entry, workbook setup, and scan results appear in the output area. Select **Disable Terminal** to return to GUI scanning.

The same executable can also open a separate console window when started with `--terminal`:

```powershell
& '.\USB-Drive-Inventory-Collector-v4.2.9.exe' --terminal
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
| Logs save location | Defaults to `Output/Logs/` beside the executable. Choose another folder in Workbook Setup; it is remembered for the current Windows account. |
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

The default workbook is saved in `Output/Inventory.xlsx` beside the executable. Session logs default to `Output/Logs/`. Workbook Setup lets you choose both locations and remembers them for the current Windows account. Changing the logs folder moves the active GUI log there. The workbook uses the XLSX format and is written directly without Excel COM automation.

## Build from Source

Build on Windows with the .NET 8 SDK:

```powershell
dotnet publish WindowsApp/USBDriveInventoryCollector.csproj --configuration Release --runtime win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true --output publish
```

The [Build Windows GUI](.github/workflows/build-native.yml) workflow builds pull requests and uploads an unsigned preview executable. The [Publish Signed Release](.github/workflows/release.yml) workflow checks the version and release notes, signs the executable through Azure Artifact Signing, verifies its publisher, and attaches the executable and SHA-256 checksum to a GitHub release.

## License

[The Unlicense](UNLICENSE). The software is provided without warranty.

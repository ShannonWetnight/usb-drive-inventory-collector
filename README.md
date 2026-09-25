# USB Drive Inventory Collector
> **AI Workflow Notice**  
> The script was written through AI prompting, then reviewed and curated by the project maintainer. AI tools also helped with troubleshooting and documentation. Users should review the script and validate it in their own environment before relying on the collected data.

## Table of Contents
- [Overview](#overview)
- [Prerequisites](#prerequisites)
    + [Dependency Check](#dependency-check)
- [Operating Details](#operating-details)
- [Supported USB Adapters](#supported-usb-adapters)
- [Setup](#setup)
    + [Stop Windows AutoPlay prompts](#stop-windows-autoplay-prompts)
- [Output](#output)
- [Usage](#usage)
    + [Windows GUI (v4.0.0)](#windows-gui-v400)
    + [Console workflow](#console-workflow)
    + [Workbook column setup](#workbook-column-setup)
    + [Duplicate Serial Numbers](#duplicate-serial-numbers)
    + [Manual drive entry](#manual-drive-entry)
    + [Read-only Behavior](#read-only-behavior)
    + [Logs](#logs)
    + [Parameters](#parameters)
- [Limitations](#limitations)
- [Roadmap](#roadmap)
- [Tests](#tests)
- [License](#license)

## Overview
USB Drive Inventory Collector is a Windows PowerShell utility for recording drive identity from USB adapters or manual entry. Version 4.0.0 includes a Windows GUI and keeps the console script available. Both write the same local `.xlsx` workbook format and create a diagnostic log for each run.

It is useful anywhere you need a repeatable drive inventory: asset tracking, intake, audits, lab work, recycling preparation, or general hardware records. It does not erase, format, partition, or otherwise modify the attached drive.

## Prerequisites

- Windows
- Windows PowerShell 5.1 or later (an STA session for the GUI)
- Administrator privileges
- Windows Storage module (`Get-Disk`)
- smartmontools / `smartctl`

Microsoft Excel is not required. The workbook is written directly in XLSX/OpenXML format.

### Dependency Check

The collector checks for `smartctl` when it starts.

If smartmontools is missing and WinGet is available, it asks before installing anything. The GUI asks in a Windows dialog; the console shows:

```text
Install smartmontools now using WinGet? [Y/N]
```

If you approve the prompt, the script installs the `smartmontools.smartmontools` package and locates `smartctl.exe` for the current PowerShell session.

To disable the install prompt and exit when the dependency is missing:

```powershell
.\USB-Drive-Inventory-Collector.ps1 -NoDependencyInstallPrompt
```

## Operating Details

By default, each drive is written to `Output\Inventory.xlsx` with five fields:

| Column | Description |
| --- | --- |
| Make | Manufacturer inferred from the reported model when it can be identified reliably. Otherwise `N/A`. |
| Model | Model reported by the underlying drive. |
| Serial Number | Serial number reported by the underlying drive. |
| Reported Capacity | Manufacturer-style decimal capacity such as `256 GB`, `512 GB`, or `1 TB`. |
| Type | Media/interface classification based on information exposed by smartctl. |

Type can be reported as values such as:

- `M.2 NVMe SSD`
- `NVMe SSD`
- `M.2 SATA SSD`
- `mSATA SSD`
- `2.5-inch SATA SSD`
- `1.8-inch SATA SSD`
- `SATA SSD`
- `2.5-inch SATA HDD`
- `3.5-inch SATA HDD`
- `SATA HDD`
- `SATA Drive`
- `SSD`
- `HDD`
- `N/A`

The collector reports a specific form factor when the drive exposes it or an exact known model identifies it. An NVMe drive is reported as `NVMe SSD` when its physical form factor is unavailable.

ATA does not automatically mean SATA. The collector checks SATA metadata and the original smartctl identity text for explicit PATA transport information. PATA drives can be reported as `PATA Drive`, `PATA HDD`, or `PATA SSD`, with a form factor when known. `WD800AAJB` has a specific fallback to `3.5-inch PATA HDD` when older identify data omits the interface or media details. Other ATA drives without enough evidence are recorded as `ATA Drive`, `ATA HDD`, or `ATA SSD`. IDE and PATA refer to the same interface family. Existing workbook rows are not reclassified automatically.

Manual entry offers these types plus 2.5-inch and 3.5-inch IDE HDDs, generic IDE drives, IDE HDDs and SSDs, SAS SSDs and HDDs, USB flash drives, SD and microSD cards, CompactFlash cards, eMMC, 3.5-inch and 5.25-inch floppy disks, and a custom **Other** choice. The added options are for manual records; they do not change what smartctl can identify automatically.

## Supported USB Adapters

Windows often sees a USB enclosure or bridge instead of the drive behind it. To get the underlying model and serial number, the collector uses smartmontools and tries several read-only transport methods.

Current probing covers:

- standard smartctl autodetection
- SAT/SATA-to-USB passthrough
- JMicron NVMe-to-USB passthrough
- Realtek NVMe-to-USB passthrough
- ASMedia NVMe-to-USB passthrough
- common older USB-to-SATA bridge modes supported by smartmontools

This should cover many NVMe-to-USB enclosures and SATA-to-USB adapters, but USB bridge behavior varies by chipset and firmware. If an adapter does not expose the drive identity through a transport smartctl understands, the collector will not substitute the adapter's serial number for the drive's. Missing information remains `N/A`, or the read fails with details in the run log.

## Setup

1. Download `USB-Drive-Inventory-Collector.ps1` and `USB-Drive-Inventory-Collector-GUI.ps1` into the same folder.
2. Open Windows PowerShell as Administrator.
3. Change to the folder containing the script.
4. If the local execution policy blocks the script, allow it for the current PowerShell process:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
```

5. Start the GUI:

```powershell
.\USB-Drive-Inventory-Collector-GUI.ps1
```

To use the console workflow, start `USB-Drive-Inventory-Collector.ps1` instead.

No output folders need to be created manually.

### Stop Windows AutoPlay prompts

If AutoPlay is enabled for the signed-in Windows user, the collector asks at startup:

```text
AutoPlay can open drive folders or show pop-ups. Disable it while collecting? [Y/N]
```

Choose `Y` to turn off that user's AutoPlay preference for the run. The collector remembers whether the setting existed and what value it held, then restores it when the script exits. Choose `N` to leave it alone. If AutoPlay is already off, there is no prompt. The script compares account SIDs before making the change and skips it if the elevated account is different from the signed-in desktop user.

This preference controls Windows AutoPlay, including the usual open-folder action. A workplace policy can override it, and another application can still open a folder or show its own prompt. The collector also suppresses critical device error dialogs raised by its own process while it runs. If Windows keeps opening folders, check **Settings → Bluetooth & devices → AutoPlay** or ask the workstation administrator about policy.

Restoration runs during normal exit, including `Ctrl+C` and handled errors. If PowerShell is forcibly terminated or the computer loses power, cleanup cannot run; check the AutoPlay setting before the next batch. If the preference changes away from the temporary value during collection, the script leaves that newer value in place instead of overwriting it.

## Output

The default layout is created beside the scripts:

```text
USB-Drive-Inventory-Collector.ps1
USB-Drive-Inventory-Collector-GUI.ps1
Output\
  Inventory.xlsx
  Logs\
    USB-Drive-Inventory-Collector-YYYYMMDD-HHMMSS.log
```

`Output\` is excluded by this repository's `.gitignore` so serial numbers and collected inventory data are not accidentally committed.

## Usage

### Windows GUI (v4.0.0)

Launch the GUI from an elevated STA PowerShell window. It starts scanning after startup, and the status line reports detection, saving, duplicates, removals, and read errors. **Recorded drives** shows existing rows and selected workbook columns; **Activity** shows recent events; **Details** shows paths, smartctl version, settings, and the selected columns. The debug log on disk keeps the full probe history.

Use **Pause scanning** to stop new scans; a probe already running will finish. **Manual entry** opens a form with the same field validation, `N/A` handling, capacity units, categorized drive types, and custom **Other** choices as the console. Review before saving; from review you can edit, save and start another record, or save and copy everything except the serial. A duplicate serial gives you the choice to change it or cancel. **Copy last** starts with the last saved drive's five standard fields and an empty serial. **Workbook setup** selects optional identity columns and makes a backup before a layout change. **Finish** waits for a probe in progress, restores the temporary AutoPlay setting, and closes the window.

The GUI runs drive probes in a background PowerShell runspace, so slow smartctl calls do not block the controls. Manual entry and setup can be opened during a probe; the result is processed after the dialog closes. The collector handles one newly inserted disk per scan and requires removal before retrying a disk that failed to read.

The GUI uses the drive classification, smartctl transport fallbacks, direct XLSX writer, logging, and AutoPlay restoration functions in the adjacent console script. Keep both files together. Windows Forms adds no Excel dependency or separate GUI package.

### Console workflow

1. Start the script in an elevated PowerShell window.
2. Answer the AutoPlay prompt if it appears.
3. Connect a drive through a USB adapter or enclosure.
4. Wait for the drive to be recorded.
5. Remove it after the script reports that the record was saved.
6. Connect the next drive.
7. Press `Ctrl+C` when finished.

The waiting screen shows the version, maintainer, repository link, and a numbered Usage section. Press `[D]` while waiting to see the full output and log paths, USB adapter scope, workbook backend, smartctl version, and probe timeout; press `[D]` again to hide them. Press `[M]` for manual entry. After a read failure or drive removal, the console reminds you that `[M]` opens manual entry. The console clears when a new drive is detected so the current result is easy to read; the log keeps the run history. Press `[Ctrl+C]` when finished.

### Workbook column setup

In the GUI, click **Workbook setup**. In the console, press `[S]` while waiting. To configure columns before any connected drive is probed, start either entry point with `-SetupOnStartup`:

```powershell
.\USB-Drive-Inventory-Collector-GUI.ps1 -SetupOnStartup
# Or: .\USB-Drive-Inventory-Collector.ps1 -SetupOnStartup
```

The five default columns stay in place. You can add Interface, Firmware Version, Model Family, Form Factor, Rotation Rate (RPM), Capacity (Bytes), Logical and Physical Sector Sizes, ATA Version, SATA Version, Reported Protocol, and Probe Transport. These values come from the identity query already used by the collector. Setup does not run SMART health tests or collect every vendor-specific attribute.

Enter a field number to toggle it, `[A]` to select all extra fields, or `[D]` for the default layout. `[Y]` applies the selection; `[C]` cancels it. Before changing columns, the collector copies the workbook to a file named `Inventory.xlsx.before-setup-<unique ID>.xlsx` beside the original. Removing a column excludes its data from the active workbook; the backup retains it.

The workbook headers remember the selection on the next run. Existing values in retained columns survive later saves. Newly added columns contain `N/A` for older records, manual records, and details the adapter does not expose. Copying a manual record copies the five core fields and asks for a new serial; optional identity details remain `N/A`. Automatic collection pauses while setup is open.

Use a separate output workbook for a different collection layout. Unsupported or duplicate headers stop the collector before it overwrites the workbook.

### Manual drive entry

Press `M` while the collector is polling to enter a drive manually. If your PowerShell host does not support direct console keys, start the script with `-ManualEntryOnStartup` instead. A key pressed during a drive probe is handled when the script returns to the polling loop.

The form asks for Make, Model, Serial Number, a numeric capacity and unit, and drive type. Capacity accepts positive whole numbers and decimals, such as `0.005`; enter the unit on the next screen. Capacity units include `B`, `KB`, `MB`, `GB`, `TB`, `PB`, and **Other**, which lets you enter a custom unit. Drive types have numbered choices grouped under **Standard**, **Enterprise**, and **Other**, sorted within each group; the final **Other** choice accepts a custom type. The form accepts plain letters, digits, spaces, and limited punctuation. Leading and trailing spaces are removed; Model and serial are converted to uppercase, while Make keeps the case you enter. For example, an amount of `2` with unit `TB` is saved as `2 TB`. Press Enter (or enter only spaces) to save `N/A` for a field. Skipping either the capacity number or unit saves the capacity as `N/A`.

The console clears between fields, the capacity amount and unit, review, and the saved record. Before saving, review the five fields. Choose `Y` to save and then decide whether to add another drive, copy it, or return to automatic collection. Choose `L` to save and immediately copy that drive for the next serial number, `E` to edit a field, or `C` to cancel. Both save options check for duplicates and wait for a successful workbook write before starting another record. When editing, `:back` or `:cancel` returns to review without changing the field. Submitting an edited value also returns to review. During initial entry, `:cancel` leaves manual entry without saving that drive. If the serial already exists in the workbook, choose `S` to enter a different serial for the current record or `C` to cancel it. The save options return when the serial is unique. `N/A` is accepted as a serial, but it cannot be checked for duplicates. A failed workbook save leaves the form open for another attempt.

When you open manual entry and the workbook already has a drive, choose `N` for a new record or `L` to copy the last saved drive. Copying fills in Make, Model, Capacity, and Type, then asks for a new serial number and shows the full review before saving. After saving, choose `A` for another new drive, `L` to copy the one you just saved with a new serial, or `R` to resume automatic collection.

During each insertion, the collector:

1. Finds USB physical disks reported by Windows.
2. Excludes disks marked as boot or system disks.
3. Gives the USB bridge a short period to initialize.
4. Probes supported smartctl transports.
5. Reads the underlying drive identity when available.
6. Appends the five standard fields and any selected extra fields to `Inventory.xlsx`.
7. Saves the workbook before waiting for the next drive.

The script waits for removal before treating another device on the same Windows disk number as a new insertion.
After a read failure, it waits for removal and reinsertion before trying that disk number again.

### Duplicate Serial Numbers

A serial number already present in the workbook is not added again. The collector prints the detected information, records the duplicate in the log, and leaves the workbook unchanged.

If a drive reports no serial number, the value is stored as `N/A`. The collector cannot use serial-based duplicate detection to distinguish two drives that both report `N/A`.

### Read-only Behavior

Automatic collection uses Windows disk metadata and smartctl identity queries. Manual collection uses the values you enter. Neither mode issues commands to:

- erase or sanitize a drive
- format a drive
- create or remove partitions
- mount or dismount volumes
- change SMART settings
- write files to the attached drive

Windows boot and system disks are excluded from collection.

### Logs

Each run creates a timestamped log under `Output\Logs` containing information useful for troubleshooting, including:

- script and PowerShell version
- dependency checks
- USB disk detection and removal
- smartctl transport probes and exit codes
- detected protocol and drive identity
- duplicate detection
- workbook save attempts
- exception type, HRESULT, stack position, and inner exceptions when available

If a drive read or workbook update fails, the console prints the path to that run's log. The GUI lists the error in **Activity** and shows the log path in **Details**.

### Parameters

The defaults are enough for normal use. Both entry points accept paths and polling settings:

```powershell
.\USB-Drive-Inventory-Collector.ps1 `
    -OutputPath "C:\Inventory\Inventory.xlsx" `
    -LogPath "C:\Inventory\collector.log" `
    -UsbDevicePattern "*" `
    -PollSeconds 1
```

| Parameter | Purpose |
| --- | --- |
| `OutputPath` | Overrides the default `Output\Inventory.xlsx` path. |
| `LogPath` | Overrides the default timestamped log path. |
| `UsbDevicePattern` | Filters Windows USB disk friendly names. Default is `*`. |
| `PollSeconds` | Controls how often Windows is checked for connected USB disks. |
| `SmartctlRetries` | Sets the number of full transport-probe cycles. |
| `InitialSettleMilliseconds` | Sets the delay after Windows first detects a drive before probing begins. |
| `RetryDelayMilliseconds` | Sets the delay between smartctl retry cycles. |
| `WorkbookSaveRetries` | Sets the number of workbook save attempts. |
| `WorkbookRetryDelayMilliseconds` | Sets the delay between workbook save attempts. |
| `SmartctlTimeoutSeconds` | Maximum time for each smartctl process; default 30 seconds. A stalled probe is stopped, and the drive is skipped until removal and reinsertion. |
| `NoDependencyInstallPrompt` | Exits instead of offering to install smartmontools when it is missing. |
| `SetupOnStartup` | Opens workbook column setup before probing connected drives. |
| `ManualEntryOnStartup` | Opens manual entry after startup; useful if the `M` console hotkey is unavailable. |

## Limitations

- USB bridge behavior is not standardized. Some adapters will expose more information than others.
- A bridge may report its own identity instead of the underlying drive. The collector rejects common bridge-style identities when possible and tries alternate smartctl transports.
- Form factor is not always exposed. In that case Type stays broader, such as `NVMe SSD` or `SATA SSD`.
- Manufacturer detection is conservative. An unknown model prefix returns `N/A` instead of a guessed manufacturer.
- Keep `Inventory.xlsx` closed while collecting. The script replaces the workbook file when saving an update.
- If a USB bridge stops responding, the collector stops an overdue smartctl process and logs the timeout. This does not reset the bridge's hardware. Unplug and reconnect a stuck adapter, then reinsert the drive. Windows device restart commands can reset a specific Plug and Play device, but restarting a shared hub or controller can interrupt other attached devices, and a restart is not guaranteed to cycle USB port power.
- Windows may take longer than the polling interval to register removal. Wait for the console's removal message before inserting the next drive; a swap that occurs entirely between polls may be missed.
- Automatic scanning pauses while a manual form is open. If you insert a USB drive then, it will be considered for automatic collection after you return to the polling loop.
- The collector records drive information only. It does not perform any follow-up action on the hardware.

## Roadmap

- Investigate safe recovery for USB adapters that stop responding. Restarting a shared controller could interrupt unrelated devices, so the current collector reports the timeout and waits for the adapter to be reconnected.

## Tests

Run `pwsh -NoProfile -File tests/Collector.Tests.ps1` and `pwsh -NoProfile -File tests/GUI.Tests.ps1` (or use `powershell` on Windows). These tests load functions without accessing attached drives. They cover classification, workbook round trips, setup, manual validation, and an asynchronous scan with a simulated disk. The Windows Forms interface still needs a real Windows workstation and USB adapter for final UI testing.

## License

This project is released under [The Unlicense](UNLICENSE). You may use, modify, share, or sell it without attribution. The software is provided without warranty.

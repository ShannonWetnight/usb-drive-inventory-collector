# USB Drive Inventory Collector
> **AI Workflow Notice**  
> Parts of this project were developed with assistance from AI tools for scripting, troubleshooting, and documentation drafting. The project maintainer reviews and curates changes before publishing them. AI assistance does not replace testing, so users should review the script and validate it in their own environment before relying on the collected data.

USB Drive Inventory Collector is a Windows PowerShell utility for collecting basic identity information from drives connected through USB adapters or enclosures. It writes the results to a local `.xlsx` workbook and keeps a diagnostic log for each run.

It is useful anywhere you need a repeatable drive inventory: asset tracking, intake, audits, lab work, recycling preparation, or general hardware records. It does not erase, format, partition, or otherwise modify the attached drive.

## Prerequisites

- Windows
- Windows PowerShell 5.1 or later
- Administrator privileges
- Windows Storage module (`Get-Disk`)
- smartmontools / `smartctl`

Microsoft Excel is not required. The workbook is written directly in XLSX/OpenXML format.

### Dependency Check

The collector checks for `smartctl` when it starts.

If smartmontools is missing and WinGet is available, it asks before installing anything:

```text
Install smartmontools now using WinGet? [Y/N]
```

If you approve the prompt, the script installs the `smartmontools.smartmontools` package and locates `smartctl.exe` for the current PowerShell session.

To disable the install prompt and exit when the dependency is missing:

```powershell
.\USB-Drive-Inventory-Collector.ps1 -NoDependencyInstallPrompt
```

## Operating Details

Each drive is written to `Output\Inventory.xlsx` with five fields:

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

The collector only reports a specific form factor when the drive or adapter exposes enough information to support it. An NVMe drive, for example, is reported as `NVMe SSD` rather than assumed to be M.2 when its physical form factor is unavailable.

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

1. Download `USB-Drive-Inventory-Collector.ps1` and place it in its own folder.
2. Open Windows PowerShell as Administrator.
3. Change to the folder containing the script.
4. If the local execution policy blocks the script, allow it for the current PowerShell process:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
```

5. Start the collector:

```powershell
.\USB-Drive-Inventory-Collector.ps1
```

No output folders need to be created manually.

## Output

The default layout is created beside the script:

```text
USB-Drive-Inventory-Collector.ps1
Output\
  Inventory.xlsx
  Logs\
    USB-Drive-Inventory-YYYYMMDD-HHMMSS.log
```

`Output\` is excluded by this repository's `.gitignore` so serial numbers and collected inventory data are not accidentally committed.

## Usage

1. Start the script in an elevated PowerShell window.
2. Connect a drive through a USB adapter or enclosure.
3. Wait for the drive to be recorded.
4. Remove it after the script reports that the record was saved.
5. Connect the next drive.
6. Press `Ctrl+C` when finished.

During each insertion, the collector:

1. Finds USB physical disks reported by Windows.
2. Excludes disks marked as boot or system disks.
3. Gives the USB bridge a short period to initialize.
4. Probes supported smartctl transports.
5. Reads the underlying drive identity when available.
6. Appends Make, Model, Serial Number, Reported Capacity, and Type to `Inventory.xlsx`.
7. Saves the workbook before waiting for the next drive.

The script waits for removal before treating another device on the same Windows disk number as a new insertion.

### Duplicate Serial Numbers

A serial number already present in the workbook is not added again. The collector prints the detected information, records the duplicate in the log, and leaves the workbook unchanged.

If a drive reports no serial number, the value is stored as `N/A`. The collector cannot use serial-based duplicate detection to distinguish two drives that both report `N/A`.

### Read-only Behavior

The collector uses Windows disk metadata and smartctl identity queries. It does not issue commands to:

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

If a drive read or workbook update fails, the console prints the path to that run's log.

### Parameters

The defaults are enough for normal use. Paths and polling behavior can also be overridden:

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
| `NoDependencyInstallPrompt` | Exits instead of offering to install smartmontools when it is missing. |

## Limitations

- USB bridge behavior is not standardized. Some adapters will expose more information than others.
- A bridge may report its own identity instead of the underlying drive. The collector rejects common bridge-style identities when possible and tries alternate smartctl transports.
- Form factor is not always exposed. In that case Type stays broader, such as `NVMe SSD` or `SATA SSD`.
- Manufacturer detection is conservative. An unknown model prefix returns `N/A` instead of a guessed manufacturer.
- Keep `Inventory.xlsx` closed while collecting. The script replaces the workbook file when saving an update.
- The collector records drive information only. It does not perform any follow-up action on the hardware.

<#
.SYNOPSIS
    General-purpose USB drive inventory collector for Windows.

.DESCRIPTION
    Watches for USB-connected physical drives and records media identity for
    inventory, asset tracking, auditing, and other drive-identification workflows.

    For each newly detected USB drive, the collector:
      - Uses smartctl transport autodetection plus safe read-only fallbacks for
        common NVMe-to-USB and SATA-to-USB bridge families.
      - Captures only:
            Make
            Model
            Serial Number
            Reported Capacity
            Type
      - Writes N/A for information that cannot be determined reliably.
      - Appends one row to an .xlsx workbook and saves immediately.
      - Prevents duplicate serial numbers from being added.

    Windows boot/system disks are explicitly excluded. All smartctl operations
    used by this script are identity/read operations; the script does not issue
    erase, format, write, or SMART configuration commands.

    The workbook is written directly in the XLSX/OpenXML format. Microsoft
    Excel is not launched or required.

    Output layout beside the script:
        Output\Inventory.xlsx
        Output\Logs\USB-Drive-Inventory-Collector-YYYYMMDD-HHMMSS.log

    The collector checks for smartmontools/smartctl at startup. If smartctl is
    missing and WinGet is available, the user is prompted before any install is
    attempted.

.REQUIREMENTS
    - Windows PowerShell 5.1 or later
    - Administrator privileges
    - smartmontools (the script can offer to install it with WinGet)
#>

param (
    [string]$OutputPath,

    [string]$LogPath,

    # Optional friendly-name filter. Default '*' allows any USB drive or adapter.
    [string]$UsbDevicePattern = "*",

    [int]$PollSeconds = 1,

    [ValidateRange(1, 5)]
    [int]$SmartctlRetries = 2,

    [ValidateRange(250, 10000)]
    [int]$InitialSettleMilliseconds = 2000,

    [ValidateRange(250, 10000)]
    [int]$RetryDelayMilliseconds = 1250,

    [ValidateRange(1, 10)]
    [int]$WorkbookSaveRetries = 5,

    [ValidateRange(100, 10000)]
    [int]$WorkbookRetryDelayMilliseconds = 750,

    # Maximum time to wait for an individual smartctl process.
    [ValidateRange(2, 120)]
    [int]$SmartctlTimeoutSeconds = 15,

    # Suppresses the install prompt and exits if smartctl is missing.
    [switch]$NoDependencyInstallPrompt,

    # Opens manual entry immediately (also available with M while polling).
    [switch]$ManualEntryOnStartup
)

$ErrorActionPreference = "Stop"
$ScriptVersion = "3.5.3"
$RunId = [guid]::NewGuid().ToString("N").Substring(0, 8)
$script:PreferredTransportByDiskNumber = @{}

$DefaultOutputDirectory = Join-Path -Path $PSScriptRoot -ChildPath "Output"

if (-not (Test-Path -LiteralPath $DefaultOutputDirectory)) {
    New-Item -ItemType Directory -Path $DefaultOutputDirectory -Force | Out-Null
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path -Path $DefaultOutputDirectory -ChildPath "Inventory.xlsx"
}

$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
$OutputDirectory = Split-Path -Parent $OutputPath

if (-not (Test-Path -LiteralPath $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogDirectory = Join-Path -Path $DefaultOutputDirectory -ChildPath "Logs"

    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }

    $LogPath = Join-Path -Path $LogDirectory -ChildPath (
        "USB-Drive-Inventory-Collector-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss")
    )
}
else {
    $LogPath = [System.IO.Path]::GetFullPath($LogPath)
    $LogDirectory = Split-Path -Parent $LogPath

    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }
}

function Write-Log {

    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("DEBUG", "INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )

    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $Line = "[$Timestamp][$Level][Run:$RunId] $Message"

    try {
        [System.IO.File]::AppendAllText(
            $LogPath,
            $Line + [Environment]::NewLine,
            [System.Text.UTF8Encoding]::new($false)
        )
    }
    catch {
        # Logging must never terminate the inventory collector.
    }
}


function Write-ExceptionLog {

    param (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,

        [string]$Context = "Unhandled exception"
    )

    $Exception = $ErrorRecord.Exception
    $ExceptionType = if ($null -ne $Exception) { $Exception.GetType().FullName } else { "N/A" }
    $Message = if ($null -ne $Exception) { $Exception.Message } else { [string]$ErrorRecord }
    $HResult = "N/A"

    if ($null -ne $Exception) {
        try {
            $HResult = '0x' + $Exception.HResult.ToString('X8')
        }
        catch {}
    }

    Write-Log -Level ERROR -Message (
        "{0} | Type={1} | HResult={2} | FQID={3} | Category={4} | Message={5}" -f
        $Context,
        $ExceptionType,
        $HResult,
        $ErrorRecord.FullyQualifiedErrorId,
        $ErrorRecord.CategoryInfo,
        $Message
    )

    if (-not [string]::IsNullOrWhiteSpace($ErrorRecord.ScriptStackTrace)) {
        Write-Log -Level ERROR -Message ("ScriptStackTrace: " + $ErrorRecord.ScriptStackTrace)
    }

    if (
        $null -ne $ErrorRecord.InvocationInfo -and
        -not [string]::IsNullOrWhiteSpace($ErrorRecord.InvocationInfo.PositionMessage)
    ) {
        Write-Log -Level ERROR -Message (
            "Position: " + ($ErrorRecord.InvocationInfo.PositionMessage -replace "[\r\n]+", " ")
        )
    }

    $Inner = if ($null -ne $Exception) { $Exception.InnerException } else { $null }
    $Depth = 0

    while ($null -ne $Inner -and $Depth -lt 5) {
        Write-Log -Level ERROR -Message (
            "InnerException[{0}]: {1}: {2}" -f $Depth, $Inner.GetType().FullName, $Inner.Message
        )
        $Inner = $Inner.InnerException
        $Depth++
    }
}



function Test-IsAdministrator {

    $Identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $Principal = [Security.Principal.WindowsPrincipal]::new($Identity)

    return $Principal.IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator
    )
}


function Find-SmartctlExecutable {

    $Command = Get-Command smartctl -ErrorAction SilentlyContinue

    if ($null -ne $Command -and -not [string]::IsNullOrWhiteSpace($Command.Source)) {
        return $Command.Source
    }

    $Candidates = [System.Collections.ArrayList]::new()

    if (-not [string]::IsNullOrWhiteSpace($env:ProgramFiles)) {
        [void]$Candidates.Add(
            (Join-Path -Path $env:ProgramFiles -ChildPath "smartmontools\bin\smartctl.exe")
        )
    }

    $ProgramFilesX86 = [Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFilesX86)
    if (-not [string]::IsNullOrWhiteSpace($ProgramFilesX86)) {
        [void]$Candidates.Add(
            (Join-Path -Path $ProgramFilesX86 -ChildPath "smartmontools\bin\smartctl.exe")
        )
    }

    foreach ($Candidate in $Candidates) {
        if (
            -not [string]::IsNullOrWhiteSpace($Candidate) -and
            (Test-Path -LiteralPath $Candidate)
        ) {
            return $Candidate
        }
    }

    return $null
}


function Ensure-SmartctlDependency {

    $Existing = Find-SmartctlExecutable

    if (-not [string]::IsNullOrWhiteSpace($Existing)) {
        return $Existing
    }

    Write-Host ""
    Write-Host "DEPENDENCY REQUIRED"
    Write-Host "-------------------"
    Write-Host "smartmontools (smartctl) is required to read drive identity through USB adapters."
    Write-Host ""

    Write-Log -Level WARN -Message "Dependency check: smartctl was not found."

    if ($NoDependencyInstallPrompt) {
        throw "smartctl is missing and -NoDependencyInstallPrompt was specified."
    }

    $Winget = Get-Command winget -ErrorAction SilentlyContinue

    if ($null -eq $Winget) {
        throw "smartctl is missing and WinGet is not available for automatic installation. Install smartmontools manually, then rerun the script."
    }

    $Response = Read-Host "Install smartmontools now using WinGet? [Y/N]"

    if ($Response -notmatch '(?i)^y(?:es)?$') {
        throw "smartctl is required. Installation was declined."
    }

    Write-Host ""
    Write-Host "Installing smartmontools..."
    Write-Log -Level INFO -Message "User approved smartmontools installation with WinGet."

    & $Winget.Source install `
        --id smartmontools.smartmontools `
        -e `
        --source winget `
        --accept-package-agreements `
        --accept-source-agreements

    $InstallExitCode = $LASTEXITCODE

    Write-Log -Level INFO -Message "WinGet smartmontools install completed with exitCode=$InstallExitCode."

    if ($InstallExitCode -ne 0) {
        throw "WinGet could not install smartmontools (exit code $InstallExitCode)."
    }

    # Refresh PATH for this PowerShell process, then check standard install paths.
    try {
        $MachinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
        $UserPath = [Environment]::GetEnvironmentVariable("Path", "User")
        $env:Path = "$MachinePath;$UserPath"
    }
    catch {}

    $Installed = Find-SmartctlExecutable

    if ([string]::IsNullOrWhiteSpace($Installed)) {
        throw "smartmontools installed, but smartctl.exe could not be located. Close and reopen PowerShell, then rerun the script."
    }

    Write-Host "smartmontools installed successfully."
    Write-Host ""

    return $Installed
}


function ConvertTo-CommandLineArgument {

    param (
        [string]$Value
    )

    if ($null -eq $Value) {
        return '""'
    }

    if ($Value -notmatch '[\s"]') {
        return $Value
    }

    return '"' + ($Value -replace '"', '\\"') + '"'
}


function Invoke-SmartctlProcess {

    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments,

        [string]$Context = "smartctl"
    )

    $StartInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $StartInfo.FileName = $script:SmartctlPath
    $StartInfo.Arguments = (($Arguments | ForEach-Object { ConvertTo-CommandLineArgument $_ }) -join " ")
    $StartInfo.UseShellExecute = $false
    $StartInfo.RedirectStandardOutput = $true
    $StartInfo.RedirectStandardError = $true
    $StartInfo.CreateNoWindow = $true

    $Process = [System.Diagnostics.Process]::new()
    $Process.StartInfo = $StartInfo

    try {
        Write-Log -Level DEBUG -Message (
            "{0}: {1} {2}" -f $Context, $script:SmartctlPath, $StartInfo.Arguments
        )

        if (-not $Process.Start()) {
            throw "smartctl process did not start."
        }

        # Read both pipes concurrently so a full stderr pipe cannot deadlock
        # while stdout is being drained. Bound the wait for a stalled bridge.
        $StdOutTask = $Process.StandardOutput.ReadToEndAsync()
        $StdErrTask = $Process.StandardError.ReadToEndAsync()

        if (-not $Process.WaitForExit($SmartctlTimeoutSeconds * 1000)) {
            try { $Process.Kill() } catch {}
            try { $Process.WaitForExit(2000) | Out-Null } catch {}
            throw [System.TimeoutException]::new("smartctl timed out after $SmartctlTimeoutSeconds seconds ($Context). The USB adapter may need to be unplugged and reconnected.")
        }

        $StdOut = $StdOutTask.GetAwaiter().GetResult()
        $StdErr = $StdErrTask.GetAwaiter().GetResult()

        return [PSCustomObject]@{
            ExitCode = $Process.ExitCode
            StdOut   = [string]$StdOut
            StdErr   = [string]$StdErr
        }
    }
    finally {
        $Process.Dispose()
    }
}


function Get-SmartctlScanDeviceType {

    param (
        [int]$DiskNumber
    )

    try {
        $Result = Invoke-SmartctlProcess `
            -Arguments @("--scan-open", "-j") `
            -Context "smartctl scan-open"

        if ([string]::IsNullOrWhiteSpace($Result.StdOut)) {
            return $null
        }

        $Json = $Result.StdOut | ConvertFrom-Json
        $DeviceName = "/dev/pd$DiskNumber"

        foreach ($ScannedDevice in @($Json.devices)) {
            if ($null -eq $ScannedDevice) {
                continue
            }

            $Name = Get-CleanValue $ScannedDevice.name
            $InfoName = Get-CleanValue $ScannedDevice.info_name

            if (
                $Name -eq $DeviceName -or
                $InfoName -match ("(?i)PhysicalDrive{0}(?:\D|$)" -f $DiskNumber)
            ) {
                $DetectedType = Get-CleanValue $ScannedDevice.type
                $DetectedProtocol = Get-CleanValue $ScannedDevice.protocol

                Write-Log -Level DEBUG -Message (
                    "scan-open matched disk={0}: name='{1}', type='{2}', protocol='{3}', infoName='{4}'" -f
                    $DiskNumber, $Name, $DetectedType, $DetectedProtocol, $InfoName
                )

                if ($DetectedType -ne "N/A") {
                    return $DetectedType
                }
            }
        }
    }
    catch {
        Write-ExceptionLog -ErrorRecord $_ -Context "smartctl --scan-open discovery failed for disk $DiskNumber"
        if ($_.Exception -is [System.TimeoutException]) { throw }
    }

    return $null
}


function Get-SmartctlTransportCandidates {

    param (
        [int]$DiskNumber
    )

    $Candidates = [System.Collections.ArrayList]::new()

    function Add-TransportCandidate {
        param ([string]$Value)

        if ([string]::IsNullOrWhiteSpace($Value)) {
            return
        }

        if (-not $Candidates.Contains($Value)) {
            [void]$Candidates.Add($Value)
        }
    }

    if ($script:PreferredTransportByDiskNumber.ContainsKey($DiskNumber)) {
        Add-TransportCandidate $script:PreferredTransportByDiskNumber[$DiskNumber]
    }

    $ScannedType = Get-SmartctlScanDeviceType -DiskNumber $DiskNumber

    if (
        -not [string]::IsNullOrWhiteSpace($ScannedType) -and
        $ScannedType -notmatch '(?i)^(auto|scsi)$'
    ) {
        Add-TransportCandidate $ScannedType
    }

    # 'auto' means omit -d and let smartctl perform its normal autodetection.
    Add-TransportCandidate "auto"

    # SATA-to-USB / SAT bridges.
    Add-TransportCandidate "sat"

    # NVMe-to-USB bridge families supported by smartmontools. The /sat forms
    # allow the same family to fall back to ATA/SATA media when supported.
    Add-TransportCandidate "sntjmicron"
    Add-TransportCandidate "sntjmicron/sat"
    Add-TransportCandidate "sntrealtek"
    Add-TransportCandidate "sntrealtek/sat"
    Add-TransportCandidate "sntasmedia"
    Add-TransportCandidate "sntasmedia/sat"

    # Older/common SATA USB pass-through families.
    Add-TransportCandidate "usbjmicron"
    Add-TransportCandidate "usbprolific"
    Add-TransportCandidate "usbsunplus"
    Add-TransportCandidate "usbcypress"

    return @($Candidates)
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem


function Get-CleanValue {

    param (
        [object]$Value
    )

    if ($null -eq $Value) {
        return "N/A"
    }

    $Text = ([string]$Value).Trim()

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return "N/A"
    }

    return $Text
}


function Get-ReportedCapacity {

    param (
        [object]$Bytes
    )

    if ($null -eq $Bytes) {
        return "N/A"
    }

    try {
        $Capacity = [decimal]$Bytes
    }
    catch {
        return "N/A"
    }

    if ($Capacity -le 0) {
        return "N/A"
    }

    # Drive manufacturers commonly report decimal capacity units:
    # 1 GB = 1,000,000,000 bytes
    # 1 TB = 1,000,000,000,000 bytes

    if ($Capacity -ge 1000000000000) {
        $TB = $Capacity / 1000000000000
        return ("{0:0.##} TB" -f $TB)
    }

    $GB = [math]::Round(
        [double]($Capacity / 1000000000),
        0
    )

    return "$GB GB"
}


function Get-DriveType {

    param (
        [object]$SmartctlJson,
        [string]$Model = "N/A"
    )

    if ($null -eq $SmartctlJson) {
        return "N/A"
    }

    $Protocol = ""
    $FormFactor = ""
    $RotationRate = $null

    if ($null -ne $SmartctlJson.device -and $null -ne $SmartctlJson.device.protocol) {
        $Protocol = ([string]$SmartctlJson.device.protocol).Trim()
    }

    if ($null -ne $SmartctlJson.form_factor -and $null -ne $SmartctlJson.form_factor.name) {
        $FormFactor = ([string]$SmartctlJson.form_factor.name).Trim()
    }

    if ($null -ne $SmartctlJson.rotation_rate) {
        try {
            $RotationRate = [int64]$SmartctlJson.rotation_rate
        }
        catch {}
    }

    $IsSolidState = $false
    $IsRotational = $false

    if ($null -ne $RotationRate) {
        if ($RotationRate -eq 0) {
            $IsSolidState = $true
        }
        elseif ($RotationRate -gt 0) {
            $IsRotational = $true
        }
    }

    if (-not $IsSolidState -and $Model -match "(?i)SSD|Solid[ _-]?State") {
        $IsSolidState = $true
    }

    if (
        $Protocol -match "(?i)NVMe" -or
        $null -ne $SmartctlJson.nvme_version
    ) {
        switch -Regex ($FormFactor) {
            "(?i)^M\.2$"           { return "M.2 NVMe SSD" }
            "(?i)^2\.5 inches$"   { return "2.5-inch NVMe SSD" }
            "(?i)^1\.8 inches$"   { return "1.8-inch NVMe SSD" }
            "(?i)^< 1\.8 inches$" { return "<1.8-inch NVMe SSD" }
            default                 { return "NVMe SSD" }
        }
    }

    if (
        $Protocol -match "(?i)ATA|SATA" -or
        $null -ne $SmartctlJson.ata_version -or
        $null -ne $SmartctlJson.sata_version
    ) {
        if ($IsSolidState) {
            switch -Regex ($FormFactor) {
                "(?i)^M\.2$"           { return "M.2 SATA SSD" }
                "(?i)^mSATA$"          { return "mSATA SSD" }
                "(?i)^2\.5 inches$"   { return "2.5-inch SATA SSD" }
                "(?i)^1\.8 inches$"   { return "1.8-inch SATA SSD" }
                "(?i)^< 1\.8 inches$" { return "<1.8-inch SATA SSD" }
                default                 { return "SATA SSD" }
            }
        }

        if ($IsRotational) {
            switch -Regex ($FormFactor) {
                "(?i)^3\.5 inches$" { return "3.5-inch SATA HDD" }
                "(?i)^2\.5 inches$" { return "2.5-inch SATA HDD" }
                "(?i)^1\.8 inches$" { return "1.8-inch SATA HDD" }
                default               { return "SATA HDD" }
            }
        }

        # The interface is known, but the media type is not.
        return "SATA Drive"
    }

    # Some bridges expose the underlying media through a generic SCSI layer.
    # Classify the media only when smartctl exposes enough information to do so.
    if ($IsSolidState) {
        switch -Regex ($FormFactor) {
            "(?i)^M\.2$"           { return "M.2 SSD" }
            "(?i)^mSATA$"          { return "mSATA SSD" }
            "(?i)^2\.5 inches$"   { return "2.5-inch SSD" }
            "(?i)^1\.8 inches$"   { return "1.8-inch SSD" }
            "(?i)^< 1\.8 inches$" { return "<1.8-inch SSD" }
            default                 { return "SSD" }
        }
    }

    if ($IsRotational) {
        switch -Regex ($FormFactor) {
            "(?i)^3\.5 inches$" { return "3.5-inch HDD" }
            "(?i)^2\.5 inches$" { return "2.5-inch HDD" }
            "(?i)^1\.8 inches$" { return "1.8-inch HDD" }
            default               { return "HDD" }
        }
    }

    return "N/A"
}


function Get-DriveMake {

    param (
        [string]$Model
    )

    if ([string]::IsNullOrWhiteSpace($Model)) {
        return "N/A"
    }

    # Conservative model matching.
    # Unknown models intentionally return N/A rather than guessing.

    switch -Regex ($Model) {

        "(?i)^LENSE" {
            return "Lenovo"
        }

        "(?i)Samsung|^MZ[A-Z0-9]" {
            return "Samsung"
        }

        "(?i)SK[\s_-]*hynix|^HFS|^HFM" {
            return "SK hynix"
        }

        "(?i)KIOXIA|^KBG|^KXG" {
            return "KIOXIA"
        }

        "(?i)TOSHIBA|^THNS" {
            return "Toshiba"
        }

        "(?i)Western Digital|WDC|^WDS" {
            return "Western Digital"
        }

        "(?i)SanDisk" {
            return "SanDisk"
        }

        "(?i)Micron|^MTFD" {
            return "Micron"
        }

        "(?i)Crucial" {
            return "Crucial"
        }

        "(?i)KINGSTON" {
            return "Kingston"
        }

        "(?i)Intel|^SSDPE|^SSDSC" {
            return "Intel"
        }

        "(?i)Solidigm" {
            return "Solidigm"
        }

        "(?i)LITEON|LITE-ON" {
            return "Lite-On"
        }

        "(?i)ADATA" {
            return "ADATA"
        }

        "(?i)Seagate|^ST[0-9]" {
            return "Seagate"
        }

        "(?i)Hitachi|HGST" {
            return "HGST"
        }

        "(?i)PNY" {
            return "PNY"
        }

        "(?i)Transcend" {
            return "Transcend"
        }

        default {
            return "N/A"
        }
    }
}


function Get-TargetUsbDisks {

    return @(
        Get-Disk -ErrorAction SilentlyContinue |
        Where-Object {
            $_.BusType -eq "USB" -and
            -not $_.IsBoot -and
            -not $_.IsSystem -and
            $_.FriendlyName -like $UsbDevicePattern
        }
    )
}


function Test-LooksLikeBridgeIdentity {

    param (
        [string]$Model,
        [string]$Serial,
        [object]$Disk,
        [string]$Protocol
    )

    if ($Protocol -match "(?i)ATA|SATA|NVMe") {
        return $false
    }

    $NormalizedModel = (Get-CleanValue $Model).Trim()
    $NormalizedSerial = (Get-CleanValue $Serial).Trim()
    $WindowsModel = (Get-CleanValue $Disk.FriendlyName).Trim()
    $WindowsSerial = (Get-CleanValue $Disk.SerialNumber).Trim()

    $GenericModel = $NormalizedModel -match '(?i)^(SSK|USB|USB Device|External|Generic|Mass Storage|JMicron|ASMedia|Realtek|SCSI Disk Device)$'
    $SameModel = $NormalizedModel -ne "N/A" -and $NormalizedModel -eq $WindowsModel
    $SameSerial = $NormalizedSerial -ne "N/A" -and $NormalizedSerial -eq $WindowsSerial

    return ($GenericModel -or ($SameModel -and $SameSerial -and $WindowsModel -match '(?i)SSK|USB|External|Generic|JMicron|ASMedia|Realtek'))
}


function Get-DriveInformation {

    param (
        [Parameter(Mandatory = $true)]
        [object]$Disk
    )

    $DiskNumber = [int]$Disk.Number
    $Device = "/dev/pd$DiskNumber"
    $LastFailure = "No supported smartctl transport returned usable media identity."
    $FallbackIdentity = $null

    for ($Cycle = 1; $Cycle -le $SmartctlRetries; $Cycle++) {
        $Transports = Get-SmartctlTransportCandidates -DiskNumber $DiskNumber

        foreach ($Transport in $Transports) {
            $Arguments = [System.Collections.ArrayList]::new()
            [void]$Arguments.Add("-i")
            [void]$Arguments.Add("-j")

            if ($Transport -ne "auto") {
                [void]$Arguments.Add("-d")
                [void]$Arguments.Add($Transport)
            }

            [void]$Arguments.Add($Device)

            Write-Log -Level DEBUG -Message (
                "smartctl identity probe: cycle={0}/{1}, disk={2}, device={3}, transport={4}" -f
                $Cycle, $SmartctlRetries, $DiskNumber, $Device, $Transport
            )

            try {
                $Result = Invoke-SmartctlProcess `
                    -Arguments @($Arguments) `
                    -Context ("smartctl identity disk={0} transport={1}" -f $DiskNumber, $Transport)

                Write-Log -Level DEBUG -Message (
                    "smartctl completed: disk={0}, transport={1}, exitCode={2}, stdoutBytes={3}, stderrBytes={4}" -f
                    $DiskNumber,
                    $Transport,
                    $Result.ExitCode,
                    ([System.Text.Encoding]::UTF8.GetByteCount([string]$Result.StdOut)),
                    ([System.Text.Encoding]::UTF8.GetByteCount([string]$Result.StdErr))
                )

                if (-not [string]::IsNullOrWhiteSpace($Result.StdErr)) {
                    Write-Log -Level DEBUG -Message (
                        "smartctl stderr: disk={0}, transport={1}, text={2}" -f
                        $DiskNumber,
                        $Transport,
                        (($Result.StdErr.Trim()) -replace "[\r\n]+", " | ")
                    )
                }

                if ([string]::IsNullOrWhiteSpace($Result.StdOut)) {
                    $LastFailure = "smartctl returned no output for transport '$Transport'."
                    continue
                }

                try {
                    $Json = $Result.StdOut | ConvertFrom-Json
                }
                catch {
                    $LastFailure = "Unable to parse smartctl JSON for transport '$Transport'."
                    Write-ExceptionLog -ErrorRecord $_ -Context $LastFailure
                    continue
                }

                $Model = Get-CleanValue $Json.model_name
                $Serial = Get-CleanValue $Json.serial_number
                $CapacityBytes = $null

                if ($null -ne $Json.user_capacity) {
                    $CapacityBytes = $Json.user_capacity.bytes
                }

                $Capacity = Get-ReportedCapacity $CapacityBytes
                $Protocol = "N/A"
                $SmartctlType = "N/A"

                if ($null -ne $Json.device) {
                    $Protocol = Get-CleanValue $Json.device.protocol
                    $SmartctlType = Get-CleanValue $Json.device.type
                }

                $Type = Get-DriveType -SmartctlJson $Json -Model $Model
                $Make = Get-DriveMake $Model

                $HasUsefulIdentity = (
                    $Model -ne "N/A" -or
                    $Serial -ne "N/A"
                ) -and $Capacity -ne "N/A"

                if (-not $HasUsefulIdentity) {
                    $LastFailure = "Transport '$Transport' opened but did not expose usable model/serial/capacity data."
                    continue
                }

                $BridgeIdentity = Test-LooksLikeBridgeIdentity `
                    -Model $Model `
                    -Serial $Serial `
                    -Disk $Disk `
                    -Protocol $Protocol

                Write-Log -Level DEBUG -Message (
                    "identity candidate: disk={0}, transport={1}, smartctlType='{2}', protocol='{3}', make='{4}', model='{5}', serial='{6}', capacity='{7}', type='{8}', bridgeIdentity={9}" -f
                    $DiskNumber, $Transport, $SmartctlType, $Protocol, $Make, $Model, $Serial, $Capacity, $Type, $BridgeIdentity
                )

                $Candidate = [PSCustomObject]@{
                    Make             = $Make
                    Model            = $Model
                    SerialNumber     = $Serial
                    Capacity         = $Capacity
                    Type             = $Type
                    Protocol         = $Protocol
                    Transport        = $Transport
                    SmartctlType     = $SmartctlType
                    SmartctlExitCode = $Result.ExitCode
                }

                # ATA/SATA or NVMe is a high-confidence media identity. Prefer it
                # over a generic SCSI bridge presentation.
                if ($Protocol -match "(?i)ATA|SATA|NVMe") {
                    $script:PreferredTransportByDiskNumber[$DiskNumber] = $Transport

                    Write-Log -Level INFO -Message (
                        "Drive identity read: disk={0}, transport={1}, protocol='{2}', make='{3}', model='{4}', serial='{5}', capacity='{6}', type='{7}'" -f
                        $DiskNumber, $Transport, $Protocol, $Make, $Model, $Serial, $Capacity, $Type
                    )

                    return $Candidate
                }

                # Keep a non-generic SCSI identity as a final fallback so the
                # drive can still be inventoried with Type=N/A if passthrough is
                # unavailable but the bridge exposes the real model/serial.
                if (-not $BridgeIdentity -and $null -eq $FallbackIdentity) {
                    $FallbackIdentity = $Candidate
                }
            }
            catch {
                $LastFailure = $_.Exception.Message
                Write-ExceptionLog -ErrorRecord $_ -Context (
                    "smartctl probe failed: disk=$DiskNumber transport=$Transport"
                )
                if ($_.Exception -is [System.TimeoutException]) { throw }
            }
        }

        if ($null -ne $FallbackIdentity) {
            Write-Log -Level WARN -Message (
                "Using fallback non-ATA/NVMe identity for disk={0}; Type remains '{1}', transport={2}, protocol='{3}'" -f
                $DiskNumber, $FallbackIdentity.Type, $FallbackIdentity.Transport, $FallbackIdentity.Protocol
            )

            return $FallbackIdentity
        }

        if ($Cycle -lt $SmartctlRetries) {
            Start-Sleep -Milliseconds $RetryDelayMilliseconds
        }
    }

    throw "Unable to read media identity from $Device after testing supported smartctl USB transports. Last error: $LastFailure"
}


function ConvertTo-XmlText {

    param (
        [object]$Value
    )

    if ($null -eq $Value) {
        return ""
    }

    return [System.Security.SecurityElement]::Escape([string]$Value)
}


function Add-ZipTextEntry {

    param (
        [System.IO.Compression.ZipArchive]$Archive,
        [string]$EntryName,
        [string]$Content
    )

    $Entry = $Archive.CreateEntry(
        $EntryName,
        [System.IO.Compression.CompressionLevel]::Optimal
    )

    $Stream = $null
    $Writer = $null

    try {
        $Stream = $Entry.Open()
        $Encoding = [System.Text.UTF8Encoding]::new($false)
        $Writer = [System.IO.StreamWriter]::new($Stream, $Encoding)
        $Writer.Write($Content)
        $Writer.Flush()
    }
    finally {
        if ($null -ne $Writer) {
            $Writer.Dispose()
        }
        elseif ($null -ne $Stream) {
            $Stream.Dispose()
        }
    }
}


function Get-ZipEntryText {

    param (
        [System.IO.Compression.ZipArchiveEntry]$Entry
    )

    if ($null -eq $Entry) {
        return $null
    }

    $Stream = $null
    $Reader = $null

    try {
        $Stream = $Entry.Open()
        $Reader = [System.IO.StreamReader]::new($Stream)
        return $Reader.ReadToEnd()
    }
    finally {
        if ($null -ne $Reader) {
            $Reader.Dispose()
        }
        elseif ($null -ne $Stream) {
            $Stream.Dispose()
        }
    }
}


function Get-XlsxCellText {

    param (
        [System.Xml.XmlElement]$Cell,
        [System.Collections.ArrayList]$SharedStrings
    )

    $CellType = $Cell.GetAttribute("t")

    if ($CellType -eq "s") {
        $ValueNode = $Cell.SelectSingleNode("./*[local-name()='v']")

        if ($null -eq $ValueNode) {
            return ""
        }

        $Index = 0

        if (
            [int]::TryParse(
                $ValueNode.InnerText,
                [ref]$Index
            ) -and
            $Index -ge 0 -and
            $Index -lt $SharedStrings.Count
        ) {
            return [string]$SharedStrings[$Index]
        }

        return ""
    }

    if ($CellType -eq "inlineStr") {
        $TextNodes = $Cell.SelectNodes("./*[local-name()='is']//*[local-name()='t']")
        $Builder = [System.Text.StringBuilder]::new()

        foreach ($TextNode in $TextNodes) {
            [void]$Builder.Append($TextNode.InnerText)
        }

        return $Builder.ToString()
    }

    $ValueNode = $Cell.SelectSingleNode("./*[local-name()='v']")

    if ($null -ne $ValueNode) {
        return $ValueNode.InnerText
    }

    return ""
}


function Read-XlsxInventory {

    param (
        [string]$Path
    )

    $Records = [System.Collections.ArrayList]::new()

    if (-not (Test-Path -LiteralPath $Path)) {
        return $Records
    }

    $FileStream = $null
    $Archive = $null

    try {
        $FileStream = [System.IO.FileStream]::new(
            $Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite
        )

        $Archive = [System.IO.Compression.ZipArchive]::new(
            $FileStream,
            [System.IO.Compression.ZipArchiveMode]::Read,
            $false
        )

        $SharedStrings = [System.Collections.ArrayList]::new()
        $SharedStringsEntry = $Archive.GetEntry("xl/sharedStrings.xml")

        if ($null -ne $SharedStringsEntry) {
            $SharedStringsXmlText = Get-ZipEntryText -Entry $SharedStringsEntry

            if (-not [string]::IsNullOrWhiteSpace($SharedStringsXmlText)) {
                $SharedStringsXml = [System.Xml.XmlDocument]::new()
                $SharedStringsXml.LoadXml($SharedStringsXmlText)

                foreach ($StringItem in $SharedStringsXml.SelectNodes("//*[local-name()='si']")) {
                    $Builder = [System.Text.StringBuilder]::new()

                    foreach ($TextNode in $StringItem.SelectNodes(".//*[local-name()='t']")) {
                        [void]$Builder.Append($TextNode.InnerText)
                    }

                    [void]$SharedStrings.Add($Builder.ToString())
                }
            }
        }

        $SheetEntry = $Archive.GetEntry("xl/worksheets/sheet1.xml")

        if ($null -eq $SheetEntry) {
            throw "The existing workbook does not contain xl/worksheets/sheet1.xml."
        }

        $SheetXmlText = Get-ZipEntryText -Entry $SheetEntry
        $SheetXml = [System.Xml.XmlDocument]::new()
        $SheetXml.LoadXml($SheetXmlText)

        foreach ($RowNode in $SheetXml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")) {
            $RowNumber = 0
            [void][int]::TryParse($RowNode.GetAttribute("r"), [ref]$RowNumber)

            if ($RowNumber -le 1) {
                continue
            }

            $Values = @{
                A = ""
                B = ""
                C = ""
                D = ""
                E = ""
            }

            foreach ($Cell in $RowNode.SelectNodes("./*[local-name()='c']")) {
                $Reference = $Cell.GetAttribute("r")

                if ($Reference -match "^([A-E])\d+$") {
                    $Column = $Matches[1]
                    $Values[$Column] = Get-XlsxCellText -Cell $Cell -SharedStrings $SharedStrings
                }
            }

            if (
                [string]::IsNullOrWhiteSpace($Values.A) -and
                [string]::IsNullOrWhiteSpace($Values.B) -and
                [string]::IsNullOrWhiteSpace($Values.C) -and
                [string]::IsNullOrWhiteSpace($Values.D) -and
                [string]::IsNullOrWhiteSpace($Values.E)
            ) {
                continue
            }

            [void]$Records.Add(
                [PSCustomObject]@{
                    Make         = Get-CleanValue $Values.A
                    Model        = Get-CleanValue $Values.B
                    SerialNumber = Get-CleanValue $Values.C
                    Capacity     = Get-CleanValue $Values.D
                    Type         = Get-CleanValue $Values.E
                }
            )
        }
    }
    catch {
        throw "Unable to read existing workbook '$Path'. $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $Archive) {
            $Archive.Dispose()
        }

        if ($null -ne $FileStream) {
            $FileStream.Dispose()
        }
    }

    return $Records
}


function Write-XlsxInventory {

    param (
        [string]$Path,
        [System.Collections.IEnumerable]$Records
    )

    $TargetDirectory = Split-Path -Parent $Path
    $TargetName = [System.IO.Path]::GetFileName($Path)

    $SheetBuilder = [System.Text.StringBuilder]::new()

    [void]$SheetBuilder.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$SheetBuilder.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$SheetBuilder.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    [void]$SheetBuilder.Append('<cols>')
    [void]$SheetBuilder.Append('<col min="1" max="1" width="20" customWidth="1"/>')
    [void]$SheetBuilder.Append('<col min="2" max="2" width="35" customWidth="1"/>')
    [void]$SheetBuilder.Append('<col min="3" max="3" width="25" customWidth="1"/>')
    [void]$SheetBuilder.Append('<col min="4" max="4" width="20" customWidth="1"/>')
    [void]$SheetBuilder.Append('<col min="5" max="5" width="18" customWidth="1"/>')
    [void]$SheetBuilder.Append('</cols>')
    [void]$SheetBuilder.Append('<sheetData>')
    [void]$SheetBuilder.Append('<row r="1">')
    [void]$SheetBuilder.Append('<c r="A1" t="inlineStr" s="1"><is><t>Make</t></is></c>')
    [void]$SheetBuilder.Append('<c r="B1" t="inlineStr" s="1"><is><t>Model</t></is></c>')
    [void]$SheetBuilder.Append('<c r="C1" t="inlineStr" s="1"><is><t>Serial Number</t></is></c>')
    [void]$SheetBuilder.Append('<c r="D1" t="inlineStr" s="1"><is><t>Reported Capacity</t></is></c>')
    [void]$SheetBuilder.Append('<c r="E1" t="inlineStr" s="1"><is><t>Type</t></is></c>')
    [void]$SheetBuilder.Append('</row>')

    $RowNumber = 2

    foreach ($Record in $Records) {
        $Make = ConvertTo-XmlText (Get-CleanValue $Record.Make)
        $Model = ConvertTo-XmlText (Get-CleanValue $Record.Model)
        $Serial = ConvertTo-XmlText (Get-CleanValue $Record.SerialNumber)
        $Capacity = ConvertTo-XmlText (Get-CleanValue $Record.Capacity)
        $Type = ConvertTo-XmlText (Get-CleanValue $Record.Type)

        [void]$SheetBuilder.Append(('<row r="{0}">' -f $RowNumber))
        [void]$SheetBuilder.Append(('<c r="A{0}" t="inlineStr"><is><t>{1}</t></is></c>' -f $RowNumber, $Make))
        [void]$SheetBuilder.Append(('<c r="B{0}" t="inlineStr"><is><t>{1}</t></is></c>' -f $RowNumber, $Model))
        [void]$SheetBuilder.Append(('<c r="C{0}" t="inlineStr"><is><t>{1}</t></is></c>' -f $RowNumber, $Serial))
        [void]$SheetBuilder.Append(('<c r="D{0}" t="inlineStr"><is><t>{1}</t></is></c>' -f $RowNumber, $Capacity))
        [void]$SheetBuilder.Append(('<c r="E{0}" t="inlineStr"><is><t>{1}</t></is></c>' -f $RowNumber, $Type))
        [void]$SheetBuilder.Append('</row>')

        $RowNumber++
    }

    [void]$SheetBuilder.Append('</sheetData>')

    if ($RowNumber -gt 2) {
        $LastRow = $RowNumber - 1
        [void]$SheetBuilder.Append(('<autoFilter ref="A1:E{0}"/>' -f $LastRow))
    }

    [void]$SheetBuilder.Append('</worksheet>')

    $ContentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
'@

    $RootRelationships = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@

    $WorkbookXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Inventory" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
'@

    $WorkbookRelationships = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'@

    $StylesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
'@

    $LastSaveError = $null

    for ($SaveAttempt = 1; $SaveAttempt -le $WorkbookSaveRetries; $SaveAttempt++) {
        $TempPath = Join-Path -Path $TargetDirectory -ChildPath (
            ".{0}.{1}.{2}.tmp" -f $TargetName, $PID, ([guid]::NewGuid().ToString("N"))
        )

        $FileStream = $null
        $Archive = $null

        try {
            Write-Log -Level DEBUG -Message (
                "Workbook save attempt {0}/{1}: records={2}, target='{3}', temp='{4}'" -f
                $SaveAttempt, $WorkbookSaveRetries, @($Records).Count, $Path, $TempPath
            )

            $FileStream = [System.IO.FileStream]::new(
                $TempPath,
                [System.IO.FileMode]::CreateNew,
                [System.IO.FileAccess]::ReadWrite,
                [System.IO.FileShare]::None
            )

            # Keep the backing FileStream open when the ZipArchive is disposed.
            # Disposing a ZipArchive in Create mode finalizes the XLSX central
            # directory. With leaveOpen=$false, that same Dispose() also closes
            # $FileStream, making the subsequent Flush() deterministically fail
            # with ObjectDisposedException ("Cannot access a closed file").
            $Archive = [System.IO.Compression.ZipArchive]::new(
                $FileStream,
                [System.IO.Compression.ZipArchiveMode]::Create,
                $true
            )

            Add-ZipTextEntry -Archive $Archive -EntryName "[Content_Types].xml" -Content $ContentTypes.Trim()
            Add-ZipTextEntry -Archive $Archive -EntryName "_rels/.rels" -Content $RootRelationships.Trim()
            Add-ZipTextEntry -Archive $Archive -EntryName "xl/workbook.xml" -Content $WorkbookXml.Trim()
            Add-ZipTextEntry -Archive $Archive -EntryName "xl/_rels/workbook.xml.rels" -Content $WorkbookRelationships.Trim()
            Add-ZipTextEntry -Archive $Archive -EntryName "xl/styles.xml" -Content $StylesXml.Trim()
            Add-ZipTextEntry -Archive $Archive -EntryName "xl/worksheets/sheet1.xml" -Content $SheetBuilder.ToString()

            Write-Log -Level DEBUG -Message "Finalizing temporary XLSX ZIP archive."
            $Archive.Dispose()
            $Archive = $null

            Write-Log -Level DEBUG -Message "Flushing finalized temporary XLSX file stream."
            $FileStream.Flush()
            $FileStream.Dispose()
            $FileStream = $null

            # Validate that the temporary package contains the worksheet before
            # replacing the last-known-good workbook.
            $ValidationStream = $null
            $ValidationArchive = $null

            try {
                $ValidationStream = [System.IO.FileStream]::new(
                    $TempPath,
                    [System.IO.FileMode]::Open,
                    [System.IO.FileAccess]::Read,
                    [System.IO.FileShare]::Read
                )

                $ValidationArchive = [System.IO.Compression.ZipArchive]::new(
                    $ValidationStream,
                    [System.IO.Compression.ZipArchiveMode]::Read,
                    $false
                )

                if (
                    $null -eq $ValidationArchive.GetEntry("xl/workbook.xml") -or
                    $null -eq $ValidationArchive.GetEntry("xl/worksheets/sheet1.xml")
                ) {
                    throw "Temporary XLSX package validation failed: required workbook entries are missing."
                }
            }
            finally {
                if ($null -ne $ValidationArchive) {
                    $ValidationArchive.Dispose()
                }
                if ($null -ne $ValidationStream) {
                    $ValidationStream.Dispose()
                }
            }

            if (Test-Path -LiteralPath $Path) {
                try {
                    [System.IO.File]::Replace($TempPath, $Path, $null)
                }
                catch {
                    Write-Log -Level WARN -Message (
                        "Atomic File.Replace failed; attempting overwrite copy. Error: {0}" -f $_.Exception.Message
                    )

                    [System.IO.File]::Copy($TempPath, $Path, $true)
                    Remove-Item -LiteralPath $TempPath -Force -ErrorAction SilentlyContinue
                }
            }
            else {
                [System.IO.File]::Move($TempPath, $Path)
            }

            $SavedFile = Get-Item -LiteralPath $Path -ErrorAction Stop
            Write-Log -Level INFO -Message (
                "Workbook saved successfully: records={0}, bytes={1}, attempt={2}, path='{3}'" -f
                @($Records).Count, $SavedFile.Length, $SaveAttempt, $Path
            )

            return
        }
        catch {
            $LastSaveError = $_
            Write-ExceptionLog -ErrorRecord $_ -Context (
                "Workbook save attempt $SaveAttempt/$WorkbookSaveRetries failed"
            )

            if (Test-Path -LiteralPath $Path) {
                try {
                    $Existing = Get-Item -LiteralPath $Path -ErrorAction Stop
                    Write-Log -Level DEBUG -Message (
                        "Existing workbook state: bytes={0}, attributes={1}, lastWrite='{2:o}'" -f
                        $Existing.Length, $Existing.Attributes, $Existing.LastWriteTime
                    )
                }
                catch {}
            }

            if ($SaveAttempt -lt $WorkbookSaveRetries) {
                Start-Sleep -Milliseconds ($WorkbookRetryDelayMilliseconds * $SaveAttempt)
            }
        }
        finally {
            if ($null -ne $Archive) {
                $Archive.Dispose()
            }

            if ($null -ne $FileStream) {
                $FileStream.Dispose()
            }

            if (Test-Path -LiteralPath $TempPath) {
                Remove-Item -LiteralPath $TempPath -Force -ErrorAction SilentlyContinue
            }
        }
    }

    $LastMessage = if ($null -ne $LastSaveError) {
        $LastSaveError.Exception.Message
    }
    else {
        "Unknown workbook save failure."
    }

    throw "Unable to save workbook '$Path' after $WorkbookSaveRetries attempts. Last error: $LastMessage"
}


# ------------------------------------------------------------
# Prerequisite checks
# ------------------------------------------------------------

$Identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$IsAdmin = Test-IsAdministrator

if (-not $IsAdmin) {
    Write-Host ""
    Write-Host "ERROR: Administrator privileges are required."
    Write-Host "Open PowerShell as Administrator and run this script again."
    Write-Host ""
    Write-Log -Level ERROR -Message "Prerequisite failure: process is not elevated."
    exit 1
}

if ($PSVersionTable.PSVersion -lt [version]"5.1") {
    Write-Host ""
    Write-Host "ERROR: Windows PowerShell 5.1 or later is required."
    Write-Host ""
    Write-Log -Level ERROR -Message ("Prerequisite failure: unsupported PowerShell version {0}." -f $PSVersionTable.PSVersion)
    exit 1
}

if ($null -eq (Get-Command Get-Disk -ErrorAction SilentlyContinue)) {
    Write-Host ""
    Write-Host "ERROR: The Windows Storage module/Get-Disk command is unavailable."
    Write-Host "This collector requires a supported Windows installation with the Storage module."
    Write-Host ""
    Write-Log -Level ERROR -Message "Prerequisite failure: Get-Disk command is unavailable."
    exit 1
}

try {
    $script:SmartctlPath = Ensure-SmartctlDependency
}
catch {
    Write-Host ""
    Write-Host "ERROR: Dependency check failed."
    Write-Host $_.Exception.Message
    Write-Host ""
    Write-Log -Level ERROR -Message ("Dependency failure: " + $_.Exception.Message)
    exit 1
}

$InteractiveUser = "N/A"
try {
    $InteractiveUser = (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).UserName
}
catch {}

$SmartctlVersion = "N/A"
try {
    $VersionResult = Invoke-SmartctlProcess -Arguments @("--version") -Context "smartctl version check"
    if (-not [string]::IsNullOrWhiteSpace($VersionResult.StdOut)) {
        $SmartctlVersion = (($VersionResult.StdOut -split "[\r\n]+" | Select-Object -First 1).Trim())
    }
}
catch {
    Write-ExceptionLog -ErrorRecord $_ -Context "Unable to read smartctl version"
}

Write-Log -Level INFO -Message "Collector startup: version=$ScriptVersion, script='$PSCommandPath'"
Write-Log -Level INFO -Message "Process identity='$($Identity.Name)', interactive user='$InteractiveUser', elevated=$IsAdmin"
Write-Log -Level INFO -Message "PowerShell=$($PSVersionTable.PSVersion), OS=$([Environment]::OSVersion.VersionString)"
Write-Log -Level INFO -Message "smartctl='$script:SmartctlPath', smartctlVersion='$SmartctlVersion', output='$OutputPath', log='$LogPath'"
Write-Log -Level INFO -Message (
    "Settings: usbDevicePattern='{0}', pollSeconds={1}, smartctlRetries={2}, settleMs={3}, retryMs={4}, workbookRetries={5}, workbookRetryMs={6}" -f
    $UsbDevicePattern, $PollSeconds, $SmartctlRetries, $InitialSettleMilliseconds,
    $RetryDelayMilliseconds, $WorkbookSaveRetries, $WorkbookRetryDelayMilliseconds
)

if ($OutputPath -match '(?i)\\OneDrive(?:\s-\s[^\\]+)?\\') {
    Write-Log -Level INFO -Message "Output path is OneDrive-backed; workbook save retry protection is enabled."
}

# ------------------------------------------------------------
# Load or create inventory
# ------------------------------------------------------------

$Inventory = [System.Collections.ArrayList]::new()
$KnownSerials = @{}

if (Test-Path -LiteralPath $OutputPath) {
    Write-Host "Opening existing inventory:"
    Write-Host "  $OutputPath"
    Write-Host ""

    foreach ($Record in @(Read-XlsxInventory -Path $OutputPath)) {
        [void]$Inventory.Add($Record)

        if (
            $Record.SerialNumber -ne "N/A" -and
            -not [string]::IsNullOrWhiteSpace($Record.SerialNumber)
        ) {
            $KnownSerials[$Record.SerialNumber] = $true
        }
    }

    Write-Log -Level INFO -Message (
        "Existing inventory loaded: records={0}, source='{1}'" -f $Inventory.Count, $OutputPath
    )
}
else {
    Write-Host "Creating inventory:"
    Write-Host "  $OutputPath"
    Write-Host ""

    Write-XlsxInventory -Path $OutputPath -Records $Inventory
    Write-Log -Level INFO -Message "New inventory workbook created."
}


Write-Host "USB Drive Inventory Collector v$ScriptVersion"
Write-Host "---------------------------------------"
Write-Host ""
Write-Host "Scope:     USB physical drives (boot/system disks excluded)"
Write-Host "Adapters:  smartctl autodetection + NVMe/SATA USB transport fallbacks"
Write-Host "Output:    $OutputPath"
Write-Host "Log:       $LogPath"
Write-Host "Backend:   Direct XLSX (no Excel COM)"
Write-Host "smartctl:  $SmartctlVersion"
Write-Host ""
Write-Host "Insert one drive at a time."
Write-Host "The workbook is saved after every drive."
Write-Host "Do not keep the workbook open in Excel while collecting drives."
Write-Host "Press M to add a manual drive record while waiting."
Write-Host "Press Ctrl+C when finished."
Write-Host ""

# Tracks only USB disks that are currently attached. A removed disk disappears
# from this table; a later insertion on the same Windows disk number is new.
$ConnectedDisks = @{}

# Windows Settings stores the current user's AutoPlay switch here. This is a
# user preference, not the machine/user AutoPlay policy managed by Group Policy.
$AutoPlayRegistryPath = 'Software\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers'
$AutoPlayRestore = $null

function Set-TemporaryAutoPlayPreference {
    $DesktopSid = $null
    if ($InteractiveUser -ne 'N/A') {
        try {
            $DesktopSid = ([Security.Principal.NTAccount]::new($InteractiveUser)).Translate(
                [Security.Principal.SecurityIdentifier]
            ).Value
        }
        catch {
            Write-ExceptionLog -ErrorRecord $_ -Context "Could not resolve desktop user '$InteractiveUser'"
        }
    }

    if ($null -eq $DesktopSid -or $DesktopSid -ne $Identity.User.Value) {
        Write-Host 'AutoPlay prompt skipped: the signed-in desktop user could not be matched to this account.'
        Write-Log -Level WARN -Message "AutoPlay preference skipped: process='$($Identity.Name)' SID='$($Identity.User.Value)', desktop='$InteractiveUser' SID='$DesktopSid'."
        return
    }

    $Key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($AutoPlayRegistryPath, $false)
    $KeyWasPresent = $null -ne $Key

    try {
        $HadValue = $KeyWasPresent -and (@($Key.GetValueNames()) -contains 'DisableAutoplay')
        if ($HadValue) {
            $OriginalKind = $Key.GetValueKind('DisableAutoplay')
            if ($OriginalKind -ne [Microsoft.Win32.RegistryValueKind]::DWord) {
                Write-Host 'AutoPlay uses an unexpected setting type. The Windows setting was not changed.'
                Write-Log -Level WARN -Message "AutoPlay preference has unexpected type: $OriginalKind"
                return
            }
            $OriginalValue = [int]$Key.GetValue('DisableAutoplay')
            if ($OriginalValue -eq 1) {
                Write-Log -Level INFO -Message 'AutoPlay is already disabled for the signed-in user.'
                return
            }
        }
        else {
            $OriginalValue = $null
        }

        do {
            $Choice = (Read-Host 'AutoPlay can open drive folders or show pop-ups. Disable it while collecting? [Y/N]').Trim()
        } while ($Choice -notin @('Y', 'N'))

        if ($Choice -eq 'N') {
            Write-Log -Level INFO -Message 'User declined temporary AutoPlay change.'
            return
        }

        # Keep the original presence and value, including an explicit zero.
        # Register for cleanup before writing in case the write partly succeeds.
        $WriteKey = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($AutoPlayRegistryPath)
        if ($null -eq $WriteKey) { throw 'Could not open the AutoPlay preference for writing.' }
        try {
            $script:AutoPlayRestore = [PSCustomObject]@{
                HadValue      = $HadValue
                Value         = $OriginalValue
                KeyWasPresent = $KeyWasPresent
            }
            $WriteKey.SetValue('DisableAutoplay', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
        }
        finally {
            $WriteKey.Dispose()
        }
        Write-Host 'AutoPlay disabled for this Windows user until the collector exits.'
        Write-Log -Level INFO -Message 'Temporarily disabled AutoPlay for the signed-in user.'
    }
    finally {
        if ($null -ne $Key) { $Key.Dispose() }
    }
}

function Restore-AutoPlayPreference {
    if ($null -eq $script:AutoPlayRestore) { return }
    $RestoreState = $script:AutoPlayRestore

    $Key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($AutoPlayRegistryPath, $true)
    if ($null -eq $Key) { throw "AutoPlay registry key disappeared: HKCU\$AutoPlayRegistryPath" }

    try {
        # Leave a setting changed by the user or a policy agent during the run.
        if ((@($Key.GetValueNames()) -notcontains 'DisableAutoplay') -or
            $Key.GetValueKind('DisableAutoplay') -ne [Microsoft.Win32.RegistryValueKind]::DWord -or
            [int]$Key.GetValue('DisableAutoplay') -ne 1) {
            Write-Host 'AutoPlay changed during this run; leaving its current value in place.'
            Write-Log -Level WARN -Message 'AutoPlay preference changed externally; original value was not restored.'
            return
        }

        if ($RestoreState.HadValue) {
            $Key.SetValue('DisableAutoplay', $RestoreState.Value, [Microsoft.Win32.RegistryValueKind]::DWord)
        }
        else {
            $Key.DeleteValue('DisableAutoplay', $false)
        }
        Write-Host 'AutoPlay preference restored.'
        Write-Log -Level INFO -Message 'Original AutoPlay preference restored.'
        $script:AutoPlayRestore = $null
    }
    finally {
        $Key.Dispose()
    }

    if (-not $RestoreState.KeyWasPresent) {
        $ParentPath = 'Software\Microsoft\Windows\CurrentVersion\Explorer'
        $ParentKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($ParentPath, $true)
        if ($null -ne $ParentKey) {
            try {
                $ChildKey = $ParentKey.OpenSubKey('AutoplayHandlers', $false)
                if ($null -ne $ChildKey) {
                    try {
                        if ($ChildKey.ValueCount -eq 0 -and $ChildKey.SubKeyCount -eq 0) {
                            $ChildKey.Dispose()
                            $ChildKey = $null
                            $ParentKey.DeleteSubKey('AutoplayHandlers', $false)
                        }
                    }
                    finally {
                        if ($null -ne $ChildKey) { $ChildKey.Dispose() }
                    }
                }
            }
            finally {
                $ParentKey.Dispose()
            }
        }
    }
}

function Read-ManualText {
    param ([string]$Label, [string]$Pattern, [int]$MaxLength, [switch]$Uppercase, [switch]$AllowBack)

    while ($true) {
        $Hint = if ($AllowBack) { 'Enter for N/A, :back to review, or :cancel' } else { 'Enter for N/A, or :cancel' }
        $Answer = Read-Host "$Label ($Hint)"
        if ($null -eq $Answer -or $Answer.Trim() -eq ':cancel') { return $null }
        if ($AllowBack -and $Answer.Trim() -eq ':back') { return ':back' }
        if ([string]::IsNullOrWhiteSpace($Answer)) { return 'N/A' }
        $Answer = $Answer.Trim()
        if ($Answer.Length -gt $MaxLength -or $Answer -cnotmatch $Pattern) {
            Write-Host "Invalid $Label. Use 1-$MaxLength plain letters, digits, and standard punctuation."
            continue
        }
        if ($Uppercase) { return $Answer.ToUpperInvariant() }
        return $Answer
    }
}

function Read-ManualSelection {
    param ([string]$Label, [string[]]$Options, [switch]$AllowBack)

    Write-Host "${Label}:"
    for ($Index = 0; $Index -lt $Options.Count; $Index++) {
        Write-Host ("  {0}. {1}" -f ($Index + 1), $Options[$Index])
    }
    while ($true) {
        $Hint = if ($AllowBack) { 'Enter for N/A, :back to review, or :cancel' } else { 'Enter for N/A, or :cancel' }
        $Answer = Read-Host "Choose a number ($Hint)"
        if ($null -eq $Answer -or $Answer.Trim() -eq ':cancel') { return $null }
        if ($AllowBack -and $Answer.Trim() -eq ':back') { return ':back' }
        if ([string]::IsNullOrWhiteSpace($Answer)) { return 'N/A' }
        $Number = 0
        if ([int]::TryParse($Answer.Trim(), [ref]$Number) -and $Number -ge 1 -and $Number -le $Options.Count) {
            return $Options[$Number - 1]
        }
        Write-Host "Enter a number from 1 to $($Options.Count)."
    }
}

function Read-ManualDriveType {
    param ([switch]$AllowBack)

    $Groups = [ordered]@{
        Standard = @(
            '1.8-inch SATA SSD', '2.5-inch SATA HDD', '2.5-inch SATA SSD',
            '3.5-inch SATA HDD', 'M.2 NVMe SSD', 'M.2 SATA SSD',
            'mSATA SSD', 'NVMe SSD', 'SATA Drive', 'SATA HDD', 'SATA SSD'
        )
        Enterprise = @(
            '2.5-inch SAS HDD', '2.5-inch SAS SSD', '3.5-inch SAS HDD',
            '3.5-inch SAS SSD', 'SAS HDD', 'SAS SSD'
        )
        Other = @(
            '3.5-inch Floppy Disk', '5.25-inch Floppy Disk',
            'CompactFlash Card', 'eMMC', 'HDD', 'microSD Card',
            'SD Card', 'SSD', 'USB Flash Drive', 'Other'
        )
    }

    $Options = @()
    Write-Host 'Drive type:'
    foreach ($Group in $Groups.Keys) {
        Write-Host "  ${Group}:"
        foreach ($Option in $Groups[$Group]) {
            $Options += $Option
            Write-Host ('    {0}. {1}' -f $Options.Count, $Option)
        }
    }
    while ($true) {
        $Hint = if ($AllowBack) { 'Enter for N/A, :back to review, or :cancel' } else { 'Enter for N/A, or :cancel' }
        $Answer = Read-Host "Choose a number ($Hint)"
        if ($null -eq $Answer -or $Answer.Trim() -eq ':cancel') { return $null }
        if ($AllowBack -and $Answer.Trim() -eq ':back') { return ':back' }
        if ([string]::IsNullOrWhiteSpace($Answer)) { return 'N/A' }
        $Number = 0
        if ([int]::TryParse($Answer.Trim(), [ref]$Number) -and $Number -ge 1 -and $Number -le $Options.Count) {
            return $Options[$Number - 1]
        }
        Write-Host "Enter a number from 1 to $($Options.Count)."
    }
}

function Read-ManualCapacity {
    param ([switch]$AllowBack)

    while ($true) {
        $Hint = if ($AllowBack) { 'Enter for N/A, :back to review, or :cancel' } else { 'Enter for N/A, or :cancel' }
        $Amount = Read-Host "Capacity (number only; $Hint)"
        if ($null -eq $Amount -or $Amount.Trim() -eq ':cancel') { return $null }
        if ($AllowBack -and $Amount.Trim() -eq ':back') { return ':back' }
        if ([string]::IsNullOrWhiteSpace($Amount)) { return 'N/A' }
        $Amount = $Amount.Trim()
        if ($Amount.Length -gt 22 -or $Amount -cnotmatch '\A[0-9]{1,15}(?:\.[0-9]{1,6})?\z') {
            Write-Host 'Enter a positive number without a unit, such as 1 or 0.005.'
            continue
        }
        $Number = [decimal]::Parse($Amount, [Globalization.CultureInfo]::InvariantCulture)
        if ($Number -gt 0) { break }
        Write-Host 'Capacity must be greater than zero.'
    }

    Clear-Host
    if ($AllowBack) { Write-Host 'EDIT CAPACITY - UNIT' }
    else { Write-Host 'Step 4 of 5 - Capacity unit' }
    Write-Host "Capacity amount: $Amount"
    Write-Host ''
    while ($true) {
        $Unit = Read-ManualSelection -Label 'Capacity unit' -Options @('B', 'KB', 'MB', 'GB', 'TB', 'PB', 'Other') -AllowBack:$AllowBack
        if ($null -eq $Unit) { return $null }
        if ($Unit -eq ':back') { return ':back' }
        if ($Unit -eq 'N/A') { return 'N/A' }
        if ($Unit -eq 'Other') {
            Clear-Host
            if ($AllowBack) { Write-Host 'EDIT CAPACITY - CUSTOM UNIT' }
            else { Write-Host 'Step 4 of 5 - Custom capacity unit' }
            Write-Host "Capacity amount: $Amount"
            Write-Host ''
            $Unit = Read-ManualText -Label 'Custom capacity unit (letters only)' `
                -Pattern '\A[A-Za-z]{1,12}\z' -MaxLength 12 -AllowBack:$AllowBack
            if ($null -eq $Unit) { return $null }
            if ($Unit -eq ':back') { return ':back' }
            if ($Unit -eq 'N/A') { return 'N/A' }
        }
        if ($Unit -ne 'B' -or $Number -eq [decimal]::Truncate($Number)) { break }
        Write-Host 'Bytes must be a whole number. Choose another unit or edit the amount later.'
    }

    return ('{0} {1}' -f $Number.ToString('0.######', [Globalization.CultureInfo]::InvariantCulture), $Unit)
}

function Read-ManualField {
    param ([ValidateRange(1, 5)][int]$Field, [switch]$AllowBack)

    switch ($Field) {
        1 { return (Read-ManualText -Label 'Make' -Pattern "\A[A-Za-z0-9][A-Za-z0-9 .&()+'/_-]{0,79}\z" -MaxLength 80 -AllowBack:$AllowBack) }
        2 { return (Read-ManualText -Label 'Model' -Pattern '\A[A-Za-z0-9][A-Za-z0-9 .+/_-]{0,99}\z' -MaxLength 100 -Uppercase -AllowBack:$AllowBack) }
        3 { return (Read-ManualText -Label 'Serial number' -Pattern '\A[A-Za-z0-9][A-Za-z0-9./_-]{0,99}\z' -MaxLength 100 -Uppercase -AllowBack:$AllowBack) }
        4 { return (Read-ManualCapacity -AllowBack:$AllowBack) }
        5 {
            $Type = Read-ManualDriveType -AllowBack:$AllowBack
            if ($null -eq $Type) { return $null }
            if ($Type -eq ':back') { return ':back' }
            if ($Type -eq 'Other') {
                Clear-Host
                if ($AllowBack) { Write-Host 'EDIT DRIVE TYPE - CUSTOM' }
                else { Write-Host 'Step 5 of 5 - Custom drive type' }
                return (Read-ManualText -Label 'Custom drive type' `
                    -Pattern '\A[A-Za-z0-9][A-Za-z0-9 .()+/_-]{0,59}\z' -MaxLength 60 -AllowBack:$AllowBack)
            }
            return $Type
        }
    }
}

function Show-ManualRecord {
    param ([object]$Record)

    Write-Host ''
    Write-Host 'REVIEW MANUAL DRIVE'
    Write-Host '-------------------'
    Write-Host "1. Make:     $($Record.Make)"
    Write-Host "2. Model:    $($Record.Model)"
    Write-Host "3. Serial:   $($Record.SerialNumber)"
    Write-Host "4. Capacity: $($Record.Capacity)"
    Write-Host "5. Type:     $($Record.Type)"
    Write-Host ''
}

function Invoke-ManualEntry {
    Write-Log -Level INFO -Message 'Manual drive recording opened.'
    $NextMode = $null

    while ($true) {
        Clear-Host
        Write-Host 'Manual drive recording initialized...'
        Write-Host 'Press Enter to record N/A for a field, or type :cancel to return to automatic recording.'
        Write-Host ''
        if ($null -eq $NextMode -and $Inventory.Count -gt 0) {
            while ($true) {
                $NextMode = Read-Host 'New drive [N], copy last saved drive [C], or return [R]'
                if ($null -eq $NextMode) { return }
                $NextMode = $NextMode.Trim()
                if ($NextMode -eq 'R') { return }
                if ($NextMode -eq 'N' -or $NextMode -eq 'C') { break }
                Write-Host 'Choose N, C, or R.'
            }
        }
        if ($null -eq $NextMode) { $NextMode = 'N' }

        if ($NextMode -eq 'C') {
            $Source = $Inventory[$Inventory.Count - 1]
            $Record = [PSCustomObject]@{
                Make = $Source.Make; Model = $Source.Model; SerialNumber = $null
                Capacity = $Source.Capacity; Type = $Source.Type
            }
        }
        else {
            $Record = [PSCustomObject]@{
                Make = $null; Model = $null; SerialNumber = $null; Capacity = $null; Type = $null
            }
        }
        $Fields = @('Make', 'Model', 'SerialNumber', 'Capacity', 'Type')
        $FieldLabels = @('Make', 'Model', 'Serial number', 'Capacity', 'Drive type')
        $FirstField = if ($NextMode -eq 'C') { 3 } else { 1 }
        $LastField = if ($NextMode -eq 'C') { 3 } else { 5 }
        for ($Index = $FirstField; $Index -le $LastField; $Index++) {
            Clear-Host
            if ($NextMode -eq 'C') {
                Write-Host 'Step 1 of 1 - Serial number'
                Write-Host 'Copying the last saved drive. Enter a new serial number.'
            }
            else {
                Write-Host "Step $Index of 5 - $($FieldLabels[$Index - 1])"
            }
            Write-Host ''
            $Value = Read-ManualField -Field $Index
            if ($null -eq $Value) {
                Write-Log -Level INFO -Message 'Manual entry cancelled before saving.'
                return
            }
            $Record.($Fields[$Index - 1]) = $Value
        }

        Clear-Host
        while ($true) {
            Show-ManualRecord -Record $Record
            $Duplicate = $Record.SerialNumber -ne 'N/A' -and $KnownSerials.ContainsKey($Record.SerialNumber)
            if ($Duplicate) {
                Write-Host 'This serial number is already in the workbook. Edit it or cancel.'
            }
            elseif ($Record.SerialNumber -eq 'N/A') {
                Write-Host 'Serial N/A cannot be checked for duplicates.'
            }

            $Action = Read-Host 'Save [Y], edit a field [E], or cancel [C]'
            if ($null -eq $Action) { return }
            $Action = $Action.Trim()
            if ($Action -eq 'C') {
                Write-Log -Level INFO -Message 'Manual entry cancelled at review.'
                return
            }
            if ($Action -eq 'E') {
                $FieldNumber = 0
                $Choice = Read-Host 'Field to edit [1-5], :back to review, or :cancel to leave'
                if ($null -eq $Choice -or $Choice.Trim() -eq ':cancel') { return }
                if ($Choice.Trim() -eq ':back') {
                    Clear-Host
                    continue
                }
                if (-not [int]::TryParse($Choice.Trim(), [ref]$FieldNumber) -or $FieldNumber -lt 1 -or $FieldNumber -gt 5) {
                    Write-Host 'Choose a field number from 1 to 5.'
                    continue
                }
                Clear-Host
                Write-Host "EDIT FIELD $FieldNumber (type :back to keep the current value)"
                Write-Host ''
                $Value = Read-ManualField -Field $FieldNumber -AllowBack
                if ($null -eq $Value) { return }
                if ($Value -eq ':back') {
                    Clear-Host
                    continue
                }
                $Record.($Fields[$FieldNumber - 1]) = $Value
                Clear-Host
                continue
            }
            if ($Action -ne 'Y') {
                Write-Host 'Choose Y, E, or C.'
                continue
            }
            if ($Duplicate) { continue }

            [void]$Inventory.Add($Record)
            try {
                Write-XlsxInventory -Path $OutputPath -Records $Inventory
            }
            catch {
                $Inventory.RemoveAt($Inventory.Count - 1)
                Write-ExceptionLog -ErrorRecord $_ -Context 'Manual drive workbook save failed'
                Write-Host "Save failed: $($_.Exception.Message)"
                Write-Host "Check the workbook, then try saving again. Debug log: $LogPath"
                continue
            }

            if ($Record.SerialNumber -ne 'N/A') { $KnownSerials[$Record.SerialNumber] = $true }
            Clear-Host
            Write-Host 'MANUAL DRIVE RECORDED'
            Write-Host '---------------------'
            Write-Host "Make:     $($Record.Make)"
            Write-Host "Model:    $($Record.Model)"
            Write-Host "Serial:   $($Record.SerialNumber)"
            Write-Host "Capacity: $($Record.Capacity)"
            Write-Host "Type:     $($Record.Type)"
            Write-Host ''
            Write-Host "Saved as row $($Inventory.Count + 1)."
            Write-Host ''
            Write-Log -Level INFO -Message (
                "Manual drive recorded: row={0}, make='{1}', model='{2}', serial='{3}', capacity='{4}', type='{5}'" -f
                ($Inventory.Count + 1), $Record.Make, $Record.Model, $Record.SerialNumber, $Record.Capacity, $Record.Type
            )
            break
        }

        while ($true) {
            $Next = Read-Host 'Add another manual drive [A], copy this drive with a new serial [C], or return [R]'
            if ($null -eq $Next) { return }
            $Next = $Next.Trim()
            if ($Next -eq 'A') { $NextMode = 'N'; break }
            if ($Next -eq 'C') { $NextMode = 'C'; break }
            if ($Next -eq 'R') { return }
            Write-Host 'Choose A, C, or R.'
        }
    }
}

function Test-ManualEntryHotkey {
    if ($script:ManualHotkeyUnavailable) { return $false }
    try {
        while ([Console]::KeyAvailable) {
            $Key = [Console]::ReadKey($true)
            if ($Key.Key -eq [ConsoleKey]::M -and $Key.Modifiers -eq 0) { return $true }
        }
    }
    catch {
        $script:ManualHotkeyUnavailable = $true
        if (-not $script:ManualHotkeyWarningShown) {
            $script:ManualHotkeyWarningShown = $true
            Write-Log -Level WARN -Message 'This PowerShell host does not support the M hotkey. Use -ManualEntryOnStartup.'
            Write-Host 'Manual hotkey unavailable in this host. Restart with -ManualEntryOnStartup if needed.'
        }
    }
    return $false
}

# Suppress critical device/read error dialogs generated by this PowerShell
# process. Explorer's AutoPlay behavior is managed separately by the temporary
# preference above; this process setting does not affect Explorer's dialogs.
$PreviousErrorMode = $null
try {
    if (-not ("DriveInventoryNativeErrorMode" -as [type])) {
        Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
public static class DriveInventoryNativeErrorMode {
    [DllImport("kernel32.dll")]
    public static extern uint SetErrorMode(uint mode);
    [DllImport("kernel32.dll")]
    public static extern uint GetErrorMode();
}
'@
    }
    $PreviousErrorMode = [DriveInventoryNativeErrorMode]::GetErrorMode()
    [void][DriveInventoryNativeErrorMode]::SetErrorMode($PreviousErrorMode -bor 0x0001 -bor 0x8000)
    Write-Log -Level INFO -Message "Critical device error dialogs suppressed for collector process."
}
catch {
    Write-ExceptionLog -ErrorRecord $_ -Context "Could not suppress process device error dialogs"
    Write-Host "Warning: Windows may display device error dialogs during this run."
}


try {
    try {
        Set-TemporaryAutoPlayPreference
    }
    catch {
        Write-ExceptionLog -ErrorRecord $_ -Context 'Temporary AutoPlay setup failed'
        Write-Host 'AutoPlay setting could not be changed. Continuing with the collector.'
    }

    if ($ManualEntryOnStartup) {
        Invoke-ManualEntry
        Clear-Host
    }

    Write-Host 'Ready for a USB drive. Insert one to begin, or press M for manual entry.'
    Write-Host 'Waiting for a drive...'
    Write-Host ''

    while ($true) {
        if (Test-ManualEntryHotkey) {
            Invoke-ManualEntry
            Clear-Host
            Write-Host 'Returning to automatic recording. Insert a USB drive or press M for manual entry.'
            Write-Host ''
        }
        $CurrentDisks = Get-TargetUsbDisks
        $CurrentNumbers = @{}

        foreach ($Disk in $CurrentDisks) {
            $DiskNumber = [int]$Disk.Number
            $CurrentNumbers[$DiskNumber] = $true

            if ($ConnectedDisks.ContainsKey($DiskNumber)) {
                continue
            }

            # Count this insertion as handled even when the read or save fails.
            # Retry only after Windows reports a removal and a new insertion.
            $ConnectedDisks[$DiskNumber] = $true

            Clear-Host
            Write-Host "USB drive detected on Disk $DiskNumber."
            Write-Host "Reading drive identity..."

            Write-Log -Level INFO -Message (
                "USB disk detected: number={0}, friendlyName='{1}', bridgeSerial='{2}', sizeBytes={3}, busType={4}, isBoot={5}, isSystem={6}" -f
                $DiskNumber,
                $Disk.FriendlyName,
                $Disk.SerialNumber,
                $Disk.Size,
                $Disk.BusType,
                $Disk.IsBoot,
                $Disk.IsSystem
            )

            # Give the USB bridge/adapter time to settle before the first query.
            Start-Sleep -Milliseconds $InitialSettleMilliseconds

            try {
                $Drive = Get-DriveInformation -Disk $Disk
            }
            catch {
                Write-Host ""
                Write-Host "ERROR: Unable to read drive."
                Write-Host $_.Exception.Message
                Write-Host ""
                Write-Host "Remove the drive and reinsert it to retry."
                Write-Host "Debug log: $LogPath"
                Write-Host ""
                continue
            }

            # Duplicate protection. A previously recorded serial number is not
            # added to the workbook a second time.
            if (
                $Drive.SerialNumber -ne "N/A" -and
                $KnownSerials.ContainsKey($Drive.SerialNumber)
            ) {
                Write-Host ""
                Write-Host "DUPLICATE DRIVE DETECTED"
                Write-Host "------------------------"
                Write-Host "Make:     $($Drive.Make)"
                Write-Host "Model:    $($Drive.Model)"
                Write-Host "Serial:   $($Drive.SerialNumber)"
                Write-Host "Capacity: $($Drive.Capacity)"
                Write-Host "Type:     $($Drive.Type)"
                Write-Host ""
                Write-Host "This serial number already exists in the workbook."
                Write-Host "No new row was added."
                Write-Host ""

                Write-Log -Level WARN -Message (
                    "Duplicate serial skipped: disk={0}, serial='{1}', model='{2}', type='{3}'" -f
                    $DiskNumber, $Drive.SerialNumber, $Drive.Model, $Drive.Type
                )

                try {
                    [console]::Beep(500, 400)
                }
                catch {}

                continue
            }


            $Record = [PSCustomObject]@{
                Make         = $Drive.Make
                Model        = $Drive.Model
                SerialNumber = $Drive.SerialNumber
                Capacity     = $Drive.Capacity
                Type         = $Drive.Type
            }

            [void]$Inventory.Add($Record)

            try {
                Write-XlsxInventory -Path $OutputPath -Records $Inventory
            }
            catch {
                # The workbook did not save, so roll back the in-memory append.
                $Inventory.RemoveAt($Inventory.Count - 1)

                Write-Host ""
                Write-Host "ERROR: Drive was read, but the workbook could not be updated."
                Write-Host $_.Exception.Message
                Write-Host ""
                Write-Host "Remove the drive, correct the workbook issue, and reinsert it to retry."
                Write-Host "Debug log: $LogPath"
                Write-Host ""
                continue
            }

            if ($Drive.SerialNumber -ne "N/A") {
                $KnownSerials[$Drive.SerialNumber] = $true
            }

            $RecordedRow = $Inventory.Count + 1

            Write-Host ""
            Write-Host "DRIVE RECORDED"
            Write-Host "--------------"
            Write-Host "Make:     $($Drive.Make)"
            Write-Host "Model:    $($Drive.Model)"
            Write-Host "Serial:   $($Drive.SerialNumber)"
            Write-Host "Capacity: $($Drive.Capacity)"
            Write-Host "Type:     $($Drive.Type)"
            Write-Host ""
            Write-Host "Saved as row $RecordedRow."
            Write-Host ""

            Write-Log -Level INFO -Message (
                "Drive recorded: row={0}, make='{1}', model='{2}', serial='{3}', capacity='{4}', type='{5}'" -f
                $RecordedRow, $Drive.Make, $Drive.Model, $Drive.SerialNumber, $Drive.Capacity, $Drive.Type
            )
            Write-Host "Remove this drive and insert the next drive."
            Write-Host ""

            try {
                [console]::Beep(1000, 150)
                [console]::Beep(1200, 150)
            }
            catch {}
        }

        # Record removal transitions as well. This gives the operator a clear
        # indication that the previous drive is fully gone before the next swap.
        foreach ($PreviousDiskNumber in @($ConnectedDisks.Keys)) {
            if (-not $CurrentNumbers.ContainsKey($PreviousDiskNumber)) {
                Write-Host "Disk $PreviousDiskNumber removed. Ready for the next drive."
                Write-Host ""
                Write-Log -Level INFO -Message "USB disk removal detected: disk=$PreviousDiskNumber"
            }
        }

        # Anything removed since the previous cycle disappears here.
        $ConnectedDisks = $CurrentNumbers

        Start-Sleep -Seconds $PollSeconds
    }
}
catch [System.Management.Automation.PipelineStoppedException] {
    Write-Log -Level INFO -Message "Collector stopped by user/pipeline interruption."
}
catch {
    Write-Host ""
    Write-Host "FATAL ERROR: The collector encountered an unexpected exception."
    Write-Host $_.Exception.Message
    Write-Host "Debug log: $LogPath"
    Write-Host ""

    Write-ExceptionLog -ErrorRecord $_ -Context "Unhandled collector exception"
}
finally {
    try {
        Restore-AutoPlayPreference
    }
    catch {
        Write-ExceptionLog -ErrorRecord $_ -Context 'AutoPlay preference could not be restored'
        Write-Host 'WARNING: AutoPlay could not be restored. Check Settings > Bluetooth & devices > AutoPlay.'
    }
    if ($null -ne $PreviousErrorMode) {
        [void][DriveInventoryNativeErrorMode]::SetErrorMode($PreviousErrorMode)
    }
    Write-Log -Level INFO -Message ("Collector stopped. Final in-memory record count={0}" -f $Inventory.Count)

    Write-Host ""
    Write-Host "Inventory collector stopped."
    Write-Host "Inventory saved:"
    Write-Host "  $OutputPath"
    Write-Host "Debug log:"
    Write-Host "  $LogPath"
    Write-Host ""
}

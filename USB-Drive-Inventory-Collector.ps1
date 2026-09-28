<#
.SYNOPSIS
    USB drive inventory collector for Windows, with GUI and console modes.

.DESCRIPTION
    Watches for USB-connected physical drives and records media identity for
    inventory, asset tracking, auditing, and other drive-identification workflows.

    For each newly detected USB drive, the collector:
      - Uses smartctl transport autodetection plus safe read-only fallbacks for
        common NVMe-to-USB and SATA-to-USB bridge families.
      - Captures these default fields:
            Make
            Model
            Serial Number
            Reported Capacity
            Type
      - Offers [S] setup for optional identity columns; -SetupOnStartup opens it before scanning.
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
    [ValidateSet('GUI', 'CLI', 'WorkerFunctions')][string]$Mode = 'GUI',
    [switch]$HideConsoleOnGuiReady,
    # The double-click launcher requests elevation through this switch.
    [switch]$ElevateOnStartup,
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
    [int]$SmartctlTimeoutSeconds = 30,

    # Suppresses the install prompt and exits if smartctl is missing.
    [switch]$NoDependencyInstallPrompt,

    # Opens manual entry immediately (also available with M while polling).
    [switch]$ManualEntryOnStartup,

    # Opens workbook column setup before probing any connected drives.
    [switch]$SetupOnStartup
)

$ErrorActionPreference = "Stop"
if ($ElevateOnStartup -and $Mode -ne 'WorkerFunctions') {
    $CurrentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $CurrentPrincipal = [Security.Principal.WindowsPrincipal]::new($CurrentIdentity)
    if (-not $CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $ChildArguments = '-NoLogo -NoProfile -STA '
        if ($Mode -eq 'GUI') { $ChildArguments += '-NoExit ' }
        $ChildArguments += '-File "{0}" -Mode {1}' -f $PSCommandPath, $Mode
        if ($HideConsoleOnGuiReady) { $ChildArguments += ' -HideConsoleOnGuiReady' }
        try {
            Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList $ChildArguments `
                -WorkingDirectory $PSScriptRoot -Verb RunAs -ErrorAction Stop
        }
        catch {
            Write-Host "Unable to launch as Administrator: $($_.Exception.Message)"
            throw
        }
        exit 0
    }
}
$ScriptVersion = "4.0.0"
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


function Show-ConsoleHeading {
    param ([string]$Title)

    Write-Host $Title
    Write-Host ('-' * $Title.Length)
}

function Ensure-SmartctlDependency {
    param ([scriptblock]$ConfirmInstall)

    $Existing = Find-SmartctlExecutable

    if (-not [string]::IsNullOrWhiteSpace($Existing)) {
        return $Existing
    }

    Write-Host ""
    Show-ConsoleHeading 'DEPENDENCY REQUIRED'
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

    $Response = if ($null -ne $ConfirmInstall) {
        if (& $ConfirmInstall) { 'Y' } else { 'N' }
    }
    else { Read-Host "Install smartmontools now using WinGet? [Y/N]" }

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


function Get-DriveInterface {
    param ([object]$SmartctlJson, [string]$Model = 'N/A')

    $Protocol = [string]$SmartctlJson.device.protocol
    $Text = @($SmartctlJson.smartctl.output) -join "`n"
    if ($Protocol -match 'NVMe' -or $null -ne $SmartctlJson.nvme_version) { return 'NVMe' }
    if ($Text -match '(?im)^Transport Type:\s*Parallel\b' -or $Protocol -match '^(PATA|IDE)$') { return 'PATA' }
    if ($null -ne $SmartctlJson.sata_version -or $Protocol -eq 'SATA' -or
        $Text -match '(?im)^SATA Version is:') { return 'SATA' }
    # WD Caviar Blue specifications identify this model as a 3.5-inch PATA HDD.
    # Use an exact model family, never a broad WD prefix or missing SATA metadata.
    if ($Model -match '^(?:WDC\s+)?WD800AAJB(?:-|$)') { return 'PATA' }
    if ($Protocol -eq 'ATA' -or $null -ne $SmartctlJson.ata_version) { return 'ATA (interface unknown)' }
    return (Get-CleanValue $Protocol)
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

    $Interface = Get-DriveInterface -SmartctlJson $SmartctlJson -Model $Model
    if ($Interface -eq 'PATA' -and $Model -match '^(?:WDC\s+)?WD800AAJB(?:-|$)') {
        if ([string]::IsNullOrWhiteSpace($FormFactor)) { $FormFactor = '3.5 inches' }
        if ($null -eq $RotationRate) { $RotationRate = 7200 }
    }

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

    if ($Interface -in @('PATA', 'ATA (interface unknown)')) {
        $Media = if ($IsSolidState) { 'SSD' } elseif ($IsRotational) { 'HDD' } else { 'Drive' }
        $Bus = if ($Interface -eq 'PATA') { 'PATA' } else { 'ATA' }
        $Size = if ($FormFactor -match '^(1\.8|2\.5|3\.5) inches$') { "$($Matches[1])-inch " } else { '' }
        return "$Size$Bus $Media"
    }

    if ($Interface -eq 'SATA') {
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
            # Include original identity text: PATA transport is not always a JSON field.
            [void]$Arguments.Add("-jo")

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
                    Interface        = Get-DriveInterface -SmartctlJson $Json -Model $Model
                    FirmwareVersion  = Get-CleanValue $Json.firmware_version
                    ModelFamily      = Get-CleanValue $Json.model_family
                    FormFactor       = Get-CleanValue $Json.form_factor.name
                    RotationRate     = Get-CleanValue $Json.rotation_rate
                    CapacityBytes    = Get-CleanValue $CapacityBytes
                    LogicalBlockSize = Get-CleanValue $Json.logical_block_size
                    PhysicalBlockSize = Get-CleanValue $Json.physical_block_size
                    AtaVersion       = Get-CleanValue $Json.ata_version.string
                    SataVersion      = Get-CleanValue $Json.sata_version.string
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


function Initialize-InventoryColumns {
    $script:ColumnCatalog = [ordered]@{
        Make = 'Make'; Model = 'Model'; SerialNumber = 'Serial Number'
        Capacity = 'Reported Capacity'; Type = 'Type'
        Interface = 'Interface'; FirmwareVersion = 'Firmware Version'
        ModelFamily = 'Model Family'; FormFactor = 'Form Factor'
        RotationRate = 'Rotation Rate (RPM)'; CapacityBytes = 'Capacity (Bytes)'
        LogicalBlockSize = 'Logical Sector Size (Bytes)'
        PhysicalBlockSize = 'Physical Sector Size (Bytes)'
        AtaVersion = 'ATA Version'; SataVersion = 'SATA Version'
        Protocol = 'Reported Protocol'; Transport = 'Probe Transport'
    }
    $script:CoreColumns = @('Make', 'Model', 'SerialNumber', 'Capacity', 'Type')
    $script:SelectedColumns = @($script:CoreColumns)
}

function Get-XlsxColumnName {
    param ([int]$Number)
    $Name = ''
    while ($Number -gt 0) {
        $Number--
        $Name = [string][char](65 + ($Number % 26)) + $Name
        $Number = [int][Math]::Floor($Number / 26)
    }
    return $Name
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

        $HeaderMap = @{}
        $LoadedColumns = @()
        $HeaderRow = $SheetXml.SelectSingleNode("//*[local-name()='sheetData']/*[local-name()='row'][@r='1']")
        if ($null -eq $HeaderRow) { throw 'The workbook has no header row.' }
        foreach ($Cell in $HeaderRow.SelectNodes("./*[local-name()='c']")) {
            $Header = (Get-XlsxCellText -Cell $Cell -SharedStrings $SharedStrings).Trim()
            if ([string]::IsNullOrWhiteSpace($Header)) { continue }
            $Key = @($script:ColumnCatalog.Keys | Where-Object { $script:ColumnCatalog[$_] -eq $Header })
            if ($Header -eq 'Capacity') { $Key = @('Capacity') }
            if ($Key.Count -ne 1 -or $LoadedColumns -contains $Key[0]) {
                throw "Unsupported or duplicate column '$Header'. Use a collector workbook; no changes were made."
            }
            if ($Cell.GetAttribute('r') -notmatch '^([A-Z]+)1$') { throw 'Invalid header cell reference.' }
            $HeaderMap[$Matches[1]] = $Key[0]
            $LoadedColumns += $Key[0]
        }
        foreach ($Key in $script:CoreColumns) {
            if ($LoadedColumns -notcontains $Key) { throw "Required column '$($script:ColumnCatalog[$Key])' is missing." }
        }
        foreach ($RowNode in $SheetXml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")) {
            $RowNumber = 0
            [void][int]::TryParse($RowNode.GetAttribute('r'), [ref]$RowNumber)
            if ($RowNumber -le 1) { continue }
            $Values = [ordered]@{}
            foreach ($Key in $LoadedColumns) { $Values[$Key] = 'N/A' }
            $HasData = $false
            foreach ($Cell in $RowNode.SelectNodes("./*[local-name()='c']")) {
                if ($Cell.GetAttribute('r') -match '^([A-Z]+)\d+$') {
                    $Column = $Matches[1]
                    $Value = Get-XlsxCellText -Cell $Cell -SharedStrings $SharedStrings
                    if ($HeaderMap.ContainsKey($Column)) {
                        $Values[$HeaderMap[$Column]] = Get-CleanValue $Value
                        if (-not [string]::IsNullOrWhiteSpace($Value)) { $HasData = $true }
                    }
                    elseif (-not [string]::IsNullOrWhiteSpace($Value)) {
                        throw "Data without a column header in row $RowNumber. No changes were made."
                    }
                }
            }
            if ($HasData) { [void]$Records.Add([PSCustomObject]$Values) }
        }
        $script:SelectedColumns = $LoadedColumns

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
    for ($Index = 0; $Index -lt $script:SelectedColumns.Count; $Index++) {
        $Width = if ($script:SelectedColumns[$Index] -eq 'Model') { 35 } else { 25 }
        [void]$SheetBuilder.Append(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f ($Index + 1), $Width))
    }
    [void]$SheetBuilder.Append('</cols><sheetData><row r="1">')
    for ($Index = 0; $Index -lt $script:SelectedColumns.Count; $Index++) {
        $Column = Get-XlsxColumnName ($Index + 1)
        $Header = ConvertTo-XmlText $script:ColumnCatalog[$script:SelectedColumns[$Index]]
        [void]$SheetBuilder.Append(('<c r="{0}1" t="inlineStr" s="1"><is><t>{1}</t></is></c>' -f $Column, $Header))
    }
    [void]$SheetBuilder.Append('</row>')
    $RowNumber = 2
    foreach ($Record in $Records) {
        [void]$SheetBuilder.Append(('<row r="{0}">' -f $RowNumber))
        for ($Index = 0; $Index -lt $script:SelectedColumns.Count; $Index++) {
            $Column = Get-XlsxColumnName ($Index + 1)
            $Value = ConvertTo-XmlText (Get-CleanValue $Record.($script:SelectedColumns[$Index]))
            [void]$SheetBuilder.Append(('<c r="{0}{1}" t="inlineStr"><is><t>{2}</t></is></c>' -f $Column, $RowNumber, $Value))
        }
        [void]$SheetBuilder.Append('</row>')
        $RowNumber++
    }
    [void]$SheetBuilder.Append('</sheetData>')
    if ($RowNumber -gt 2) {
        $LastColumn = Get-XlsxColumnName $script:SelectedColumns.Count
        [void]$SheetBuilder.Append(('<autoFilter ref="A1:{0}{1}"/>' -f $LastColumn, ($RowNumber - 1)))
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


if ($Mode -eq 'WorkerFunctions') { return }

if ($Mode -eq 'CLI') {
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

Initialize-InventoryColumns
$Inventory = [System.Collections.ArrayList]::new()
$KnownSerials = @{}

if (Test-Path -LiteralPath $OutputPath) {
    Write-Host 'Opening existing inventory...'

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
    Write-Host 'Creating inventory...'

    Write-XlsxInventory -Path $OutputPath -Records $Inventory
    Write-Log -Level INFO -Message "New inventory workbook created."
}


}

function Show-CollectorHeader {
    Show-ConsoleHeading "USB Drive Inventory Collector v$ScriptVersion"
    Write-Host 'Maintainer: Shannon Wetnight | https://github.com/ShannonWetnight/usb-drive-inventory-collector'
    Write-Host ''
}

function Show-CollectorUsage {
    Show-ConsoleHeading 'USAGE'
    Write-Host '1. Insert one USB drive at a time. The workbook is saved after each drive.'
    Write-Host '2. Remove the recorded drive, then insert the next.'
    Write-Host '3. Press [M] for manual entry or [D] for technical details while waiting.'
    Write-Host '4. Press [S] for setup to choose extra workbook columns.'
    Write-Host '5. Press [Ctrl+C] when finished.'
    Write-Host ''
}

function Show-CollectorWaitingScreen {
    Clear-Host
    Show-CollectorHeader
    Show-CollectorUsage
    Write-Host 'Waiting for a USB drive...'
    Write-Host ''
}

function Show-CollectorTechnicalDetails {
    Clear-Host
    Show-CollectorHeader
    Show-ConsoleHeading 'TECHNICAL DETAILS'
    Write-Host 'Scope:     USB physical drives (boot and system disks excluded)'
    Write-Host 'Adapters:  smartctl auto-detection with NVMe/SATA USB transport fallbacks'
    Write-Host "Output:    $OutputPath"
    Write-Host "Log:       $LogPath"
    Write-Host 'Backend:   Direct XLSX (no Excel COM)'
    Write-Host "smartctl:  $SmartctlVersion"
    Write-Host "Timeout:   $SmartctlTimeoutSeconds seconds per smartctl process"
    Write-Host "Columns:   $($script:SelectedColumns -join ', ')"
    Write-Host ''
    Write-Host 'Press [D] to hide details, [M] for manual entry, or [S] for setup.'
    Write-Host 'Press [Ctrl+C] to stop.'
    Write-Host ''
}

function Invoke-CollectorSetup {
    $Optional = @($script:ColumnCatalog.Keys | Where-Object { $script:CoreColumns -notcontains $_ })
    $Chosen = @($script:SelectedColumns)
    $Message = ''
    while ($true) {
        Clear-Host
        Show-CollectorHeader
        Show-ConsoleHeading 'WORKBOOK SETUP'
        Write-Host 'Default columns: Make, Model, Serial Number, Reported Capacity, Type.'
        Write-Host 'Choose optional identity fields below. Unavailable values are recorded as N/A.'
        Write-Host 'Manual records and older rows receive N/A for added fields.'
        Write-Host ''
        for ($Index = 0; $Index -lt $Optional.Count; $Index++) {
            $Mark = if ($Chosen -contains $Optional[$Index]) { 'X' } else { ' ' }
            Write-Host ('{0,2}. [{1}] {2}' -f ($Index + 1), $Mark, $script:ColumnCatalog[$Optional[$Index]])
        }
        Write-Host ''
        Write-Host 'Enter a field number to toggle it.'
        Write-Host '[A] Select all optional fields  [D] Restore default columns'
        Write-Host '[Y] Apply to this workbook     [C] Cancel setup'
        Write-Host 'Note: A backup is created before changing columns.'
        Write-Host 'Removed columns will be excluded from this workbook; their data stays in the backup.'
        if ($Message) { Write-Host ''; Write-Host $Message }
        $Choice = Read-Host 'Choose an option'
        if ($null -eq $Choice) { return }
        $Choice = $Choice.Trim()
        $Message = ''
        if ($Choice -eq 'C') { return }
        if ($Choice -eq 'A') { $Chosen = @($script:ColumnCatalog.Keys); continue }
        if ($Choice -eq 'D') { $Chosen = @($script:CoreColumns); continue }
        if ($Choice -eq 'Y') {
            $Previous = @($script:SelectedColumns)
            $Next = @($script:ColumnCatalog.Keys | Where-Object { $Chosen -contains $_ })
            if (($Previous -join '|') -eq ($Next -join '|')) { return }
            try {
                $Backup = "$OutputPath.before-setup-$([guid]::NewGuid().ToString('N')).xlsx"
                Copy-Item -LiteralPath $OutputPath -Destination $Backup -ErrorAction Stop
                Write-Log -Level INFO -Message "Setup backup created: '$Backup'"
                $script:SelectedColumns = $Next
                Write-XlsxInventory -Path $OutputPath -Records $Inventory
                Write-Log -Level INFO -Message "Workbook columns updated: $($Next -join ', ')"
                return
            }
            catch {
                $script:SelectedColumns = $Previous
                Write-ExceptionLog -ErrorRecord $_ -Context 'Workbook setup failed'
                $Message = "Setup was not applied. $($_.Exception.Message)"
                continue
            }
        }
        $Number = 0
        if ([int]::TryParse($Choice, [ref]$Number) -and $Number -ge 1 -and $Number -le $Optional.Count) {
            $Key = $Optional[$Number - 1]
            if ($Chosen -contains $Key) { $Chosen = @($Chosen | Where-Object { $_ -ne $Key }) }
            else { $Chosen += $Key }
        }
        else { $Message = "Choose a number from 1 to $($Optional.Count), [A], [D], [Y], or [C]." }
    }
}
if ($Mode -eq 'CLI') {

Clear-Host
Show-CollectorHeader
Show-CollectorUsage

# Tracks only USB disks that are currently attached. A removed disk disappears
# from this table; a later insertion on the same Windows disk number is new.
$ConnectedDisks = @{}

# Windows Settings stores the current user's AutoPlay switch here. This is a
# user preference, not the machine/user AutoPlay policy managed by Group Policy.
$AutoPlayRegistryPath = 'Software\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers'
$AutoPlayRestore = $null

}

function Set-TemporaryAutoPlayPreference {
    param ([scriptblock]$AskDisable)
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

        if ($null -ne $AskDisable) {
            $Choice = if (& $AskDisable) { 'Y' } else { 'N' }
        }
        else {
            do {
                $Choice = (Read-Host 'AutoPlay can open drive folders or show pop-ups. Disable it while collecting? [Y/N]').Trim()
            } while ($Choice -notin @('Y', 'N'))
        }

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
        $Hint = if ($AllowBack) { '[Enter] for N/A; :back or :cancel to review' } else { '[Enter] for N/A; :cancel to stop' }
        $Answer = Read-Host "$Label ($Hint)"
        if ($null -eq $Answer) { return $null }
        if ($Answer.Trim() -eq ':cancel' -or $Answer.Trim() -eq ':back') {
            if ($AllowBack) { return ':back' }
            if ($Answer.Trim() -eq ':cancel') { return $null }
        }
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

    Write-Host ''
    Write-Host "${Label}:"
    for ($Index = 0; $Index -lt $Options.Count; $Index++) {
        Write-Host ("  {0}. {1}" -f ($Index + 1), $Options[$Index])
    }
    Write-Host ''
    while ($true) {
        $Hint = if ($AllowBack) { '[Enter] for N/A; :back or :cancel to review' } else { '[Enter] for N/A; :cancel to stop' }
        $Answer = Read-Host "Choose a number ($Hint)"
        if ($null -eq $Answer) { return $null }
        if ($Answer.Trim() -eq ':cancel' -or $Answer.Trim() -eq ':back') {
            if ($AllowBack) { return ':back' }
            if ($Answer.Trim() -eq ':cancel') { return $null }
        }
        if ([string]::IsNullOrWhiteSpace($Answer)) { return 'N/A' }
        $Number = 0
        if ([int]::TryParse($Answer.Trim(), [ref]$Number) -and $Number -ge 1 -and $Number -le $Options.Count) {
            return $Options[$Number - 1]
        }
        Write-Host "Enter a number from 1 to $($Options.Count)."
    }
}

function Get-ManualDriveTypeGroups {
    return [ordered]@{
        Standard = @(
            '1.8-inch SATA SSD', '2.5-inch IDE HDD', '2.5-inch SATA HDD',
            '2.5-inch SATA SSD', '3.5-inch IDE HDD', '3.5-inch SATA HDD',
            'IDE Drive', 'IDE HDD', 'IDE SSD', 'M.2 NVMe SSD',
            'M.2 SATA SSD', 'mSATA SSD', 'NVMe SSD',
            'SATA Drive', 'SATA HDD', 'SATA SSD'
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
}

function Read-ManualDriveType {
    param ([switch]$AllowBack)

    $Groups = Get-ManualDriveTypeGroups

    $Options = @()
    Write-Host ''
    Write-Host 'Drive type:'
    foreach ($Group in $Groups.Keys) {
        Write-Host ''
        Write-Host "  ${Group}:"
        foreach ($Option in $Groups[$Group]) {
            $Options += $Option
            Write-Host ('    {0}. {1}' -f $Options.Count, $Option)
        }
    }
    Write-Host ''
    while ($true) {
        $Hint = if ($AllowBack) { '[Enter] for N/A; :back or :cancel to review' } else { '[Enter] for N/A; :cancel to stop' }
        $Answer = Read-Host "Choose a number ($Hint)"
        if ($null -eq $Answer) { return $null }
        if ($Answer.Trim() -eq ':cancel' -or $Answer.Trim() -eq ':back') {
            if ($AllowBack) { return ':back' }
            if ($Answer.Trim() -eq ':cancel') { return $null }
        }
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
        $Hint = if ($AllowBack) { '[Enter] for N/A; :back or :cancel to review' } else { '[Enter] for N/A; :cancel to stop' }
        $Amount = Read-Host "Capacity (number only; $Hint)"
        if ($null -eq $Amount) { return $null }
        if ($Amount.Trim() -eq ':cancel' -or $Amount.Trim() -eq ':back') {
            if ($AllowBack) { return ':back' }
            if ($Amount.Trim() -eq ':cancel') { return $null }
        }
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
    if ($AllowBack) { Show-ConsoleHeading 'EDIT CAPACITY - UNIT' }
    else { Show-ConsoleHeading 'Step 4 of 5 - Capacity unit' }
    Write-Host "Capacity amount: $Amount"
    Write-Host ''
    while ($true) {
        $Unit = Read-ManualSelection -Label 'Capacity unit' -Options @('B', 'KB', 'MB', 'GB', 'TB', 'PB', 'Other') -AllowBack:$AllowBack
        if ($null -eq $Unit) { return $null }
        if ($Unit -eq ':back') { return ':back' }
        if ($Unit -eq 'N/A') { return 'N/A' }
        if ($Unit -eq 'Other') {
            Clear-Host
            if ($AllowBack) { Show-ConsoleHeading 'EDIT CAPACITY - CUSTOM UNIT' }
            else { Show-ConsoleHeading 'Step 4 of 5 - Custom capacity unit' }
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
                if ($AllowBack) { Show-ConsoleHeading 'EDIT DRIVE TYPE - CUSTOM' }
                else { Show-ConsoleHeading 'Step 5 of 5 - Custom drive type' }
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
    Show-ConsoleHeading 'REVIEW MANUAL DRIVE'
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
    $SavedRow = $null

    while ($true) {
        Clear-Host
        Show-ConsoleHeading 'MANUAL DRIVE ENTRY'
        Write-Host 'Note: Press [Enter] to record N/A for a field.'
        Write-Host 'Type :cancel during entry to return to automatic recording.'
        Write-Host ''
        if ($null -eq $NextMode -and $Inventory.Count -gt 0) {
            while ($true) {
                Write-Host '[N] New drive'
                Write-Host '[L] Copy last saved drive with a new serial'
                Write-Host '[R] Return to automatic recording'
                $NextMode = Read-Host 'Choose [N/L/R]'
                if ($null -eq $NextMode) { return }
                $NextMode = $NextMode.Trim()
                if ($NextMode -eq 'R') { return }
                if ($NextMode -eq 'N' -or $NextMode -eq 'L') { break }
                Write-Host 'Choose [N], [L], or [R].'
            }
        }
        if ($null -eq $NextMode) { $NextMode = 'N' }

        if ($NextMode -eq 'L') {
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
        $FirstField = if ($NextMode -eq 'L') { 3 } else { 1 }
        $LastField = if ($NextMode -eq 'L') { 3 } else { 5 }
        for ($Index = $FirstField; $Index -le $LastField; $Index++) {
            Clear-Host
            if ($NextMode -eq 'L') {
                Show-ConsoleHeading 'Step 1 of 1 - Serial number'
                Write-Host 'Copying the last saved drive. Enter a new serial number.'
                if ($null -ne $SavedRow) { Write-Host "Previous drive saved as row $SavedRow." }
            }
            else {
                Show-ConsoleHeading "Step $Index of 5 - $($FieldLabels[$Index - 1])"
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
        :ReviewLoop while ($true) {
            Show-ManualRecord -Record $Record
            $Duplicate = $Record.SerialNumber -ne 'N/A' -and $KnownSerials.ContainsKey($Record.SerialNumber)
            if ($Duplicate) {
                Write-Host 'This serial number is already in the workbook.'
                Write-Host ''
                Write-Host '[S] Enter a different serial for this record'
                Write-Host '[C] Cancel this record'
                $DuplicateAction = Read-Host 'Choose [S/C]'
                if ($null -eq $DuplicateAction) { return }
                $DuplicateAction = $DuplicateAction.Trim()
                if ($DuplicateAction -eq 'C') {
                    Write-Log -Level INFO -Message 'Manual entry cancelled at duplicate serial review.'
                    return
                }
                if ($DuplicateAction -eq 'S') {
                    Clear-Host
                    Show-ConsoleHeading 'CHANGE SERIAL NUMBER'
                    Write-Host 'Note: :back or :cancel keeps the current value.'
                    Write-Host ''
                    $NewSerial = Read-ManualField -Field 3 -AllowBack
                    if ($null -eq $NewSerial) { return }
                    if ($NewSerial -ne ':back') { $Record.SerialNumber = $NewSerial }
                }
                Clear-Host
                continue ReviewLoop
            }
            elseif ($Record.SerialNumber -eq 'N/A') {
                Write-Host 'Serial N/A cannot be checked for duplicates.'
                Write-Host ''
            }

            Write-Host '[Y] Save and choose next action'
            Write-Host '[L] Save and copy with a new serial'
            Write-Host '[E] Edit a field'
            Write-Host '[C] Cancel this record'
            $Action = Read-Host 'Choose [Y/L/E/C]'
            if ($null -eq $Action) { return }
            $Action = $Action.Trim()
            if ($Action -eq 'C') {
                Write-Log -Level INFO -Message 'Manual entry cancelled at review.'
                return
            }
            if ($Action -eq 'E') {
                $FieldNumber = 0
                Write-Host 'Note: [Enter], :back, or :cancel returns to review.'
                $Choice = Read-Host 'Field to edit [1-5]'
                if ($null -eq $Choice) { return }
                if ([string]::IsNullOrWhiteSpace($Choice) -or $Choice.Trim() -eq ':back' -or $Choice.Trim() -eq ':cancel') {
                    Clear-Host
                    continue ReviewLoop
                }
                if (-not [int]::TryParse($Choice.Trim(), [ref]$FieldNumber) -or $FieldNumber -lt 1 -or $FieldNumber -gt 5) {
                    Write-Host 'Choose a field number from [1-5].'
                    continue
                }
                Clear-Host
                Show-ConsoleHeading "EDIT FIELD $FieldNumber"
                Write-Host 'Note: :back or :cancel keeps the current value.'
                Write-Host ''
                $Value = Read-ManualField -Field $FieldNumber -AllowBack
                if ($null -eq $Value) { return }
                if ($Value -eq ':back') {
                    Clear-Host
                    continue ReviewLoop
                }
                $Record.($Fields[$FieldNumber - 1]) = $Value
                Clear-Host
                continue ReviewLoop
            }
            if ($Action -ne 'Y' -and $Action -ne 'L') {
                Clear-Host
                continue ReviewLoop
            }
            if ($Duplicate) {
                Clear-Host
                continue ReviewLoop
            }

            $SaveAndCopy = $Action -eq 'L'
            [void]$Inventory.Add($Record)
            try {
                Write-XlsxInventory -Path $OutputPath -Records $Inventory
            }
            catch {
                $Inventory.RemoveAt($Inventory.Count - 1)
                Write-ExceptionLog -ErrorRecord $_ -Context 'Manual drive workbook save failed'
                Write-Host "Save failed: $($_.Exception.Message)"
                Write-Host 'Check the workbook, then try saving again.'
                Write-Host "Debug log: $LogPath"
                Write-Host ''
                continue
            }

            if ($Record.SerialNumber -ne 'N/A') { $KnownSerials[$Record.SerialNumber] = $true }
            $SavedRow = $Inventory.Count + 1
            Clear-Host
            Show-ConsoleHeading 'MANUAL DRIVE RECORDED'
            Write-Host "Make:     $($Record.Make)"
            Write-Host "Model:    $($Record.Model)"
            Write-Host "Serial:   $($Record.SerialNumber)"
            Write-Host "Capacity: $($Record.Capacity)"
            Write-Host "Type:     $($Record.Type)"
            Write-Host ''
            Write-Host "Saved as row $SavedRow."
            Write-Host ''
            Write-Log -Level INFO -Message (
                "Manual drive recorded: row={0}, make='{1}', model='{2}', serial='{3}', capacity='{4}', type='{5}'" -f
                ($Inventory.Count + 1), $Record.Make, $Record.Model, $Record.SerialNumber, $Record.Capacity, $Record.Type
            )
            break
        }

        if ($SaveAndCopy) {
            $NextMode = 'L'
            continue
        }

        while ($true) {
            Write-Host '[A] Add another manual drive'
            Write-Host '[L] Copy this drive with a new serial'
            Write-Host '[R] Return to automatic recording'
            $Next = Read-Host 'Choose [A/L/R]'
            if ($null -eq $Next) { return }
            $Next = $Next.Trim()
            if ($Next -eq 'A') { $NextMode = 'N'; break }
            if ($Next -eq 'L') { $NextMode = 'L'; break }
            if ($Next -eq 'R') { return }
            Write-Host 'Choose [A], [L], or [R].'
        }
    }
}

function Read-CollectorHotkey {
    if ($script:ConsoleHotkeysUnavailable) { return $null }
    try {
        while ([Console]::KeyAvailable) {
            $Key = [Console]::ReadKey($true)
            if ($Key.Modifiers -ne 0) { continue }
            if ($Key.Key -eq [ConsoleKey]::M) { return 'M' }
            if ($Key.Key -eq [ConsoleKey]::D) { return 'D' }
            if ($Key.Key -eq [ConsoleKey]::S) { return 'S' }
        }
    }
    catch {
        $script:ConsoleHotkeysUnavailable = $true
        if (-not $script:ConsoleHotkeyWarningShown) {
            $script:ConsoleHotkeyWarningShown = $true
            Write-Log -Level WARN -Message 'This PowerShell host does not support console hotkeys. Use -ManualEntryOnStartup for manual entry.'
            Write-Host 'Console hotkeys unavailable. Use -ManualEntryOnStartup or -SetupOnStartup when starting the script.'
        }
    }
    return $null
}

if ($Mode -eq 'GUI') {
$ErrorActionPreference = 'Stop'
if (-not $IsWindows -and [Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    throw 'The GUI requires Windows.'
}
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA) {
    throw 'Launch the GUI in an STA PowerShell session (powershell.exe -STA or pwsh -STA).'
}
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:CollectorScriptPath = $PSCommandPath

$DefaultOutputDirectory = Join-Path $PSScriptRoot 'Output'
[void](New-Item -ItemType Directory -Path $DefaultOutputDirectory -Force)
if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = Join-Path $DefaultOutputDirectory 'Inventory.xlsx' }
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
[void](New-Item -ItemType Directory -Path (Split-Path -Parent $OutputPath) -Force)
if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogDirectory = Join-Path $DefaultOutputDirectory 'Logs'
    [void](New-Item -ItemType Directory -Path $LogDirectory -Force)
    $LogPath = Join-Path $LogDirectory ('USB-Drive-Inventory-Collector-{0}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
}
else {
    $LogPath = [IO.Path]::GetFullPath($LogPath)
    [void](New-Item -ItemType Directory -Path (Split-Path -Parent $LogPath) -Force)
}
$script:AutoPlayRegistryPath = 'Software\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers'
$script:AutoPlayRestore = $null
$script:ConnectedDisks = @{}
$script:Inventory = [Collections.ArrayList]::new()
$script:KnownSerials = @{}
$script:ActiveWorker = $null
$script:WorkerHandle = $null
$script:ModalActive = $false
$script:Paused = $false
$script:ClosingAfterProbe = $false
$script:Started = $false
$script:PreviousErrorMode = $null
$script:ProgressQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

function New-GuiLabel {
    param($Text, $X, $Y, $Width, $Height = 24, $Size = 10, $Bold = $false)
    $Label = [System.Windows.Forms.Label]::new()
    $Label.Text = $Text; $Label.SetBounds($X, $Y, $Width, $Height)
    $Style = if ($Bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    $Label.Font = [System.Drawing.Font]::new('Segoe UI', $Size, $Style)
    return $Label
}

function New-GuiButton {
    param($Text, $X, $Y, $Width, $Height = 34)
    $Button = [System.Windows.Forms.Button]::new()
    $Button.Text = $Text; $Button.SetBounds($X, $Y, $Width, $Height)
    $Button.Font = [System.Drawing.Font]::new('Segoe UI', 9)
    return $Button
}

function New-GuiTextBox {
    param($X, $Y, $Width)
    $Box = [System.Windows.Forms.TextBox]::new()
    $Box.SetBounds($X, $Y, $Width, 27)
    $Box.Font = [System.Drawing.Font]::new('Segoe UI', 10)
    return $Box
}

function Show-GuiError {
    param ([string]$Message)
    [void][System.Windows.Forms.MessageBox]::Show($Message, 'USB Drive Inventory Collector', 'OK', 'Error')
}

$Form = [System.Windows.Forms.Form]::new()
$Form.Text = "USB Drive Inventory Collector v$ScriptVersion"
$Form.ClientSize = [System.Drawing.Size]::new(1020, 690)
$Form.MinimumSize = [System.Drawing.Size]::new(880, 570)
$Form.StartPosition = 'CenterScreen'
$Form.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$Form.BackColor = [System.Drawing.Color]::FromArgb(246, 248, 251)
$Layout = [System.Windows.Forms.TableLayoutPanel]::new()
$Layout.Dock = 'Fill'; $Layout.RowCount = 5; $Layout.ColumnCount = 1
[void]$Layout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new('Percent',100))
foreach ($Height in @(85,56,74)) {
    [void]$Layout.RowStyles.Add([System.Windows.Forms.RowStyle]::new('Absolute',$Height))
}
[void]$Layout.RowStyles.Add([System.Windows.Forms.RowStyle]::new('Percent',100))
[void]$Layout.RowStyles.Add([System.Windows.Forms.RowStyle]::new('Absolute',42))
$Form.Controls.Add($Layout)

$Header = [System.Windows.Forms.Panel]::new()
$Header.Dock = 'Top'; $Header.Height = 85; $Header.BackColor = [System.Drawing.Color]::FromArgb(32, 55, 78)
$Title = New-GuiLabel "USB Drive Inventory Collector  v$ScriptVersion" 22 13 640 37 18 $true
$Title.ForeColor = [System.Drawing.Color]::White
$Header.Controls.Add($Title)
$Subtitle = New-GuiLabel 'One drive at a time. Every saved record is written to the workbook.' 24 51 800 24 9
$Subtitle.ForeColor = [System.Drawing.Color]::FromArgb(215, 229, 238)
$Header.Controls.Add($Subtitle)
$Layout.Controls.Add($Header,0,0)

$Actions = [System.Windows.Forms.FlowLayoutPanel]::new()
$Actions.Dock = 'Top'; $Actions.Height = 56; $Actions.Padding = [System.Windows.Forms.Padding]::new(12, 10, 0, 0)
$Actions.WrapContents = $false
$Layout.Controls.Add($Actions,0,1)
$PauseButton = New-GuiButton 'Pause scanning' 0 0 126
$ManualButton = New-GuiButton 'Manual entry' 0 0 120
$CopyButton = New-GuiButton 'Copy last' 0 0 105
$SetupButton = New-GuiButton 'Workbook setup' 0 0 137
$DetailsButton = New-GuiButton 'Technical details' 0 0 138
$ExitButton = New-GuiButton 'Finish' 0 0 90
foreach ($Button in @($PauseButton,$ManualButton,$CopyButton,$SetupButton,$DetailsButton,$ExitButton)) {
    $Button.Margin = [System.Windows.Forms.Padding]::new(0, 0, 8, 0)
    [void]$Actions.Controls.Add($Button)
}

$StatusPanel = [System.Windows.Forms.Panel]::new()
$StatusPanel.Dock = 'Top'; $StatusPanel.Height = 74
$StatusLabel = New-GuiLabel 'Initializing...' 24 5 920 28 12 $true
$GuidanceLabel = New-GuiLabel 'Insert one USB drive at a time.' 24 36 950 23 9
$StatusPanel.Controls.AddRange(@($StatusLabel,$GuidanceLabel))
$Layout.Controls.Add($StatusPanel,0,2)

$Tabs = [System.Windows.Forms.TabControl]::new()
$Tabs.Dock = 'Fill'
$Layout.Controls.Add($Tabs,0,3)
$RecordTab = [System.Windows.Forms.TabPage]::new('Recorded drives')
$ActivityTab = [System.Windows.Forms.TabPage]::new('Activity')
$DetailsTab = [System.Windows.Forms.TabPage]::new('Details')
$Tabs.TabPages.AddRange(@($RecordTab,$ActivityTab,$DetailsTab))
$Grid = [System.Windows.Forms.DataGridView]::new()
$Grid.Dock = 'Fill'; $Grid.ReadOnly = $true; $Grid.AllowUserToAddRows = $false
$Grid.AllowUserToDeleteRows = $false; $Grid.RowHeadersVisible = $false
$Grid.AutoSizeColumnsMode = 'DisplayedCells'; $Grid.SelectionMode = 'FullRowSelect'
$RecordTab.Controls.Add($Grid)
$Activity = [System.Windows.Forms.ListBox]::new()
$Activity.Dock = 'Fill'; $Activity.HorizontalScrollbar = $true
$Activity.Font = [System.Drawing.Font]::new('Consolas', 9)
$ActivityTab.Controls.Add($Activity)
$Details = [System.Windows.Forms.TextBox]::new()
$Details.Multiline = $true; $Details.ReadOnly = $true
$Details.ScrollBars = 'Both'; $Details.WordWrap = $false; $Details.Dock = 'Fill'
$Details.Font = [System.Drawing.Font]::new('Consolas', 9)
$DetailsTab.Controls.Add($Details)
$Bottom = [System.Windows.Forms.Panel]::new()
$Bottom.Dock = 'Bottom'; $Bottom.Height = 42
$CountLabel = New-GuiLabel 'Records: 0' 18 9 300 24 9
$PathLabel = New-GuiLabel '' 312 9 685 24 9
$PathLabel.AutoEllipsis = $true
$Bottom.Controls.AddRange(@($CountLabel,$PathLabel))
$Layout.Controls.Add($Bottom,0,4)

function Add-Activity {
    param ([string]$Message, [string]$Level = 'INFO')
    $StatusLabel.Text = $Message
    $Line = '{0:HH:mm:ss}  {1,-5}  {2}' -f (Get-Date), $Level, $Message
    $Activity.Items.Insert(0, $Line)
    if ($Activity.Items.Count -gt 500) { $Activity.Items.RemoveAt(500) }
}

function Update-Details {
    $PathLabel.Text = "Workbook: $OutputPath"
    $CountLabel.Text = "Records: $($script:Inventory.Count)"
    $Details.Lines = @(
        "USB Drive Inventory Collector v$ScriptVersion",
        'Maintainer: Shannon Wetnight',
        'Repository: https://github.com/ShannonWetnight/usb-drive-inventory-collector',
        "Workbook: $OutputPath", "Debug log: $LogPath", "smartctl: $SmartctlVersion",
        "Scope: USB physical drives; boot and system disks excluded",
        "Transport: smartctl autodetection plus USB adapter fallbacks",
        'Workbook backend: Direct XLSX (no Excel COM)',
        "Timeout: $SmartctlTimeoutSeconds seconds per smartctl process",
        "Workbook columns: $(@($script:SelectedColumns | ForEach-Object { $script:ColumnCatalog[$_] }) -join ', ')"
    )
}

function Update-Grid {
    $Grid.Columns.Clear()
    foreach ($Key in $script:SelectedColumns) {
        $Column = [System.Windows.Forms.DataGridViewTextBoxColumn]::new()
        $Column.Name = $Key; $Column.HeaderText = $script:ColumnCatalog[$Key]
        [void]$Grid.Columns.Add($Column)
    }
    $Grid.Rows.Clear()
    foreach ($Record in $script:Inventory) {
        [void]$Grid.Rows.Add([object[]]@($script:SelectedColumns | ForEach-Object { Get-CleanValue $Record.$_ }))
    }
    Update-Details
}

function Save-GuiRecord {
    param ([object]$Record, [string]$Source)
    if ($Record.SerialNumber -ne 'N/A' -and $script:KnownSerials.ContainsKey($Record.SerialNumber)) {
        throw "Serial $($Record.SerialNumber) is already in the workbook."
    }
    foreach ($Key in $script:SelectedColumns) {
        if ($null -eq $Record.PSObject.Properties[$Key]) {
            $Record | Add-Member -NotePropertyName $Key -NotePropertyValue 'N/A'
        }
    }
    [void]$script:Inventory.Add($Record)
    try { Write-XlsxInventory -Path $OutputPath -Records $script:Inventory }
    catch {
        $script:Inventory.RemoveAt($script:Inventory.Count - 1)
        Write-ExceptionLog -ErrorRecord $_ -Context "$Source workbook save failed"
        throw
    }
    if ($Record.SerialNumber -ne 'N/A') { $script:KnownSerials[$Record.SerialNumber] = $true }
    $Row = $script:Inventory.Count + 1
    Update-Grid
    $CopyButton.Enabled = $true
    Add-Activity "$Source recorded as row $Row`: $($Record.Model) / $($Record.SerialNumber)"
    Write-Log -Level INFO -Message "$Source recorded: row=$Row, model='$($Record.Model)', serial='$($Record.SerialNumber)', type='$($Record.Type)'"
    return $Row
}

function Get-GuiCollectorState {
    return [PSCustomObject]@{
        CoreColumns = @($script:CoreColumns)
        SelectedColumns = @($script:SelectedColumns)
        ColumnCatalog = $script:ColumnCatalog
        Inventory = $script:Inventory
        KnownSerials = $script:KnownSerials
    }
}

function Set-GuiSelectedColumns {
    param ([string[]]$Columns)
    $script:SelectedColumns = @($Columns)
}

function ConvertTo-ManualGuiRecord {
    param ([string]$Make, [string]$Model, [string]$Serial,
        [string]$Amount, [string]$Unit, [string]$CustomUnit,
        [string]$DriveType, [string]$CustomType)
    $Make = $Make.Trim(); $Model = $Model.Trim(); $Serial = $Serial.Trim()
    $Amount = $Amount.Trim(); $Unit = $Unit.Trim(); $CustomUnit = $CustomUnit.Trim()
    $DriveType = $DriveType.Trim(); $CustomType = $CustomType.Trim()
    if ($Make -and ($Make.Length -gt 80 -or $Make -cnotmatch "\A[A-Za-z0-9][A-Za-z0-9 .&()+'/_-]{0,79}\z")) {
        throw 'Make: use up to 80 plain letters, digits, spaces, or standard punctuation.'
    }
    if ($Model -and ($Model.Length -gt 100 -or $Model -cnotmatch '\A[A-Za-z0-9][A-Za-z0-9 .+/_-]{0,99}\z')) {
        throw 'Model: use up to 100 plain letters, digits, spaces, or standard punctuation.'
    }
    if ($Serial -and ($Serial.Length -gt 100 -or $Serial -cnotmatch '\A[A-Za-z0-9][A-Za-z0-9./_-]{0,99}\z')) {
        throw 'Serial: use up to 100 letters, digits, periods, slashes, underscores, or hyphens.'
    }
    $Capacity = 'N/A'
    if ($Amount) {
        if ($Amount.Length -gt 22 -or $Amount -cnotmatch '\A[0-9]{1,15}(?:\.[0-9]{1,6})?\z') {
            throw 'Capacity: enter a positive number, such as 1 or 0.005, without the unit.'
        }
        $Number = [decimal]::Parse($Amount, [Globalization.CultureInfo]::InvariantCulture)
        if ($Number -le 0) { throw 'Capacity must be greater than zero.' }
        if ($Unit -eq 'Other') {
            if ($CustomUnit -cnotmatch '\A[A-Za-z]{1,12}\z') { throw 'Custom capacity unit must be 1–12 letters.' }
            $Unit = $CustomUnit
        }
        if ($Unit -and $Unit -notin @('B','KB','MB','GB','TB','PB') -and $Unit -ne $CustomUnit) {
            throw 'Choose a listed capacity unit or enter a custom unit.'
        }
        if ($Unit) {
            if ($Unit -eq 'B' -and $Number -ne [decimal]::Truncate($Number)) { throw 'Bytes must be a whole number.' }
            $Capacity = '{0} {1}' -f $Number.ToString('0.######', [Globalization.CultureInfo]::InvariantCulture), $Unit
        }
    }
    $Type = 'N/A'
    if ($DriveType -eq 'Other') {
        if ($CustomType -and ($CustomType.Length -gt 60 -or
                $CustomType -cnotmatch '\A[A-Za-z0-9][A-Za-z0-9 .()+/_-]{0,59}\z')) {
            throw 'Custom drive type must use up to 60 plain letters, digits, spaces, or standard punctuation.'
        }
        if ($CustomType) { $Type = $CustomType }
    }
    elseif ($DriveType) { $Type = $DriveType }
    return [PSCustomObject]@{
        Make = if ($Make) { $Make } else { 'N/A' }
        Model = if ($Model) { $Model.ToUpperInvariant() } else { 'N/A' }
        SerialNumber = if ($Serial) { $Serial.ToUpperInvariant() } else { 'N/A' }
        Capacity = $Capacity; Type = $Type
    }
}

function Show-GuiReview {
    param ([object]$Record, [bool]$Duplicate)
    $Dialog = [System.Windows.Forms.Form]::new()
    $Dialog.Text = if ($Duplicate) { 'Duplicate serial number' } else { 'Review manual drive' }
    $Dialog.ClientSize = [System.Drawing.Size]::new(490, 346)
    $Dialog.StartPosition = 'CenterParent'; $Dialog.FormBorderStyle = 'FixedDialog'
    $Dialog.MaximizeBox = $false; $Dialog.MinimizeBox = $false
    $Lines = @("Make:     $($Record.Make)", "Model:    $($Record.Model)",
        "Serial:   $($Record.SerialNumber)", "Capacity: $($Record.Capacity)", "Type:     $($Record.Type)")
    $Summary = [System.Windows.Forms.TextBox]::new()
    $Summary.Multiline = $true; $Summary.ReadOnly = $true
    $Summary.Font = [System.Drawing.Font]::new('Consolas', 11)
    $Summary.SetBounds(20, 20, 450, 158)
    $Summary.Lines = $Lines
    $Dialog.Controls.Add($Summary)
    $Note = New-GuiLabel '' 20 188 450 48 9
    if ($Duplicate) { $Note.Text = 'This serial is already in the workbook. Change it or cancel this record.' }
    elseif ($Record.SerialNumber -eq 'N/A') { $Note.Text = 'Serial N/A cannot be checked for duplicates.' }
    else { $Note.Text = 'Review these values before saving a new row.' }
    $Dialog.Controls.Add($Note)
    $Dialog.Tag = 'Edit'
    if ($Duplicate) {
        $Change = New-GuiButton 'Change serial' 20 272 140
        $Cancel = New-GuiButton 'Cancel record' 180 272 140
        $Change.Add_Click({ $Dialog.Tag = 'Serial'; $Dialog.Close() }.GetNewClosure())
        $Cancel.Add_Click({ $Dialog.Tag = 'Cancel'; $Dialog.Close() }.GetNewClosure())
        $Dialog.Controls.AddRange(@($Change,$Cancel))
    }
    else {
        $Save = New-GuiButton 'Save & next' 14 272 105
        $Copy = New-GuiButton 'Save & copy' 125 272 105
        $Edit = New-GuiButton 'Edit fields' 236 272 105
        $Cancel = New-GuiButton 'Cancel record' 347 272 128
        $Save.Add_Click({ $Dialog.Tag = 'Save'; $Dialog.Close() }.GetNewClosure())
        $Copy.Add_Click({ $Dialog.Tag = 'Copy'; $Dialog.Close() }.GetNewClosure())
        $Edit.Add_Click({ $Dialog.Tag = 'Edit'; $Dialog.Close() }.GetNewClosure())
        $Cancel.Add_Click({ $Dialog.Tag = 'Cancel'; $Dialog.Close() }.GetNewClosure())
        $Dialog.Controls.AddRange(@($Save,$Copy,$Edit,$Cancel))
    }
    try { [void]$Dialog.ShowDialog($Form); return [string]$Dialog.Tag }
    finally { $Dialog.Dispose() }
}

function Show-ManualEntry {
    param ([switch]$CopyLast)
    $script:ModalActive = $true
    $Dialog = [System.Windows.Forms.Form]::new()
    $Dialog.Text = if ($CopyLast) { 'Copy last saved drive' } else { 'Manual drive entry' }
    $Dialog.ClientSize = [System.Drawing.Size]::new(560, 410)
    $Dialog.StartPosition = 'CenterParent'; $Dialog.FormBorderStyle = 'FixedDialog'
    $Dialog.MaximizeBox = $false; $Dialog.MinimizeBox = $false
    $Dialog.Font = [System.Drawing.Font]::new('Segoe UI', 10)
    $Note = New-GuiLabel 'Leave a field blank for N/A. Model and serial are saved in uppercase.' 20 12 520 30 9
    $Dialog.Controls.Add($Note)
    $Labels = @('1. Make', '2. Model', '3. Serial number', '4. Capacity (number only)', 'Capacity unit', '5. Drive type')
    for ($Index = 0; $Index -lt $Labels.Count; $Index++) {
        $Y = 47 + ($Index * 43)
        $Dialog.Controls.Add((New-GuiLabel $Labels[$Index] 20 $Y 193 27 9))
    }
    $MakeBox = New-GuiTextBox 219 47 320
    $ModelBox = New-GuiTextBox 219 90 320
    $SerialBox = New-GuiTextBox 219 133 320
    $AmountBox = New-GuiTextBox 219 176 175
    $ModelBox.CharacterCasing = 'Upper'; $SerialBox.CharacterCasing = 'Upper'
    $UnitBox = [System.Windows.Forms.ComboBox]::new()
    $UnitBox.DropDownStyle = 'DropDownList'; $UnitBox.SetBounds(219, 219, 135, 28)
    [void]$UnitBox.Items.AddRange([object[]]@('N/A','B','KB','MB','GB','TB','PB','Other'))
    $UnitBox.SelectedIndex = 0
    $OtherUnitBox = New-GuiTextBox 365 219 174
    $OtherUnitBox.Visible = $false
    $TypeBox = [System.Windows.Forms.ComboBox]::new()
    $TypeBox.DropDownStyle = 'DropDownList'; $TypeBox.SetBounds(219, 262, 320, 28)
    [void]$TypeBox.Items.Add('N/A')
    $Groups = Get-ManualDriveTypeGroups
    foreach ($Group in $Groups.Keys) {
        foreach ($Option in $Groups[$Group]) {
            [void]$TypeBox.Items.Add("$Group / $Option")
        }
    }
    $TypeBox.SelectedIndex = 0
    $OtherTypeBox = New-GuiTextBox 219 299 320
    $OtherTypeBox.Visible = $false
    $UnitBox.Add_SelectedIndexChanged({ $OtherUnitBox.Visible = $UnitBox.Text -eq 'Other' }.GetNewClosure())
    $TypeBox.Add_SelectedIndexChanged({ $OtherTypeBox.Visible = $TypeBox.Text -eq 'Other / Other' }.GetNewClosure())
    $Dialog.Controls.AddRange(@($MakeBox,$ModelBox,$SerialBox,$AmountBox,$UnitBox,$OtherUnitBox,$TypeBox,$OtherTypeBox))

    if ($CopyLast -and $script:Inventory.Count -gt 0) {
        $Previous = $script:Inventory[$script:Inventory.Count - 1]
        foreach ($Pair in @(@($MakeBox,$Previous.Make),@($ModelBox,$Previous.Model))) {
            if ($Pair[1] -ne 'N/A') { $Pair[0].Text = [string]$Pair[1] }
        }
        if ($Previous.Capacity -match '^([0-9]+(?:\.[0-9]+)?)\s+(.+)$') {
            $AmountBox.Text = $Matches[1]
            $Unit = $Matches[2]
            if ($UnitBox.Items.Contains($Unit)) { $UnitBox.SelectedItem = $Unit }
            else { $UnitBox.SelectedItem = 'Other'; $OtherUnitBox.Text = $Unit }
        }
        $TypeValue = [string]$Previous.Type
        foreach ($Item in $TypeBox.Items) {
            if ([string]$Item -like "* / $TypeValue") { $TypeBox.SelectedItem = $Item; break }
        }
        if ($TypeBox.SelectedIndex -eq 0 -and $TypeValue -ne 'N/A') {
            $TypeBox.SelectedItem = 'Other / Other'; $OtherTypeBox.Text = $TypeValue
        }
    }
    $ReviewButton = New-GuiButton 'Review drive' 20 355 130
    $ReturnButton = New-GuiButton 'Return to scanning' 162 355 165
    $ReturnButton.Add_Click({ $Dialog.Close() }.GetNewClosure())
    $ReviewButton.Add_Click({
        try {
            $Type = if ($TypeBox.SelectedIndex -gt 0) { ([string]$TypeBox.SelectedItem -split ' / ',2)[1] } else { '' }
            $Record = ConvertTo-ManualGuiRecord -Make $MakeBox.Text -Model $ModelBox.Text `
                -Serial $SerialBox.Text -Amount $AmountBox.Text -Unit $(if ($UnitBox.Text -eq 'N/A') { '' } else { $UnitBox.Text }) `
                -CustomUnit $OtherUnitBox.Text -DriveType $Type -CustomType $OtherTypeBox.Text
            $State = Get-GuiCollectorState
            $Duplicate = $Record.SerialNumber -ne 'N/A' -and $State.KnownSerials.ContainsKey($Record.SerialNumber)
            $Action = Show-GuiReview $Record $Duplicate
            if ($Action -eq 'Serial') { $SerialBox.Focus(); $SerialBox.SelectAll(); return }
            if ($Action -eq 'Cancel') { $Dialog.Close(); return }
            if ($Action -eq 'Edit') { return }
            try { $Row = Save-GuiRecord $Record 'Manual drive' }
            catch { Show-GuiError $_.Exception.Message; return }
            [void][System.Windows.Forms.MessageBox]::Show("Drive saved as row $Row.", 'Drive recorded', 'OK', 'Information')
            $SerialBox.Clear()
            if ($Action -eq 'Save') {
                $MakeBox.Clear(); $ModelBox.Clear(); $AmountBox.Clear()
                $UnitBox.SelectedIndex = 0; $OtherUnitBox.Clear()
                $TypeBox.SelectedIndex = 0; $OtherTypeBox.Clear()
            }
            else { $Dialog.Text = 'Copy saved drive with a new serial' }
            $SerialBox.Focus()
        }
        catch { Show-GuiError $_.Exception.Message }
    }.GetNewClosure())
    $Dialog.Controls.AddRange(@($ReviewButton,$ReturnButton))
    try { [void]$Dialog.ShowDialog($Form) }
    finally { $Dialog.Dispose(); $script:ModalActive = $false }
}

function Show-WorkbookSetup {
    $script:ModalActive = $true
    $Dialog = [System.Windows.Forms.Form]::new()
    $Dialog.Text = 'Workbook setup'
    $Dialog.ClientSize = [System.Drawing.Size]::new(550, 540)
    $Dialog.StartPosition = 'CenterParent'; $Dialog.FormBorderStyle = 'FixedDialog'
    $Dialog.MaximizeBox = $false; $Dialog.MinimizeBox = $false
    $Dialog.Controls.Add((New-GuiLabel 'Default columns: Make, Model, Serial Number, Reported Capacity, Type.' 20 15 510 26 9))
    $Dialog.Controls.Add((New-GuiLabel 'Select extra identity fields to save for each drive:' 20 43 510 26 9))
    $Choices = [System.Windows.Forms.CheckedListBox]::new()
    $Choices.CheckOnClick = $true; $Choices.SetBounds(20, 75, 510, 322)
    $Choices.Font = [System.Drawing.Font]::new('Segoe UI', 10)
    $Optional = @($script:ColumnCatalog.Keys | Where-Object { $script:CoreColumns -notcontains $_ })
    foreach ($Key in $Optional) {
        [void]$Choices.Items.Add($script:ColumnCatalog[$Key], ($script:SelectedColumns -contains $Key))
    }
    $Dialog.Controls.Add($Choices)
    $Note = New-GuiLabel 'Older and manual records receive N/A for extra fields. Removing a field drops it from the active workbook; a backup keeps the old data.' 20 402 505 46 9
    $Dialog.Controls.Add($Note)
    $All = New-GuiButton 'Select all' 20 478 112
    $Defaults = New-GuiButton 'Defaults' 140 478 112
    $Apply = New-GuiButton 'Apply' 290 478 112
    $Cancel = New-GuiButton 'Cancel' 410 478 112
    $All.Add_Click({ for ($I=0;$I -lt $Choices.Items.Count;$I++) { $Choices.SetItemChecked($I,$true) } }.GetNewClosure())
    $Defaults.Add_Click({ for ($I=0;$I -lt $Choices.Items.Count;$I++) { $Choices.SetItemChecked($I,$false) } }.GetNewClosure())
    $Cancel.Add_Click({ $Dialog.Close() }.GetNewClosure())
    $Apply.Add_Click({
        $State = Get-GuiCollectorState
        $Chosen = @($State.CoreColumns)
        for ($I=0;$I -lt $Optional.Count;$I++) {
            if ($Choices.GetItemChecked($I)) { $Chosen += $Optional[$I] }
        }
        if (($Chosen -join '|') -eq ($State.SelectedColumns -join '|')) { $Dialog.Close(); return }
        $Previous = @($State.SelectedColumns)
        try {
            $Backup = "$OutputPath.before-setup-$([guid]::NewGuid().ToString('N')).xlsx"
            Copy-Item -LiteralPath $OutputPath -Destination $Backup -ErrorAction Stop
            Write-Log -Level INFO -Message "GUI setup backup created: '$Backup'"
            Set-GuiSelectedColumns $Chosen
            Write-XlsxInventory -Path $OutputPath -Records $State.Inventory
            Update-Grid
            Add-Activity 'Workbook columns updated.'
            Write-Log -Level INFO -Message "GUI columns updated: $($Chosen -join ', ')"
            $Dialog.Close()
        }
        catch {
            Set-GuiSelectedColumns $Previous
            Write-ExceptionLog -ErrorRecord $_ -Context 'GUI workbook setup failed'
            Show-GuiError "Setup was not applied. $($_.Exception.Message)"
        }
    }.GetNewClosure())
    $Dialog.Controls.AddRange(@($All,$Defaults,$Apply,$Cancel))
    try { [void]$Dialog.ShowDialog($Form) }
    finally { $Dialog.Dispose(); $script:ModalActive = $false }
}

$script:WorkerScript = @'
param($SourceFile, $KnownNumbers, $Preferred, $Pattern, $Retries, $Settle, $RetryDelay,
    $Timeout, $Smartctl, $LogFile, $RunIdentifier, $ProgressQueue)
. $SourceFile -Mode WorkerFunctions

$UsbDevicePattern = $Pattern
$SmartctlRetries = $Retries
$InitialSettleMilliseconds = $Settle
$RetryDelayMilliseconds = $RetryDelay
$SmartctlTimeoutSeconds = $Timeout
$script:SmartctlPath = $Smartctl
$script:PreferredTransportByDiskNumber = $Preferred
$LogPath = $LogFile
$RunId = $RunIdentifier
$Numbers = @()
try {
    $Disks = @(Get-TargetUsbDisks)
    $Numbers = @($Disks | ForEach-Object { [int]$_.Number })
    $Target = @($Disks | Where-Object { $KnownNumbers -notcontains [int]$_.Number } | Select-Object -First 1)
    if ($Target.Count -gt 0) {
        $Disk = $Target[0]
        $ProgressQueue.Enqueue("USB drive detected on Disk $($Disk.Number). Reading drive identity...")
        Start-Sleep -Milliseconds $InitialSettleMilliseconds
        try {
            $Drive = Get-DriveInformation -Disk $Disk
            [PSCustomObject]@{ Numbers=$Numbers; DiskNumber=[int]$Disk.Number
                Drive=$Drive; Error=$null; Preferred=$script:PreferredTransportByDiskNumber }
        }
        catch {
            Write-ExceptionLog -ErrorRecord $_ -Context "GUI identity disk $($Disk.Number) failed"
            [PSCustomObject]@{ Numbers=$Numbers; DiskNumber=[int]$Disk.Number
                Drive=$null; Error=$_.Exception.Message; Preferred=$script:PreferredTransportByDiskNumber }
        }
    }
    else {
        [PSCustomObject]@{ Numbers=$Numbers; DiskNumber=$null
            Drive=$null; Error=$null; Preferred=$script:PreferredTransportByDiskNumber }
    }
}
catch {
    Write-ExceptionLog -ErrorRecord $_ -Context 'GUI disk scan failed'
    [PSCustomObject]@{ Numbers=$null; DiskNumber=$null
        Drive=$null; Error=$_.Exception.Message; Preferred=$script:PreferredTransportByDiskNumber }
}
'@

function Start-ScanWorker {
    $KnownNumbers = @($script:ConnectedDisks.Keys | ForEach-Object { [int]$_ })
    $Worker = [PowerShell]::Create()
    try {
        $null = $Worker.AddScript($script:WorkerScript).AddArgument($script:CollectorScriptPath).AddArgument($KnownNumbers).AddArgument(
            $script:PreferredTransportByDiskNumber).AddArgument($UsbDevicePattern).AddArgument(
            $SmartctlRetries).AddArgument($InitialSettleMilliseconds).AddArgument(
            $RetryDelayMilliseconds).AddArgument($SmartctlTimeoutSeconds).AddArgument(
            $script:SmartctlPath).AddArgument($LogPath).AddArgument($RunId).AddArgument($script:ProgressQueue)
        $script:ActiveWorker = $Worker
        $script:WorkerHandle = $Worker.BeginInvoke()
    }
    catch {
        $script:ActiveWorker = $null; $script:WorkerHandle = $null
        $Worker.Dispose()
        throw
    }
}

function Complete-ScanWorker {
    $Worker = $script:ActiveWorker
    try {
        $Result = @($Worker.EndInvoke($script:WorkerHandle))
        if ($Worker.HadErrors) {
            foreach ($ErrorItem in $Worker.Streams.Error) {
                Write-Log -Level WARN -Message "GUI scan warning: $ErrorItem"
            }
        }
        if ($Result.Count -eq 0) { throw 'Disk scan returned no result.' }
        $Sample = $Result[$Result.Count - 1]
        if ($null -eq $Sample.Numbers) { throw $Sample.Error }
        $Numbers = @($Sample.Numbers)
        foreach ($OldNumber in @($script:ConnectedDisks.Keys)) {
            if ($Numbers -notcontains [int]$OldNumber) {
                Add-Activity "Disk $OldNumber removed. Ready for another drive."
                Write-Log -Level INFO -Message "USB disk removal detected: disk=$OldNumber"
            }
        }
        $NewConnected = @{}
        foreach ($Number in $Numbers) {
            if ($script:ConnectedDisks.ContainsKey([int]$Number)) { $NewConnected[[int]$Number] = $true }
        }
        if ($null -ne $Sample.DiskNumber) { $NewConnected[[int]$Sample.DiskNumber] = $true }
        $script:ConnectedDisks = $NewConnected
        $script:PreferredTransportByDiskNumber = $Sample.Preferred
        if ($null -ne $Sample.DiskNumber) {
            $DiskNumber = [int]$Sample.DiskNumber
            if ($Sample.Error) {
                Add-Activity "Disk $DiskNumber could not be read: $($Sample.Error)" 'ERROR'
                $GuidanceLabel.Text = 'Remove and reinsert to retry, or use Manual entry.'
                Write-Log -Level ERROR -Message "Disk $DiskNumber read failed: $($Sample.Error)"
            }
            else {
                try {
                    $Row = Save-GuiRecord $Sample.Drive 'USB drive'
                    $GuidanceLabel.Text = "Saved as row $Row. Remove this drive and insert the next."
                    [System.Media.SystemSounds]::Asterisk.Play()
                }
                catch {
                    if ($_.Exception.Message -match '^Serial .+ is already in the workbook') {
                        Add-Activity "Disk $DiskNumber duplicate: $($Sample.Drive.SerialNumber). No row added." 'WARN'
                        Write-Log -Level WARN -Message "Duplicate serial skipped: disk=$DiskNumber, serial='$($Sample.Drive.SerialNumber)'"
                        [System.Media.SystemSounds]::Exclamation.Play()
                    }
                    else {
                        Add-Activity "Disk $DiskNumber was read, but the workbook could not be saved: $($_.Exception.Message)" 'ERROR'
                    }
                    $GuidanceLabel.Text = 'Remove and reinsert to retry, or use Manual entry.'
                }
            }
        }
    }
    finally {
        $Worker.Dispose()
        $script:ActiveWorker = $null
        $script:WorkerHandle = $null
    }
}

$Timer = [System.Windows.Forms.Timer]::new()
$Timer.Interval = [Math]::Max(250, ($PollSeconds * 1000))
$Timer.Add_Tick({
    $ProgressText = $null
    while ($script:ProgressQueue.TryDequeue([ref]$ProgressText)) {
        Add-Activity $ProgressText
        $GuidanceLabel.Text = 'Reading drive identity. Slow adapters may reach the per-probe timeout.'
        $ProgressText = $null
    }
    if ($null -ne $script:ActiveWorker) {
        if (-not $script:WorkerHandle.IsCompleted -or $script:ModalActive) { return }
        try { Complete-ScanWorker }
        catch {
            Write-ExceptionLog -ErrorRecord $_ -Context 'GUI scan result failed'
            Add-Activity "Scan error: $($_.Exception.Message)" 'ERROR'
        }
    }
    if ($script:ClosingAfterProbe) { $Form.Close(); return }
    if (-not $script:Paused -and -not $script:ModalActive -and $null -eq $script:ActiveWorker) {
        try { Start-ScanWorker }
        catch {
            Write-ExceptionLog -ErrorRecord $_ -Context 'GUI scan start failed'
            Add-Activity "Unable to start scan: $($_.Exception.Message)" 'ERROR'
            $script:Paused = $true; $PauseButton.Text = 'Resume scanning'
        }
    }
})

$PauseButton.Add_Click({
    $script:Paused = -not $script:Paused
    $PauseButton.Text = if ($script:Paused) { 'Resume scanning' } else { 'Pause scanning' }
    Add-Activity $(if ($script:Paused) { 'Scanning paused.' } else { 'Waiting for a USB drive...' })
})
$ManualButton.Add_Click({ Show-ManualEntry })
$CopyButton.Add_Click({ if ($script:Inventory.Count -gt 0) { Show-ManualEntry -CopyLast } })
$SetupButton.Add_Click({ Show-WorkbookSetup })
$DetailsButton.Add_Click({ $Tabs.SelectedTab = $DetailsTab; Update-Details })
$ExitButton.Add_Click({ $Form.Close() })

$Form.Add_Shown({
    if ($script:Started) { return }
    $script:Started = $true
    try {
        if (-not (Test-IsAdministrator)) { throw 'Run PowerShell as Administrator.' }
        if ($PSVersionTable.PSVersion -lt [version]'5.1') { throw 'PowerShell 5.1 or later is required.' }
        if ($null -eq (Get-Command Get-Disk -ErrorAction SilentlyContinue)) { throw 'Windows Storage module (Get-Disk) is unavailable.' }
        $script:SmartctlPath = Ensure-SmartctlDependency -ConfirmInstall {
            if ($NoDependencyInstallPrompt) { return $false }
            $Answer = [System.Windows.Forms.MessageBox]::Show(
                'smartmontools is required for automatic drive detection. Install it with WinGet now?',
                'Install smartmontools', 'YesNo', 'Question')
            return $Answer -eq [System.Windows.Forms.DialogResult]::Yes
        }
        $script:Identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $script:InteractiveUser = 'N/A'
        try { $script:InteractiveUser = (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).UserName }
        catch { Write-ExceptionLog -ErrorRecord $_ -Context 'Unable to find signed-in desktop user' }
        $script:SmartctlVersion = 'N/A'
        try {
            $Version = Invoke-SmartctlProcess -Arguments @('--version') -Context 'smartctl version check'
            $script:SmartctlVersion = (($Version.StdOut -split '[\r\n]+' | Select-Object -First 1).Trim())
        }
        catch { Write-ExceptionLog -ErrorRecord $_ -Context 'Unable to read smartctl version' }
        Initialize-InventoryColumns
        if (Test-Path -LiteralPath $OutputPath) {
            foreach ($Record in @(Read-XlsxInventory $OutputPath)) {
                [void]$script:Inventory.Add($Record)
                if ($Record.SerialNumber -ne 'N/A' -and -not [string]::IsNullOrWhiteSpace($Record.SerialNumber)) {
                    $script:KnownSerials[$Record.SerialNumber] = $true
                }
            }
        }
        else { Write-XlsxInventory -Path $OutputPath -Records $script:Inventory }
        Update-Grid
        $CopyButton.Enabled = $script:Inventory.Count -gt 0
        Write-Log -Level INFO -Message "GUI collector startup: version=$ScriptVersion, output='$OutputPath', records=$($script:Inventory.Count)"
        try {
            if (-not ('DriveInventoryNativeErrorMode' -as [type])) {
                Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
public static class DriveInventoryNativeErrorMode {
    [DllImport("kernel32.dll")] public static extern uint SetErrorMode(uint mode);
    [DllImport("kernel32.dll")] public static extern uint GetErrorMode();
}
'@
            }
            $script:PreviousErrorMode = [DriveInventoryNativeErrorMode]::GetErrorMode()
            [void][DriveInventoryNativeErrorMode]::SetErrorMode($script:PreviousErrorMode -bor 0x0001 -bor 0x8000)
        }
        catch { Write-ExceptionLog -ErrorRecord $_ -Context 'Could not suppress process device dialogs' }
        try {
            Set-TemporaryAutoPlayPreference -AskDisable {
                $Answer = [System.Windows.Forms.MessageBox]::Show(
                    'AutoPlay may open drive folders or show pop-ups. Disable it until the collector exits?',
                    'Temporary AutoPlay setting', 'YesNo', 'Question')
                return $Answer -eq [System.Windows.Forms.DialogResult]::Yes
            }
            $DesktopSid = $null
            try {
                if ($script:InteractiveUser -ne 'N/A') {
                    $DesktopSid = ([Security.Principal.NTAccount]::new($script:InteractiveUser)).Translate(
                        [Security.Principal.SecurityIdentifier]).Value
                }
            }
            catch {}
            if ($DesktopSid -ne $script:Identity.User.Value) {
                Add-Activity 'AutoPlay prompt skipped: the signed-in desktop user differs from this account.' 'WARN'
            }
            elseif ($null -ne $script:AutoPlayRestore) { Add-Activity 'AutoPlay temporarily disabled for this Windows user.' }
        }
        catch { Write-ExceptionLog -ErrorRecord $_ -Context 'GUI AutoPlay setup failed'; Add-Activity 'AutoPlay setting could not be changed.' 'WARN' }
        if ($SetupOnStartup) { Show-WorkbookSetup }
        if ($ManualEntryOnStartup) { Show-ManualEntry }
        Add-Activity 'Waiting for a USB drive...'
        $GuidanceLabel.Text = 'Insert one drive at a time. Use Manual entry or Workbook setup at any time.'
        $Timer.Start()
        if ($HideConsoleOnGuiReady) {
            try {
                if (-not ('DriveInventoryConsoleWindow' -as [type])) {
                    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class DriveInventoryConsoleWindow {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
'@
                }
                $ConsoleWindow = [DriveInventoryConsoleWindow]::GetConsoleWindow()
                if ($ConsoleWindow -ne [IntPtr]::Zero) {
                    [void][DriveInventoryConsoleWindow]::ShowWindow($ConsoleWindow, 0)
                }
            }
            catch { Write-ExceptionLog -ErrorRecord $_ -Context 'Unable to hide GUI console window' }
        }
    }
    catch {
        Write-ExceptionLog -ErrorRecord $_ -Context 'GUI startup failed' 
        Show-GuiError $_.Exception.Message
        $Form.Close()
    }
})

$Form.Add_FormClosing({
    if ($null -ne $script:ActiveWorker) {
        $script:ClosingAfterProbe = $true
        $script:Paused = $true
        $StatusLabel.Text = 'Finishing the current drive probe before exit...'
        $_.Cancel = $true
        return
    }
    $Timer.Stop()
    try { Restore-AutoPlayPreference }
    catch {
        Write-ExceptionLog -ErrorRecord $_ -Context 'GUI AutoPlay restore failed'
        Show-GuiError 'AutoPlay could not be restored. Check Windows Settings > Bluetooth & devices > AutoPlay.'
    }
    if ($null -ne $script:PreviousErrorMode) {
        [void][DriveInventoryNativeErrorMode]::SetErrorMode($script:PreviousErrorMode)
    }
    Write-Log -Level INFO -Message "GUI collector stopped. Records=$($script:Inventory.Count)"
})

try { [void]$Form.ShowDialog() }
finally { $Timer.Dispose(); $Form.Dispose() }

if ($HideConsoleOnGuiReady) { exit 0 }
return
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

    $script:TechnicalDetailsVisible = $false
    if ($SetupOnStartup) {
        Invoke-CollectorSetup
        Show-CollectorWaitingScreen
    }
    if ($ManualEntryOnStartup) {
        Invoke-ManualEntry
        Show-CollectorWaitingScreen
    }
    elseif (-not $SetupOnStartup) {
        Write-Host 'Waiting for a USB drive...'
        Write-Host ''
    }

    while ($true) {
        $Hotkey = Read-CollectorHotkey
        if ($Hotkey -eq 'M') {
            Invoke-ManualEntry
            $script:TechnicalDetailsVisible = $false
            Show-CollectorWaitingScreen
        }
        elseif ($Hotkey -eq 'S') {
            Invoke-CollectorSetup
            $script:TechnicalDetailsVisible = $false
            Show-CollectorWaitingScreen
        }
        elseif ($Hotkey -eq 'D') {
            if ($script:TechnicalDetailsVisible) { Show-CollectorWaitingScreen }
            else { Show-CollectorTechnicalDetails }
            $script:TechnicalDetailsVisible = -not $script:TechnicalDetailsVisible
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

            $script:TechnicalDetailsVisible = $false
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
                Show-ConsoleHeading 'ERROR: Unable to read drive'
                Write-Host $_.Exception.Message
                Write-Host ""
                Write-Host '1. Remove the drive and reinsert it to retry.'
                Write-Host '2. Press [M] to record this drive manually while waiting.'
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
                Show-ConsoleHeading 'DUPLICATE DRIVE DETECTED'
                Write-Host "Make:     $($Drive.Make)"
                Write-Host "Model:    $($Drive.Model)"
                Write-Host "Serial:   $($Drive.SerialNumber)"
                Write-Host "Capacity: $($Drive.Capacity)"
                Write-Host "Type:     $($Drive.Type)"
                Write-Host ""
                Write-Host "This serial number already exists in the workbook."
                Write-Host "No new row was added."
                Write-Host ""
                Write-Host 'Remove this drive and insert the next drive, or press [M] for manual entry.'
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

            foreach ($Key in $script:SelectedColumns) {
                if ($script:CoreColumns -notcontains $Key) {
                    $Record | Add-Member -NotePropertyName $Key -NotePropertyValue (Get-CleanValue $Drive.$Key)
                }
            }
            [void]$Inventory.Add($Record)

            try {
                Write-XlsxInventory -Path $OutputPath -Records $Inventory
            }
            catch {
                # The workbook did not save, so roll back the in-memory append.
                $Inventory.RemoveAt($Inventory.Count - 1)

                Write-Host ""
                Show-ConsoleHeading 'ERROR: Workbook could not be updated'
                Write-Host 'The drive was read, but its record was not saved.'
                Write-Host $_.Exception.Message
                Write-Host ""
                Write-Host "Remove the drive, correct the workbook issue, and reinsert it to retry."
                Write-Host 'Press [M] for manual entry while waiting.'
                Write-Host "Debug log: $LogPath"
                Write-Host ""
                continue
            }

            if ($Drive.SerialNumber -ne "N/A") {
                $KnownSerials[$Drive.SerialNumber] = $true
            }

            $RecordedRow = $Inventory.Count + 1

            Write-Host ""
            Show-ConsoleHeading 'DRIVE RECORDED'
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
            Write-Host 'Remove this drive and insert the next drive.'
            Write-Host 'Press [M] for manual entry or [D] for technical details while waiting.'
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
                Write-Host "Disk $PreviousDiskNumber removed. Insert the next drive or press [M] for manual entry."
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
    Show-ConsoleHeading 'FATAL ERROR'
    Write-Host 'The collector encountered an unexpected exception.'
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
    Show-ConsoleHeading 'Inventory collector stopped'
    Write-Host "Inventory saved:"
    Write-Host "  $OutputPath"
    Write-Host "Debug log:"
    Write-Host "  $LogPath"
    Write-Host ""
}

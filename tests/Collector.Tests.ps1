$ErrorActionPreference = 'Stop'
$Source = Join-Path (Split-Path $PSScriptRoot -Parent) 'USB-Drive-Inventory-Collector.ps1'
$Tokens = $null
$Errors = $null
$Ast = [System.Management.Automation.Language.Parser]::ParseFile($Source, [ref]$Tokens, [ref]$Errors)
if ($Errors.Count) { throw ($Errors | Out-String) }
# Load functions without running Windows initialization or accessing attached disks.
foreach ($Function in $Ast.FindAll({ param($Node) $Node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
    Invoke-Expression $Function.Extent.Text
}
function Write-Log { param($Level, $Message) }
function Write-ExceptionLog { param($ErrorRecord, $Context) }
function Assert-Equal($Actual, $Expected, $Label) {
    if ($Actual -cne $Expected) { throw "$Label : expected '$Expected', got '$Actual'" }
}

$Ata = '{"device":{"protocol":"ATA"},"ata_version":{"string":"ATA/ATAPI-7"}}' | ConvertFrom-Json
Assert-Equal (Get-DriveType $Ata 'WDC WD800AAJB-00J3A0') '3.5-inch PATA HDD' 'Reported WD model'
Assert-Equal (Get-DriveType $Ata 'UNKNOWN') 'ATA Drive' 'ATA alone must not mean SATA or PATA'
$Pata = '{"device":{"protocol":"ATA"},"rotation_rate":5400,"form_factor":{"name":"2.5 inches"},"smartctl":{"output":["Transport Type: Parallel, ATA8-APT"]}}' | ConvertFrom-Json
Assert-Equal (Get-DriveType $Pata 'TEST') '2.5-inch PATA HDD' 'PATA text evidence'
$Sata = '{"device":{"protocol":"ATA"},"sata_version":{"string":"SATA 3.3"},"rotation_rate":0,"form_factor":{"name":"2.5 inches"}}' | ConvertFrom-Json
Assert-Equal (Get-DriveType $Sata 'TEST') '2.5-inch SATA SSD' 'SATA evidence'
$Nvme = '{"device":{"protocol":"NVMe"},"form_factor":{"name":"M.2"}}' | ConvertFrom-Json
Assert-Equal (Get-DriveType $Nvme 'TEST') 'M.2 NVMe SSD' 'NVMe unchanged'
Assert-Equal (Get-XlsxColumnName 27) 'AA' 'Column addressing'

$SmartctlRetries = 1
$script:PreferredTransportByDiskNumber = @{}
function Get-SmartctlTransportCandidates { param($DiskNumber) return 'auto' }
function Test-LooksLikeBridgeIdentity { param($Model, $Serial, $Disk, $Protocol) return $false }
function Invoke-SmartctlProcess {
    param($Arguments, $Context)
    if ($Arguments -notcontains '-jo') { throw 'Identity query must include original text.' }
    return [PSCustomObject]@{
        ExitCode = 0; StdErr = ''; StdOut = '{"device":{"protocol":"ATA","type":"sat"},"model_name":"WDC WD800AAJB-00J3A0","serial_number":"TEST0001","firmware_version":"01.00","user_capacity":{"bytes":80000000000},"logical_block_size":512,"physical_block_size":512,"ata_version":{"string":"ATA/ATAPI-7"}}'
    }
}
$Drive = Get-DriveInformation ([PSCustomObject]@{ Number=1; FriendlyName='Test adapter'; SerialNumber='BRIDGE' })
Assert-Equal $Drive.Type '3.5-inch PATA HDD' 'Identity query classifies reported model'
Assert-Equal $Drive.Interface 'PATA' 'Interface metadata'
Assert-Equal $Drive.FirmwareVersion '01.00' 'Firmware metadata'
Assert-Equal $Drive.LogicalBlockSize '512' 'Sector metadata'

$TestDirectory = Join-Path ([IO.Path]::GetTempPath()) ('collector-tests-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $TestDirectory)
$WorkbookSaveRetries = 1
$WorkbookRetryDelayMilliseconds = 1
$OutputPath = Join-Path $TestDirectory 'Inventory.xlsx'
$ScriptVersion = 'test'
try {
    Initialize-InventoryColumns
    $Inventory = [System.Collections.ArrayList]::new()
    [void]$Inventory.Add([PSCustomObject]@{ Make='A & B'; Model='MODEL <1>'; SerialNumber='0000123'; Capacity='80 GB'; Type='PATA HDD' })
    Write-XlsxInventory $OutputPath $Inventory
    $Loaded = @(Read-XlsxInventory $OutputPath)
    Assert-Equal $Loaded.Count 1 'Default round trip row count'
    Assert-Equal $Loaded[0].SerialNumber '0000123' 'Leading zero serial'
    Assert-Equal $Loaded[0].Make 'A & B' 'XML escaping'
    Assert-Equal $script:SelectedColumns.Count 5 'Default schema'

    $script:SelectedColumns += @('FirmwareVersion', 'RotationRate')
    $Inventory[0] | Add-Member FirmwareVersion '=literal&text'
    $Inventory[0] | Add-Member RotationRate '0'
    Write-XlsxInventory $OutputPath $Inventory
    Initialize-InventoryColumns
    $Loaded = @(Read-XlsxInventory $OutputPath)
    Assert-Equal ($script:SelectedColumns -join ',') 'Make,Model,SerialNumber,Capacity,Type,FirmwareVersion,RotationRate' 'Schema survives restart'
    Assert-Equal $Loaded[0].FirmwareVersion '=literal&text' 'Extra value remains literal text'
    Assert-Equal $Loaded[0].RotationRate '0' 'SSD zero RPM retained'
    [void]$Inventory.Add([PSCustomObject]@{ Make='Manual'; Model='TEST'; SerialNumber='0000456'; Capacity='1 GB'; Type='IDE HDD' })
    Write-XlsxInventory $OutputPath $Inventory
    $Loaded = @(Read-XlsxInventory $OutputPath)
    Assert-Equal $Loaded[1].FirmwareVersion 'N/A' 'Manual optional values'
    Assert-Equal $Loaded[0].FirmwareVersion '=literal&text' 'Existing extra value retained on append'

    # Exercise setup cancel, apply, backup, and save failure without an interactive console.
    function Clear-Host {}
    function Write-Host { param($Object) }
    function Read-Host { param($Prompt) return $script:Answers.Dequeue() }
    $Before = [IO.File]::ReadAllBytes($OutputPath)
    $script:Answers = [Collections.Generic.Queue[string]]::new()
    foreach ($Answer in @('A', 'C')) { $script:Answers.Enqueue($Answer) }
    Invoke-CollectorSetup
    Assert-Equal ([Convert]::ToBase64String([IO.File]::ReadAllBytes($OutputPath))) ([Convert]::ToBase64String($Before)) 'Cancel leaves workbook untouched'
    foreach ($Answer in @('A', 'Y')) { $script:Answers.Enqueue($Answer) }
    Invoke-CollectorSetup
    Assert-Equal $script:SelectedColumns.Count $script:ColumnCatalog.Count 'Apply all fields'
    Assert-Equal @(Get-ChildItem "$OutputPath.before-setup-*.xlsx").Count 1 'Setup backup'
    $BackupPath = (Get-ChildItem "$OutputPath.before-setup-*.xlsx")[0].FullName
    Assert-Equal ([Convert]::ToBase64String([IO.File]::ReadAllBytes($BackupPath))) ([Convert]::ToBase64String($Before)) 'Backup exact original'
    Initialize-InventoryColumns
    $Loaded = @(Read-XlsxInventory $OutputPath)
    Assert-Equal $script:SelectedColumns.Count $script:ColumnCatalog.Count 'All fields persist'
    Assert-Equal $Loaded[0].SerialNumber '0000123' 'Serial preserved with optional columns'
    Assert-Equal $Loaded[0].FirmwareVersion '=literal&text' 'Value preserved through setup'
    # Refuse an unknown column rather than silently dropping it on the next save.
    $Archive = [IO.Compression.ZipFile]::Open($OutputPath, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $Entry = $Archive.GetEntry('xl/worksheets/sheet1.xml')
        $Xml = (Get-ZipEntryText $Entry).Replace('Firmware Version', 'Unrecognized Field')
        $Entry.Delete()
        Add-ZipTextEntry $Archive 'xl/worksheets/sheet1.xml' $Xml
    }
    finally { $Archive.Dispose() }
    $Rejected = $false
    try { $null = Read-XlsxInventory $OutputPath } catch { $Rejected = $_.Exception.Message -match 'Unsupported or duplicate column' }
    Assert-Equal $Rejected $true 'Unknown schema rejected'
    $SelectedBefore = $script:SelectedColumns -join ','
    function Write-XlsxInventory { param($Path, $Records) throw 'Simulated save failure' }
    foreach ($Answer in @('D', 'Y', 'C')) { $script:Answers.Enqueue($Answer) }
    Invoke-CollectorSetup
    Assert-Equal ($script:SelectedColumns -join ',') $SelectedBefore 'Failed setup restores selection'
    [Console]::WriteLine('PASS: syntax, classification, XLSX round trips, setup cancel/apply/backup/failure.')
}
finally {
    Remove-Item -LiteralPath $TestDirectory -Recurse -Force
}

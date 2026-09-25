$ErrorActionPreference = 'Stop'
$Root = Split-Path $PSScriptRoot -Parent
$ConsolePath = Join-Path $Root 'USB-Drive-Inventory-Collector.ps1'
$GuiPath = Join-Path $Root 'USB-Drive-Inventory-Collector-GUI.ps1'
$Tokens = $null; $Errors = $null
$ConsoleAst = [Management.Automation.Language.Parser]::ParseFile($ConsolePath,[ref]$Tokens,[ref]$Errors)
if ($Errors.Count) { throw ($Errors | Out-String) }
$GuiAst = [Management.Automation.Language.Parser]::ParseFile($GuiPath,[ref]$Tokens,[ref]$Errors)
if ($Errors.Count) { throw ($Errors | Out-String) }
$ConsoleFunctions = @($ConsoleAst.FindAll({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] },$false))
foreach ($Function in $ConsoleFunctions) { Invoke-Expression $Function.Extent.Text }
$GuiFunctions = @($GuiAst.FindAll({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] },$false))
foreach ($Function in $GuiFunctions) { Invoke-Expression $Function.Extent.Text }
function Assert-Equal($Actual,$Expected,$Label) {
    if ($Actual -cne $Expected) { throw "$Label`: expected '$Expected', got '$Actual'" }
}

$Manual = ConvertTo-ManualGuiRecord -Make ' Western Digital ' -Model ' wd800aajb ' `
    -Serial ' 000abc ' -Amount '0.005' -Unit 'GB' -CustomUnit '' -DriveType 'IDE HDD' -CustomType ''
Assert-Equal $Manual.Make 'Western Digital' 'Make trimmed'
Assert-Equal $Manual.Model 'WD800AAJB' 'Model uppercased'
Assert-Equal $Manual.SerialNumber '000ABC' 'Serial uppercased with leading zeros'
Assert-Equal $Manual.Capacity '0.005 GB' 'Fractional capacity'
Assert-Equal $Manual.Type 'IDE HDD' 'Selected type'
$Blank = ConvertTo-ManualGuiRecord -Make ' ' -Model '' -Serial ' ' -Amount '' -Unit '' -CustomUnit '' -DriveType '' -CustomType ''
Assert-Equal $Blank.SerialNumber 'N/A' 'Skipped serial'
Assert-Equal $Blank.Capacity 'N/A' 'Skipped capacity'
$Bad = $false
try { $null = ConvertTo-ManualGuiRecord -Make 'A' -Model 'ok' -Serial '=formula' -Amount '1' -Unit 'GB' -CustomUnit '' -DriveType 'Other' -CustomType 'ok' }
catch { $Bad = $true }
Assert-Equal $Bad $true 'Formula-like serial rejected'
$Bad = $false
try { $null = ConvertTo-ManualGuiRecord -Make 'A' -Model 'ok' -Serial 'abc' -Amount '0' -Unit 'GB' -CustomUnit '' -DriveType '' -CustomType '' }
catch { $Bad = $true }
Assert-Equal $Bad $true 'Zero capacity rejected'
$Bad = $false
try { $null = ConvertTo-ManualGuiRecord -Make 'A' -Model 'ok' -Serial 'abc' -Amount '0.5' -Unit 'B' -CustomUnit '' -DriveType '' -CustomType '' }
catch { $Bad = $true }
Assert-Equal $Bad $true 'Fractional bytes rejected'
$Groups = Get-ManualDriveTypeGroups
foreach ($Name in @('Standard','Enterprise','Other')) {
    if ($Groups.Keys -notcontains $Name) { throw "Missing drive group $Name" }
}
Assert-Equal $Groups.Other[-1] 'Other' 'Custom option last'

# Windows Forms event handlers use GetNewClosure; shared state must still refer
# to the GUI script, rather than to the callback's dynamic module scope.
$script:CoreColumns = @('Make','SerialNumber')
$script:SelectedColumns = @('Make','SerialNumber')
$script:KnownSerials = @{ ABC = $true }
$Handler = {
    $State = Get-GuiCollectorState
    if (-not $State.KnownSerials.ContainsKey('ABC')) { throw 'Callback lost known serials.' }
    Set-GuiSelectedColumns @('Make','SerialNumber','Type')
}.GetNewClosure()
& $Handler
Assert-Equal ($script:SelectedColumns -join ',') 'Make,SerialNumber,Type' 'Callback updated GUI columns'

$Threshold = $ConsoleAst.Find({ param($Node)
    $Node -is [Management.Automation.Language.AssignmentStatementAst] -and
    $Node.Left.Extent.Text -eq '$Identity'
},$false).Extent.StartOffset
$script:WorkerDefinitions = ($ConsoleFunctions | Where-Object {
    $_.Extent.StartOffset -lt $Threshold
} | ForEach-Object Extent | ForEach-Object Text) -join "`n`n"
$WorkerAssignment = $GuiAst.Find({ param($Node)
    $Node -is [Management.Automation.Language.AssignmentStatementAst] -and
    $Node.Left.Extent.Text -eq '$script:WorkerScript'
},$false)
Invoke-Expression $WorkerAssignment.Extent.Text
[void][Management.Automation.Language.Parser]::ParseInput($script:WorkerScript,[ref]$Tokens,[ref]$Errors)
if ($Errors.Count) { throw "Background scan has syntax errors: $($Errors | Out-String)" }
$script:PreferredTransportByDiskNumber = @{}
function Get-Disk { return @() }
$WorkerResult = & ([scriptblock]::Create($script:WorkerScript)) @() @{} '*' 1 250 250 2 'smartctl' '/tmp/test.log' 'test' ([Collections.Concurrent.ConcurrentQueue[string]]::new())
Assert-Equal @($WorkerResult).Count 1 'Background scan result'
Assert-Equal @($WorkerResult.Numbers).Count 0 'Empty disk scan'

$ProbeScript = $script:WorkerScript.Replace('$Numbers = @()' + "`n" + 'try {', @'
function Get-Disk {
    [PSCustomObject]@{ Number=7; BusType='USB'; IsBoot=$false; IsSystem=$false; FriendlyName='Fixture disk' }
    [PSCustomObject]@{ Number=8; BusType='USB'; IsBoot=$false; IsSystem=$false; FriendlyName='Second disk' }
}
function Get-DriveInformation {
    param($Disk)
    [PSCustomObject]@{ Make='TEST'; Model='TEST MODEL'; SerialNumber='TEST0007'; Capacity='1 GB'; Type='USB Flash Drive' }
}
$Numbers = @()
try {
'@)
$PowerShell = [PowerShell]::Create()
try {
    $Progress = [Collections.Concurrent.ConcurrentQueue[string]]::new()
    $null = $PowerShell.AddScript($ProbeScript).AddArgument(@()).AddArgument(@{}).AddArgument('*').AddArgument(1).AddArgument(250).AddArgument(250).AddArgument(2).AddArgument('smartctl').AddArgument('/tmp/test.log').AddArgument('test').AddArgument($Progress)
    $Handle = $PowerShell.BeginInvoke()
    if (-not $Handle.AsyncWaitHandle.WaitOne(5000)) { throw 'Background probe did not complete.' }
    $Detected = @($PowerShell.EndInvoke($Handle))
    Assert-Equal $Detected.Count 1 'Background probe result count'
    Assert-Equal $Detected[0].DiskNumber 7 'Background drive detection'
    Assert-Equal $Detected[0].Drive.SerialNumber 'TEST0007' 'Background identity delivered'
    Assert-Equal @($Detected[0].Numbers).Count 2 'Both disks remain visible'
    $Text = $null
    Assert-Equal ($Progress.TryDequeue([ref]$Text)) $true 'Progress delivered while probing'
}
finally { $PowerShell.Dispose() }
$Next = & ([scriptblock]::Create($ProbeScript)) @(7) @{} '*' 1 250 250 2 'smartctl' '/tmp/test.log' 'test' ([Collections.Concurrent.ConcurrentQueue[string]]::new())
Assert-Equal $Next.DiskNumber 8 'Next unprocessed disk scanned'

Write-Output 'PASS: GUI parser, manual validation, shared type choices, and asynchronous background scan.'

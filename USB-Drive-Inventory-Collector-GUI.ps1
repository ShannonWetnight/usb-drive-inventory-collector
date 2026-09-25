<#
.SYNOPSIS
    Windows GUI for USB Drive Inventory Collector v4.0.0.
.DESCRIPTION
    Uses the collector's shared drive, workbook, and AutoPlay functions. Keep
    this file beside USB-Drive-Inventory-Collector.ps1. Launch as Administrator.
#>
param (
    [string]$OutputPath,
    [string]$LogPath,
    [string]$UsbDevicePattern = '*',
    [ValidateRange(1, 60)][int]$PollSeconds = 1,
    [ValidateRange(1, 5)][int]$SmartctlRetries = 2,
    [ValidateRange(250, 10000)][int]$InitialSettleMilliseconds = 2000,
    [ValidateRange(250, 10000)][int]$RetryDelayMilliseconds = 1250,
    [ValidateRange(1, 10)][int]$WorkbookSaveRetries = 5,
    [ValidateRange(100, 10000)][int]$WorkbookRetryDelayMilliseconds = 750,
    [ValidateRange(2, 120)][int]$SmartctlTimeoutSeconds = 30,
    [switch]$NoDependencyInstallPrompt,
    [switch]$SetupOnStartup,
    [switch]$ManualEntryOnStartup
)

$ErrorActionPreference = 'Stop'
if (-not $IsWindows -and [Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    throw 'The GUI requires Windows.'
}
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA) {
    throw 'Launch the GUI in an STA PowerShell session (powershell.exe -STA or pwsh -STA).'
}
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$ScriptVersion = '4.0.0'
$RunId = [guid]::NewGuid().ToString('N').Substring(0, 8)
$script:PreferredTransportByDiskNumber = @{}
$script:SourcePath = Join-Path $PSScriptRoot 'USB-Drive-Inventory-Collector.ps1'
if (-not (Test-Path -LiteralPath $script:SourcePath)) { throw "Collector script is missing: $script:SourcePath" }
$Tokens = $null
$ParseErrors = $null
$Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:SourcePath, [ref]$Tokens, [ref]$ParseErrors)
if ($ParseErrors.Count) { throw "Collector script has syntax errors: $($ParseErrors | Out-String)" }
$Definitions = @($Ast.FindAll({ param($Node) $Node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false))
$Required = @('Get-TargetUsbDisks', 'Get-DriveInformation', 'Get-ManualDriveTypeGroups',
    'Read-XlsxInventory', 'Write-XlsxInventory', 'Restore-AutoPlayPreference')
foreach ($Name in $Required) {
    if (@($Definitions | Where-Object Name -eq $Name).Count -ne 1) { throw "Collector function missing or ambiguous: $Name" }
}
foreach ($Definition in $Definitions) { Invoke-Expression $Definition.Extent.Text }
# Worker runspaces execute only functions defined before prerequisite checks.
$script:WorkerDefinitions = ($Definitions | Where-Object {
    $_.Extent.StartOffset -lt $Ast.Find({ param($Node)
        $Node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
        $Node.Left.Extent.Text -eq '$Identity'
    }, $false).Extent.StartOffset
} | ForEach-Object Extent | ForEach-Object Text) -join "`n`n"

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
param($KnownNumbers, $Preferred, $Pattern, $Retries, $Settle, $RetryDelay,
    $Timeout, $Smartctl, $LogFile, $RunIdentifier, $ProgressQueue)
'@ + "`n" + $script:WorkerDefinitions + @'

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
        $null = $Worker.AddScript($script:WorkerScript).AddArgument($KnownNumbers).AddArgument(
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

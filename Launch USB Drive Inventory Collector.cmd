@echo off
setlocal DisableDelayedExpansion
set "COLLECTOR_SCRIPT=%~dp0USB-Drive-Inventory-Collector.ps1"
if not exist "%COLLECTOR_SCRIPT%" (
    echo USB Drive Inventory Collector was not found beside this launcher.
    echo "%COLLECTOR_SCRIPT%"
    pause
    exit /b 1
)

echo USB Drive Inventory Collector
echo.
echo [G] GUI ^(default^)    [C] Console    [Q] Quit
setlocal EnableDelayedExpansion
set "CHOICE="
set /p "CHOICE=Choose mode [G/C/Q]: "
if /i "!CHOICE!"=="Q" exit /b 0
set "COLLECTOR_MODE=GUI"
if /i "!CHOICE!"=="C" set "COLLECTOR_MODE=CLI"
if not "!CHOICE!"=="" if /i not "!CHOICE!"=="G" if /i not "!CHOICE!"=="C" (
    echo Choose G, C, or Q.
    pause
    exit /b 1
)

powershell.exe -NoLogo -NoProfile -Command "$ErrorActionPreference='Stop'; $path=$env:COLLECTOR_SCRIPT; $mode=$env:COLLECTOR_MODE; $arguments='-NoLogo -NoProfile -STA '; if ($mode -eq 'GUI') { $arguments += '-NoExit ' }; $arguments += '-File ' + [char]34 + $path + [char]34 + ' -Mode ' + $mode; if ($mode -eq 'GUI') { $arguments += ' -HideConsoleOnGuiReady' }; Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList $arguments -WorkingDirectory (Split-Path -Parent $path) -Verb RunAs"
if errorlevel 1 (
    echo Unable to launch the collector. Check the error above.
    pause
    exit /b 1
)

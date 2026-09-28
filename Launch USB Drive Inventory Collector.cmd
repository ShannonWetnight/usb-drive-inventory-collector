@echo off
setlocal
echo USB Drive Inventory Collector
echo.
echo [G] GUI (default)    [C] Console    [Q] Quit
set "MODE="
set /p "MODE=Choose mode [G/C/Q]: "
if /i "%MODE%"=="Q" exit /b 0
if /i "%MODE%"=="C" goto console
if not "%MODE%"=="" if /i not "%MODE%"=="G" (
    echo Choose G, C, or Q.
    pause
    exit /b 1
)
set "TARGET_SCRIPT=%~dp0USB-Drive-Inventory-Collector-GUI.ps1"
set "WINDOW_STYLE=-WindowStyle Hidden"
goto launch
:console
set "TARGET_SCRIPT=%~dp0USB-Drive-Inventory-Collector.ps1"
set "WINDOW_STYLE="
:launch
if not exist "%TARGET_SCRIPT%" (
    echo The selected script was not found next to this launcher:
    echo "%TARGET_SCRIPT%"
    pause
    exit /b 1
)
powershell.exe -NoLogo -NoProfile -STA -Command "$ErrorActionPreference='Stop'; $scriptPath=$env:TARGET_SCRIPT; $arguments='-NoLogo -NoProfile -STA -File ' + [char]34 + $scriptPath + [char]34; if ($env:WINDOW_STYLE) { Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList $arguments -WorkingDirectory (Split-Path -Parent $scriptPath) -Verb RunAs -WindowStyle Hidden } else { Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList $arguments -WorkingDirectory (Split-Path -Parent $scriptPath) -Verb RunAs }"
if errorlevel 1 (
    echo Unable to start the selected mode. Check the error above.
    pause
    exit /b 1
)

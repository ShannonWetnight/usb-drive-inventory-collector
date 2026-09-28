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
set "CHOICE="
set /p "CHOICE=Choose mode [G/C/Q]: "
if /i "%CHOICE%"=="Q" exit /b 0
set "COLLECTOR_MODE=GUI"
if /i "%CHOICE%"=="C" set "COLLECTOR_MODE=CLI"
if not "%CHOICE%"=="" if /i not "%CHOICE%"=="G" if /i not "%CHOICE%"=="C" (
    echo Choose G, C, or Q.
    pause
    exit /b 1
)

if /i "%COLLECTOR_MODE%"=="GUI" (
    powershell.exe -NoLogo -NoProfile -STA -NoExit -File "%COLLECTOR_SCRIPT%" -Mode GUI -HideConsoleOnGuiReady -ElevateOnStartup
) else (
    powershell.exe -NoLogo -NoProfile -STA -File "%COLLECTOR_SCRIPT%" -Mode CLI -ElevateOnStartup
)
if errorlevel 1 (
    echo Unable to launch the collector. Check the error above.
    pause
    exit /b 1
)

@echo off
setlocal
set "GUI_SCRIPT=%~dp0USB-Drive-Inventory-Collector-GUI.ps1"
if not exist "%GUI_SCRIPT%" (
    echo The GUI script was not found next to this launcher:
    echo "%GUI_SCRIPT%"
    pause
    exit /b 1
)
powershell.exe -NoLogo -NoProfile -STA -Command "$ErrorActionPreference='Stop'; $scriptPath=$env:GUI_SCRIPT; $arguments='-NoLogo -NoProfile -STA -File ' + [char]34 + $scriptPath + [char]34; Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList $arguments -WorkingDirectory (Split-Path -Parent $scriptPath) -Verb RunAs"
if errorlevel 1 (
    echo Unable to start the GUI. Check the error above.
    pause
    exit /b 1
)

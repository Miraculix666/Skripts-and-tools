@echo off
title ?? opsi WinPE Builder - ADK Auto
echo [1/4] Suche ADK Deployment Shell...
echo.

REM ADK Shell finden (verschiedene Versionen)
set ADK_CMD1="%ProgramFiles(x86)%\Windows Kits\10\Assessment and Deployment Kit\Deployment and Imaging Tools\adkwinpesdkdeploy.cmd"
set ADK_CMD2="%ProgramFiles(x86)%\Windows Kits\10\Assessment and Deployment Kit\Deployment and Imaging Tools\DandISetEnv.bat"
set ADK_CMD3="%ProgramW6432%\Windows Kits\10\Assessment and Deployment Kit\Deployment and Imaging Tools\adkwinpesdkdeploy.cmd"

if exist %ADK_CMD1% (
    echo [2/4] Starte ADK Shell: %ADK_CMD1%
    call %ADK_CMD1% && goto :run_ps
)
if exist %ADK_CMD2% (
    echo [2/4] Starte ADK Shell: %ADK_CMD2%
    call %ADK_CMD2% && goto :run_ps
)
if exist %ADK_CMD3% (
    echo [2/4] Starte ADK Shell: %ADK_CMD3%
    call %ADK_CMD3% && goto :run_ps
)

echo ? ADK Deployment Tools NICHT GEFUNDEN!
echo.
echo 1. Lade ADK: https://go.microsoft.com/fwlink/?linkid=2269546
echo 2. Lade WinPE Add-on: https://go.microsoft.com/fwlink/?linkid=2269547
echo 3. Installiere "Deployment Tools" + "WinPE"
pause
exit /b 1

:run_ps
echo [3/4] F?hre WinPE Builder aus...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0WinPE_maker_Per.ps1" -BootWIM "H:\sources\boot.wim" -opsiwinpepath "c:\temp\opsiPE_per\win11-x64"
echo [4/4] WinPE bereit: c:\temp\opsiPE_per\win11-x64\winpe\
pause

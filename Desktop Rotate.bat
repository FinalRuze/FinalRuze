@echo off

:loop
powershell -Command "(Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\GraphicsDrivers\Configuration\*\00\00\*\00\00\00\00').Rotation = 1"
timeout /t 5
powershell -Command "(Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\GraphicsDrivers\Configuration\*\00\00\*\00\00\00\00').Rotation = 0"
timeout /t 5
goto loop

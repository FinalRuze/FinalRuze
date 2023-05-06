@echo off

:loop

rem Set the coordinates range of the desktop
for /f "tokens=2 delims=:." %%w in ('powershell "(Get-ItemProperty -Path 'HKCU:\Control Panel\Desktop\WindowMetrics' -Name AppliedDPI).AppliedDPI"') do set "dpi=%%w"
set /a "xMin=0, xMax=%%~fswidth-100, yMin=0, yMax=%%~fsheight-100"
set /a "xMin=!xMin!*96/!dpi!, xMax=!xMax!*96/!dpi!, yMin=!yMin!*96/!dpi!, yMax=!yMax!*96/!dpi!"

rem Loop through all the files on the desktop
for %%f in (%userprofile%\Desktop\*.lnk) do (
  rem Get a random position on the desktop
  set /a "x=!random! * (%xMax% - %xMin% + 1) / 32768 + %xMin%"
  set /a "y=!random! * (%yMax% - %yMin% + 1) / 32768 + %yMin%"

  rem Move the file to the random position
  %SystemRoot%\System32\cmd.exe /c "start /b /min powershell -Command \"(New-Object -ComObject 'Shell.Application').Namespace('%userprofile%\Desktop').ParseName('%~nxf').InvokeVerb('move')\""
  %SystemRoot%\System32\cmd.exe /c "start /b /min powershell -Command \"(New-Object -ComObject 'Shell.Application').Namespace('%userprofile%\Desktop').MoveHere((New-Object -ComObject 'Shell.Application').Namespace('%userprofile%\Desktop').ParseName('%~nxf').Self.Path), 0x10 + 0x40 + 0x200 + ($x -as [int]) * 0x10000 + ($y -as [int]))\""
)

rem Wait for 2 seconds before repeating the loop
ping -n 2 127.0.0.1 > nul
goto loop

@echo off
:loop
echo Set oShell = CreateObject("WScript.Shell") > %temp%\temp.vbs
echo oShell.Run "PowerShell -Command ""[System.Media.SystemSounds]::Beep.Play()"", 0, True >> %temp%\temp.vbs
cscript /nologo %temp%\temp.vbs
del %temp%\temp.vbs
goto loop

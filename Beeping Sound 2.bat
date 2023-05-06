@echo off
:loop
PowerShell -Command "Start-Process -FilePath 'C:\Windows\Media\chimes.wav' -WindowStyle Hidden"
goto loop

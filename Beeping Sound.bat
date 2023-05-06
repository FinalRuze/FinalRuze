@echo off
:loop
PowerShell -Command "[System.Media.SystemSounds]::Beep.Play()"
goto loop

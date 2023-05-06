@echo off
:loop
PowerShell -Command "Add-Type -AssemblyName PresentationCore; [System.Media.SystemSounds]::Beep.Play(); $null = [System.Windows.MessageBox]::Show('Beep!', '', 'OK', 'None', 'DefaultButton1', 'MessageBoxIcon.Information', 'SystemModal')"
goto loop

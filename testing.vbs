Set objShell = CreateObject("Shell.Application")
Set objWindows = objShell.Windows

Do While True
    For Each objWindow In objWindows
        If objWindow.Document.Folder.Self.Path = "C:\Users\<username>\Desktop" Then
            Randomize
            xPos = Int((Screen.Width - objWindow.Width) * Rnd)
            yPos = Int((Screen.Height - objWindow.Height) * Rnd)
            objWindow.Left = xPos
            objWindow.Top = yPos
        End If
    Next
    CreateObject("WScript.Shell").Run "PowerShell -Command ""[System.Media.SystemSounds]::Beep.Play()""", 0, True
Loop

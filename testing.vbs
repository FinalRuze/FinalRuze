Do While True
    ' Play a beep sound
    CreateObject("WScript.Shell").Run "PowerShell -Command ""[System.Media.SystemSounds]::Beep.Play()""", 0, True
    
    ' Play a hand sound
    CreateObject("WScript.Shell").Run "PowerShell -Command ""[System.Media.SystemSounds]::Hand.Play()""", 0, True
    
    ' Play an exclamation sound
    CreateObject("WScript.Shell").Run "PowerShell -Command ""[System.Media.SystemSounds]::Exclamation.Play()""", 0, True
    
    ' Play an asterisk sound
    CreateObject("WScript.Shell").Run "PowerShell -Command ""[System.Media.SystemSounds]::Asterisk.Play()""", 0, True
    
    ' Play a question sound
    CreateObject("WScript.Shell").Run "PowerShell -Command ""[System.Media.SystemSounds]::Question.Play()""", 0, True
    
Loop

Do While True
    CreateObject("WScript.Shell").Run "PowerShell -Command ""[System.Media.SystemSounds]::Beep.Play()""", 0, True
Loop

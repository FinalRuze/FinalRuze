Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

posX = Int((objShell.AppActivate("Program Manager"))/2)
posY = Int((objShell.AppActivate("Program Manager"))/2)

Do While True
    posX = Int((objShell.AppActivate("Program Manager"))/2)
    posY = Int((objShell.AppActivate("Program Manager"))/2)
    MsgBox "Error Message", 16, "Error"
    Set objShell = CreateObject("WScript.Shell")
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colMonitors = objWMIService.ExecQuery("Select * from Win32_DesktopMonitor")
    For Each objMonitor in colMonitors
        screenWidth = objMonitor.ScreenWidth
        screenHeight = objMonitor.ScreenHeight
    Next
    Randomize
    xPos = Int((screenWidth - 250) * Rnd + 1)
    yPos = Int((screenHeight - 150) * Rnd + 1)
    Set WshShell = WScript.CreateObject("WScript.Shell")
    WshShell.Run "mshta.exe vbscript:msgbox(""" & "This error cannot be closed." & """,16,""Error""):close"
    Do Until WshShell.AppActivate("Error") : WScript.Sleep 100 : Loop
    WScript.Sleep 500
    WshShell.SendKeys "{ENTER}"
    Do Until WshShell.AppActivate("Error") : WScript.Sleep 100 : Loop
    WshShell.SendKeys "{TAB}"
    WshShell.SendKeys "{TAB}"
    WshShell.SendKeys "{TAB}"
    WshShell.SendKeys "{TAB}"
    WshShell.SendKeys "{ENTER}"
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colMonitors = objWMIService.ExecQuery("Select * from Win32_DesktopMonitor")
    For Each objMonitor in colMonitors
        screenWidth = objMonitor.ScreenWidth
        screenHeight = objMonitor.ScreenHeight
    Next
    Randomize
    xPos = Int((screenWidth - 250) * Rnd + 1)
    yPos = Int((screenHeight - 150) * Rnd + 1)
    Do Until (WshShell.AppActivate("Error") Or WshShell.AppActivate("Task Switching"))
        WshShell.MoveTo xPos, yPos
        WScript.Sleep 100
    Loop
    WshShell.SendKeys "{ESC}"
    WshShell.SendKeys "{ENTER}"
Loop

Set objShell = WScript.CreateObject("WScript.Shell")

' Create first icon
strDesktopPath = objShell.SpecialFolders("Desktop")
Set objShortcut = objShell.CreateShortcut(strDesktopPath & "\Icon1.lnk")
objShortcut.TargetPath = "C:\Windows\System32\notepad.exe"
objShortcut.Save
objShortcut.Move 100, 100

' Create second icon
Set objShortcut = objShell.CreateShortcut(strDesktopPath & "\Icon2.lnk")
objShortcut.TargetPath = "C:\Windows\System32\calc.exe"
objShortcut.Save
objShortcut.Move 200, 200

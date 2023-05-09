Const DesktopFolder = &H10&

Set objShell = CreateObject("Shell.Application")

' Create first error icon
Set objFolder = objShell.Namespace(DesktopFolder)
Set objLink = objFolder.Self.CreateShortcut(objFolder.Self.Path & "\Error1.lnk")
objLink.TargetPath = "C:\NonExistentPath\NonExistentFile1.exe"
objLink.Save

' Create second error icon
Set objLink = objFolder.Self.CreateShortcut(objFolder.Self.Path & "\Error2.lnk")
objLink.TargetPath = "D:\NonExistentPath\NonExistentFile2.exe"
objLink.Save

' Move icons to specific locations on the screen
Set objShellWindows = CreateObject("Shell.Application").Windows
Set objShellWindow = objShellWindows.Item(objShellWindows.Count - 1)
objShellWindow.Left = 100
objShellWindow.Top = 100

Set objShellWindow = objShellWindows.Item(objShellWindows.Count - 2)
objShellWindow.Left = 200
objShellWindow.Top = 200

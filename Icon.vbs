Const DesktopFolder = &H10&

Set objShell = CreateObject("Shell.Application")

' Create first error icon
Set objFolder = objShell.Namespace(DesktopFolder)
Set objLink = objFolder.Items().InvokeVerb("New Shortcut")
Set objWshShell = CreateObject("WScript.Shell")
objWshShell.SendKeys "{TAB}{TAB}{TAB}{TAB}C:\NonExistentPath\NonExistentFile1.exe{ENTER}"
objLink.InvokeVerb("SaveShortcut")

' Create second error icon
Set objLink = objFolder.Items().InvokeVerb("New Shortcut")
objWshShell.SendKeys "{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}D:\NonExistentPath\NonExistentFile2.exe{ENTER}"
objLink.InvokeVerb("SaveShortcut")

' Move icons to specific locations on the screen
Set objShellWindows = CreateObject("Shell.Application").Windows
Set objShellWindow = objShellWindows.Item(objShellWindows.Count - 1)
objShellWindow.Left = 100
objShellWindow.Top = 100

Set objShellWindow = objShellWindows.Item(objShellWindows.Count - 2)
objShellWindow.Left = 200
objShellWindow.Top = 200

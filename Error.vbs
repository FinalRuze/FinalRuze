Set WshShell = WScript.CreateObject("WScript.Shell")
Set IE = CreateObject("InternetExplorer.Application")

IE.Navigate("about:blank")
IE.ToolBar = 0
IE.MenuBar = 0
IE.StatusBar = 0
IE.Resizable = 0
IE.Width = 400
IE.Height = 200
IE.Top = Int((WshShell.ScreenHeight - IE.Height) * Rnd)
IE.Left = Int((WshShell.ScreenWidth - IE.Width) * Rnd)
IE.Document.Title = "Error"
IE.Document.Body.InnerHTML = "<center><h1>Error</h1><p>An error has occurred.</p></center>"
IE.Visible = 1

Do While IE.Busy
    WScript.Sleep 100
Loop

Do While IE.Visible
    Set window = WshShell.Windows.Item(IE.Document.Title)
    If Not IsNull(window) Then
        Dim closeBtn : closeBtn = False
        For Each btn In IE.Document.getElementsByTagName("button")
            btn.Onmouseover = "MoveError()"
            btn.Onmouseout = "ResetError()"
            If btn.getAttribute("type") = "button" Then
                btn.OnClick = "closeBtn = True"
            End If
        Next
        
        While Not closeBtn
            Randomize
            IE.Top = Int((WshShell.ScreenHeight - IE.Height) * Rnd)
            IE.Left = Int((WshShell.ScreenWidth - IE.Width) * Rnd)
            WScript.Sleep 500
            If closeBtn Then Exit While
        Wend
        IE.Quit
    End If
    WScript.Sleep 100
Loop

Sub MoveError()
    window.MoveTo Int((WshShell.ScreenWidth - IE.Width) * Rnd), Int((WshShell.ScreenHeight - IE.Height) * Rnd)
End Sub

Sub ResetError()
    window.MoveTo IE.Left, IE.Top
End Sub

Set objShell = CreateObject("WScript.Shell")

'Get the dimensions of the screen
ScreenWidth = objShell.Screen.Width
ScreenHeight = objShell.Screen.Height

'Create the message box
Do
    Set objMsgBox = CreateObject("WScript.Shell")
    xPos = Int((ScreenWidth - 400) * Rnd + 1)
    yPos = Int((ScreenHeight - 150) * Rnd + 1)
    objMsgBox.Popup "An Error has occurred!", 0, "Error", 0x10 + 0x1 + 0x100, xPos, yPos
    
    'Check if the mouse is over the Close button
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Navigate "about:blank"
    objIE.ToolBar = 0
    objIE.StatusBar = 0
    objIE.Width = 400
    objIE.Height = 150
    objIE.Left = xPos
    objIE.Top = yPos
    objIE.Document.Body.InnerHTML = "<h1>An Error has occurred!</h1>"
    objIE.Visible = True
    
    Do While objIE.Busy
        WScript.Sleep 100
    Loop
    
    Set objDoc = objIE.Document
    Set objCloseButton = objDoc.ParentWindow.document.getElementById("closebutton")
    If Not objCloseButton Is Nothing Then
        objCloseButton.onmouseover = GetRef("CloseHover")
    End If
    
    'Wait for a few seconds
    WScript.Sleep 5000
Loop

Sub CloseHover()
    xPos = Int((ScreenWidth - 400) * Rnd + 1)
    yPos = Int((ScreenHeight - 150) * Rnd + 1)
    objIE.MoveTo xPos, yPos
End Sub

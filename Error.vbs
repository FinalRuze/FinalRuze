Set objShell = CreateObject("WScript.Shell")
Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")

' Get screen dimensions
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController", , 48)
For Each objItem in colItems
    intScreenHeight = objItem.CurrentVerticalResolution
    intScreenWidth = objItem.CurrentHorizontalResolution
Next

' Create error message box
Set objError = objShell.Popup("This is an error message.", 0, "Error", 0 + 16)

' Move message box to random position
If objError = 1 Then
    Set objWshShell = CreateObject("WScript.Shell")
    Do
        intLeft = Int((intScreenWidth - objError.width) * Rnd + 1)
        intTop = Int((intScreenHeight - objError.height) * Rnd + 1)
        objWshShell.AppActivate "Error"
        objWshShell.MoveTo intLeft, intTop
    Loop Until intLeft > (intScreenWidth - 100) Or intTop > (intScreenHeight - 100)
End If

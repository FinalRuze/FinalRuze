Option Explicit

Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")

Dim x, y
Randomize
x = Int(objShell.AppActivate(WScript.ScriptFullName) * Rnd)
y = Int(objShell.AppActivate(WScript.ScriptFullName) * Rnd)

Do Until False
    Dim errorMessage
    errorMessage = "An error has occurred."

    Set objError = objShell.Popup(errorMessage, 0, "Error", vbCritical + vbSystemModal + vbOKCancel, vbDefaultButton2, x, y)

    If objError = vbCancel Then
        WScript.Quit
    End If

    x = Int(objShell.AppActivate(WScript.ScriptFullName) * Rnd)
    y = Int(objShell.AppActivate(WScript.ScriptFullName) * Rnd)
Loop

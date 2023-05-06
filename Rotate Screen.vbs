' Rotate the screen to the left by default upon execution
WshShell.SendKeys("{LEFT}")
currentOrientation = DMDO_270

' Loop to continuously rotate the screen
Do While True
    ' Rotate the screen to the right after 5 seconds
    WScript.Sleep 5000
    WshShell.SendKeys("{RIGHT}")
    If currentOrientation = DMDO_DEFAULT Then
        currentOrientation = DMDO_90
    ElseIf currentOrientation = DMDO_90 Then
        currentOrientation = DMDO_180
    ElseIf currentOrientation = DMDO_180 Then
        currentOrientation = DMDO_270
    ElseIf currentOrientation = DMDO_270 Then
        currentOrientation = DMDO_DEFAULT
    End If
Loop

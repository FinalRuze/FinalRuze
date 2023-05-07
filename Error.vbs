Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")

' Loop indefinitely
Do While True
    ' Get random X and Y coordinates
    Randomize
    Dim xCoord
    Dim yCoord
    xCoord = Int((1910 * Rnd) + 1)
    yCoord = Int((1070 * Rnd) + 1)

    ' Display the message box at the random position
    objShell.Popup "Error: Your computer has encountered a critical error and needs to restart." & vbNewLine & vbNewLine & "Contact your system administrator for assistance.", 0, "Critical Error", 16 + 4096, xCoord, yCoord
Loop

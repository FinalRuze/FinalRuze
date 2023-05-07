Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Modify these values to change the number of error messages and the message text
numErrors = 10
errorMsg = "Error: Something went wrong!"

' Loop through the number of errors and display messages in random positions on the screen
For i = 1 to numErrors
    ' Calculate random x and y positions for the message
    xPos = Int((objShell.AppActivate("Program Manager").ScreenWidth - 400) * Rnd + 1)
    yPos = Int((objShell.AppActivate("Program Manager").ScreenHeight - 100) * Rnd + 1)

    ' Display the message box
    objShell.Popup errorMsg, 5, "Error " & i, 16 + 0 + 4096, xPos, yPos
Next

Set objShell = CreateObject("WScript.Shell")
Set objRandom = CreateObject("System.Random")

For i = 1 to 3
    ' Create a new cursor with a different ID for each instance
    cursorID = objShell.CreateCursor(i)

    ' Set the position of the cursor using SetCursorPos with random values for x and y
    xPos = objRandom.Next(0, 1024) ' Set a random value between 0 and 1024 for the x coordinate
    yPos = objRandom.Next(0, 768) ' Set a random value between 0 and 768 for the y coordinate
    objShell.SetCursorPos xPos, yPos, cursorID
Next

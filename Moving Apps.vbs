Set objShell = WScript.CreateObject("WScript.Shell")
Set objShellApp = CreateObject("Shell.Application")
Set objDesktop = objShellApp.Namespace(0)
Set objFolder = objDesktop.Self

Do While True
    For Each objItem In objFolder.Items
            'get random x and y coordinates
            Randomize
            xPos = Int((Screen.Width - objItem.Width) * Rnd)
            yPos = Int((Screen.Height - objItem.Height) * Rnd)
            
            'move the item to the random position
            objShellApp.MoveItem objItem, xPos, yPos
        End If
    Next
    WScript.Sleep 5000 'wait for 5 seconds before moving items again
Loop

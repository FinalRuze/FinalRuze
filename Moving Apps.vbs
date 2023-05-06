Option Explicit
Dim objShell, objFolder, objItem, i

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(16)

Do While True
    For i = 0 to objFolder.Items.Count - 1
        Set objItem = objFolder.Items.Item(i)
        Randomize
        objItem.InvokeVerb("Move...")
        WScript.Sleep 1000 'wait for the "Move Items" dialog to appear
        objShell.SendKeys "{TAB}{TAB}"
        objShell.SendKeys "{DOWN " & Int((Screen.Height - objItem.Height) * Rnd) & "}"
        objShell.SendKeys "{RIGHT " & Int((Screen.Width - objItem.Width) * Rnd) & "}"
        objShell.SendKeys "{ENTER}"
    Next
    WScript.Sleep 5000 'wait for 5 seconds before moving items again
Loop

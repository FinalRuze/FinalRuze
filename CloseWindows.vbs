Set objShell = CreateObject("Shell.Application")
For Each win in objShell.Windows
    win.Quit()
Next

Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

Do While True
    Set objFolder = objFSO.GetFolder("%userprofile%")
    Set colFiles = objFolder.Files
    For Each objFile in colFiles
        If objFSO.GetExtensionName(objFile.Name) = "txt" Then
            x = Int((objShell.AppActivate("%random% - Notepad").Height - 50 + 1) * Rnd + 50)
            y = Int((objShell.AppActivate("%random% - Notepad").Width - 50 + 1) * Rnd + 50)
            objShell.AppActivate "%random% - Notepad"
            objShell.MoveTo x, y
        End If
    Next
    WScript.Sleep 1000
Loop

Option Explicit
Dim objShell, objFolder, objItem, i, desktopHeight, desktopWidth

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(16)

desktopHeight = GetSystemMetrics(1)
desktopWidth = GetSystemMetrics(0)

Do While True
    For i = 0 to objFolder.Items.Count - 1
        Set objItem = objFolder.Items.Item(i)
        Randomize
        objItem.InvokeVerb("Move...")
        WScript.Sleep 1000 'wait for the "Move Items" dialog to appear
        MoveItemDialog desktopWidth - objItem.Width, desktopHeight - objItem.Height
        objShell.Namespace(0).Self.Refresh 'refresh the desktop to show the moved items
    Next
    WScript.Sleep 5000 'wait for 5 seconds before moving items again
Loop

Sub MoveItemDialog(maxLeft, maxTop)
    Dim hwnd, itemDialog, okButton, leftPos, topPos
    hwnd = FindWindowEx(0, 0, "Dialog", "Move Items")
    Do While hwnd = 0
        hwnd = FindWindowEx(0, 0, "Dialog", "Move Items")
        WScript.Sleep 100
    Loop
    Set itemDialog = CreateObject("WScript.Shell").AppActivate("Move Items")
    WScript.Sleep 1000 'wait for the "Move Items" dialog to fully load
    leftPos = Int(Rnd * maxLeft)
    topPos = Int(Rnd * maxTop)
    Set okButton = FindWindowEx(hwnd, 0, "Button", "&Move")
    SendMessage hwnd, &H84, 0, MakeLParam(okButton.Left + 10, okButton.Top + 10)
    SendMessage hwnd, &HF5, 0, 0
    Set okButton = Nothing
    Set itemDialog = Nothing
End Sub

Function MakeLParam(wLow, wHigh)
    MakeLParam = wLow Or (wHigh * &H10000)
End Function

Function FindWindowEx(parentHandle, childAfter, className, windowTitle)
    Dim handle
    handle = FindWindowExA(parentHandle, childAfter, StrPtr(className), StrPtr(windowTitle))
    If handle = 0 Then
        handle = FindWindowExW(parentHandle, childAfter, StrPtr(className), StrPtr(windowTitle))
    End If
    FindWindowEx = handle
End Function

Function FindWindowExA(parentHandle, childAfter, className, windowTitle)
    Dim buffer, handle
    buffer = String(256, Chr(0))
    handle = FindWindowExW(parentHandle, childAfter, className, windowTitle)
    Do While handle <> 0
        GetClassNameA handle, buffer, Len(buffer)
        If StrComp(Left(buffer, InStr(buffer, Chr(0)) - 1), className, vbTextCompare) = 0 Then
            FindWindowExA = handle
            Exit Function
        End If
        handle = FindWindowExW(parentHandle, handle, className, windowTitle)
    Loop
    FindWindowExA = handle
End Function

Function FindWindowExW(parentHandle, childAfter, className, windowTitle)
    FindWindowExW = FindWindowExByString(parentHandle, childAfter, StrPtr(className), StrPtr(windowTitle))
End Function

Function FindWindowExByString(parentHandle, childAfter, className

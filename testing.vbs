Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_DesktopMonitor")

For Each objItem in colItems
    ' Check if SetDMPSupported is available
    If InStr(1, objItem.Qualifiers_.Item("SupportedMethods").Value, "SetDMPSupported") > 0 Then
        objItem.SetDMPSupported(False)
    End If

    objItem.SetCurrentHorizontalResolution(1024)
    objItem.SetCurrentVerticalResolution(768)
Next

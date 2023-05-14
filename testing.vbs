Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_DesktopMonitor")

For Each objItem in colItems
    objItem.SetDMPSupported(False)
    objItem.SetCurrentHorizontalResolution(1024)
    objItem.SetCurrentVerticalResolution(768)
Next


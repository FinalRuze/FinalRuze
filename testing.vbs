Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_DisplayConfiguration")

For Each objItem in colItems
    objItem.SetPanning(0,0)
    objItem.SetDisplayConfig 1024, 768, 0
Next

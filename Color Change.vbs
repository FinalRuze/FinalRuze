Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")

Dim colors(4)
colors(0) = "0000FF" ' Blue
colors(1) = "00FF00" ' Green
colors(2) = "FF0000" ' Red
colors(3) = "FFFF00" ' Yellow

Do While True
    ' Select a random color
    Randomize
    index = Int((UBound(colors) + 1) * Rnd)
    color = colors(index)
    
    ' Change the background color
    WshShell.RegWrite "HKCU\Control Panel\Colors\Background", "0 " & color, "REG_SZ"
    
    ' Wait for 1 second
    WScript.Sleep 1000
Loop

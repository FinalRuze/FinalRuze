Set objShell = CreateObject("WScript.Shell")
Do While True
    Randomize
    R = Int((255-0+1)*Rnd+0)
    G = Int((255-0+1)*Rnd+0)
    B = Int((255-0+1)*Rnd+0)
    objShell.RegWrite "HKCU\Control Panel\Colors\Background", _
    "0 " & R & " " & G & " " & B
    objShell.Run "RUNDLL32.EXE USER32.DLL,UpdatePerUserSystemParameters", 1, True
    WScript.Sleep 1000 ' Change the sleep time to adjust the frequency of color changes
Loop

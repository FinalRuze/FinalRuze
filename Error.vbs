Set objShell = CreateObject("WScript.Shell")

While True
    randomX = Int((1900-100+1)*Rnd+100)
    randomY = Int((1000-100+1)*Rnd+100)
    objShell.Run "mshta.exe vbscript:msgbox(""ERROR: Your computer has been hacked!"",48,""Windows Security"")(window.resizeTo(600,500));window.moveTo " & randomX & "," & randomY & ";close", 0, False
    WScript.Sleep 5000
Wend

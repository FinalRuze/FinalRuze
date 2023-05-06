Set WshShell = CreateObject("WScript.Shell")

Do
  Randomize
  Direction = Int((4 * Rnd) + 1)
  Select Case Direction
    Case 1
      WshShell.SendKeys "{UP}"
    Case 2
      WshShell.SendKeys "{DOWN}"
    Case 3
      WshShell.SendKeys "{LEFT}"
    Case 4
      WshShell.SendKeys "{RIGHT}"
  End Select
  WScript.Sleep(1000) ' Pause for 1 second before moving cursor again
Loop

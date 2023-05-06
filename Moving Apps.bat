@echo off
setlocal EnableDelayedExpansion

set DPI=96
set W=1920
set H=1080

set /a "x_min=0", "x_max=(%W%*%DPI%/96)-150"
set /a "y_min=0", "y_max=(%H%*%DPI%/96)-150"

:loop
for /f "delims=" %%a in ('dir /b /a "C:\Users\%username%\Desktop\*.lnk"') do (
    set /a "x=!RANDOM! * (%x_max% - %x_min% + 1) / 32768 + %x_min%",
    set /a "y=!RANDOM! * (%y_max% - %y_min% + 1) / 32768 + %y_min%"
    powershell -command "& { (New-Object -ComObject WScript.Shell).CreateShortcut('C:\Users\%username%\Desktop\%%a').Move($x, $y) }"
)
timeout /t 2 /nobreak >nul
goto loop

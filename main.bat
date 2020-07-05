@echo off

set /p code=How many letters we trying to crack:

goto pre


:pre

  echo Set WshShell = WScript.CreateObject("WScript.Shell")> code.vbs
  echo Set objShell = CreateObject("WScript.Shell")>> code.vbs
  echo var1 = 1 >> code.vbs
  echo WScript.Sleep 5000>> code.vbs
  echo.>> code.vbs
  echo Do>> code.vbs
  echo WScript.Sleep 1000>> code.vbs
  echo WshShell.SendKeys "admin">> code.vbs
  echo.>> code.vbs
  echo With CreateObject("WScript.Shell")>> code.vbs
  echo     .Run "nircmd setcursor 761 334‬", 0, True>> code.vbs
  echo     .Run "nircmd sendmouse left click", 0, True>> code.vbs
  echo End With>> code.vbs
  echo WScript.Sleep 500>> code.vbs
echo.>> code.vbs
goto start

:start
set x=-1
set /A result=%x% + %code%
(
   echo var%code% = var%code% + "1"
   echo If var%code% = 27 Then
   echo    var%code% = 1 
   echo code = %result%
   echo var%code% = var%code% + 1
) >> code.vbs 
set x=1
set /A result=%x% + %code%
(
   echo code = %result%
   echo End If
  echo If var%code% = "1" Then
echo WshShell.SendKeys "a"
echo ElseIf var%code% = "2" Then
echo WshShell.SendKeys "b"
echo ElseIf var%code% = "3" Then
echo WshShell.SendKeys "c"
echo ElseIf var%code% = "4" Then
echo WshShell.SendKeys "d"
echo ElseIf var%code% = "5" Then
echo WshShell.SendKeys "e"
echo ElseIf var%code% = "6" Then
echo WshShell.SendKeys "f"
echo ElseIf var%code% = "7" Then
echo WshShell.SendKeys "g"
echo ElseIf var%code% = "8" Then
echo WshShell.SendKeys "h"
echo ElseIf var%code% = "9" Then
echo WshShell.SendKeys "i"
echo ElseIf var%code% = "10" Then
echo WshShell.SendKeys "j"
echo ElseIf var%code% = "11" Then
echo WshShell.SendKeys "k"
echo ElseIf var%code% = "12" Then
echo WshShell.SendKeys "l"
echo ElseIf var%code% = "13" Then
echo WshShell.SendKeys "m"
echo ElseIf var%code% = "14" Then
echo WshShell.SendKeys "n"
echo ElseIf var%code% = "15" Then
echo WshShell.SendKeys "o"
echo ElseIf var%code% = "16" Then
echo WshShell.SendKeys "p"
echo ElseIf var%code% = "17" Then
echo WshShell.SendKeys "q"
echo ElseIf var%code% = "18" Then
echo WshShell.SendKeys "r"
echo ElseIf var%code% = "19" Then
echo WshShell.SendKeys "s"
echo ElseIf var%code% = "20" Then
echo WshShell.SendKeys "t"
echo ElseIf var%code% = "21" Then
echo WshShell.SendKeys "u"
echo ElseIf var%code% = "22" Then
echo WshShell.SendKeys "v"
echo ElseIf var%code% = "23" Then
echo WshShell.SendKeys "w"
echo ElseIf var%code% = "24" Then
echo WshShell.SendKeys "x"
echo ElseIf var%code% = "25" Then
echo WshShell.SendKeys "y"
echo ElseIf var%code% = "26" Then
echo WshShell.SendKeys "z"
echo End If
  echo.
) >> code.vbs

set x=-1
set /A result=%x% + %code%

set code=%result%
if %code%==0 (goto final) else (goto start3)


:start3
set x=-1
set /A result=%x% + %code%
(
   echo var%code% = var%code% + "1"
   echo If var%code% = 27 Then
   echo    var%code% = 1 
   echo code = %result%
   echo var%code% = var%code% + 1
) >> code.vbs 
set x=1
set /A result=%x% + %code%
(
   echo code = %result%
   echo End If
  echo If var%code% = "1" Then
echo WshShell.SendKeys "a"
echo ElseIf var%code% = "2" Then
echo WshShell.SendKeys "b"
echo ElseIf var%code% = "3" Then
echo WshShell.SendKeys "c"
echo ElseIf var%code% = "4" Then
echo WshShell.SendKeys "d"
echo ElseIf var%code% = "5" Then
echo WshShell.SendKeys "e"
echo ElseIf var%code% = "6" Then
echo WshShell.SendKeys "f"
echo ElseIf var%code% = "7" Then
echo WshShell.SendKeys "g"
echo ElseIf var%code% = "8" Then
echo WshShell.SendKeys "h"
echo ElseIf var%code% = "9" Then
echo WshShell.SendKeys "i"
echo ElseIf var%code% = "10" Then
echo WshShell.SendKeys "j"
echo ElseIf var%code% = "11" Then
echo WshShell.SendKeys "k"
echo ElseIf var%code% = "12" Then
echo WshShell.SendKeys "l"
echo ElseIf var%code% = "13" Then
echo WshShell.SendKeys "m"
echo ElseIf var%code% = "14" Then
echo WshShell.SendKeys "n"
echo ElseIf var%code% = "15" Then
echo WshShell.SendKeys "o"
echo ElseIf var%code% = "16" Then
echo WshShell.SendKeys "p"
echo ElseIf var%code% = "17" Then
echo WshShell.SendKeys "q"
echo ElseIf var%code% = "18" Then
echo WshShell.SendKeys "r"
echo ElseIf var%code% = "19" Then
echo WshShell.SendKeys "s"
echo ElseIf var%code% = "20" Then
echo WshShell.SendKeys "t"
echo ElseIf var%code% = "21" Then
echo WshShell.SendKeys "u"
echo ElseIf var%code% = "22" Then
echo WshShell.SendKeys "v"
echo ElseIf var%code% = "23" Then
echo WshShell.SendKeys "w"
echo ElseIf var%code% = "24" Then
echo WshShell.SendKeys "x"
echo ElseIf var%code% = "25" Then
echo WshShell.SendKeys "y"
echo ElseIf var%code% = "26" Then
echo WshShell.SendKeys "z"
echo code=%code% - "1"
echo var%code%= var%code% + "1"
echo code=%code% + "1"
echo End If
  echo.
) >> code.vbs

set x=-1
set /A result=%x% + %code%

set code=%result%
goto start2


:start2
if %code%==0 (goto final) else (goto start)

:final
(
echo WshShell.SendKeys "{ENTER}"
echo var1 = var1 - 1
echo loop
) >> code.vbs
timeout /t 1 >nul
start code.vbs

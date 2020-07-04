Set WshShell = WScript.CreateObject("WScript.Shell")
Set objShell = CreateObject("WScript.Shell")
first = 1
WScript.Sleep 5000

Do
WScript.Sleep 1000
WshShell.SendKeys "admin"
If first = 27 Then
   first = 1
End If

With CreateObject("WScript.Shell")
    .Run "nircmd setcursor 761 334‬", 0, True
    .Run "nircmd sendmouse left click", 0, True
End With
WScript.Sleep 500

If first = "1" Then
WshShell.SendKeys "a"
ElseIf first = "2" Then
WshShell.SendKeys "b"
ElseIf first = "3" Then
WshShell.SendKeys "c"
ElseIf first = "4" Then
WshShell.SendKeys "d"
ElseIf first = "5" Then
WshShell.SendKeys "e"
ElseIf first = "6" Then
WshShell.SendKeys "f"
ElseIf first = "7" Then
WshShell.SendKeys "g"
ElseIf first = "8" Then
WshShell.SendKeys "h"
ElseIf first = "9" Then
WshShell.SendKeys "i"
ElseIf first = "10" Then
WshShell.SendKeys "j"
ElseIf first = "11" Then
WshShell.SendKeys "k"
ElseIf first = "12" Then
WshShell.SendKeys "l"
ElseIf first = "13" Then
WshShell.SendKeys "m"
ElseIf first = "14" Then
WshShell.SendKeys "n"
ElseIf first = "15" Then
WshShell.SendKeys "o"
ElseIf first = "16" Then
WshShell.SendKeys "p"
ElseIf first = "17" Then
WshShell.SendKeys "q"
ElseIf first = "18" Then
WshShell.SendKeys "r"
ElseIf first = "19" Then
WshShell.SendKeys "s"
ElseIf first = "20" Then
WshShell.SendKeys "t"
ElseIf first = "21" Then
WshShell.SendKeys "u"
ElseIf first = "22" Then
WshShell.SendKeys "v"
ElseIf first = "23" Then
WshShell.SendKeys "w"
ElseIf first = "24" Then
WshShell.SendKeys "x"
ElseIf first = "25" Then
WshShell.SendKeys "y"
ElseIf first = "26" Then
WshShell.SendKeys "z"
End If

WshShell.SendKeys "{ENTER}"
first = first + 1
loop

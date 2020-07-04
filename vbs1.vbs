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
End If

WshShell.SendKeys "{ENTER}"
first = first + 1
loop

Set WshShell = WScript.CreateObject("WScript.Shell")
Set objShell = CreateObject("WScript.Shell")
var1 = 1
WScript.Sleep 5000

Do
WScript.Sleep 1000
WshShell.SendKeys "admin"
If var1 = 27 Then
   var1 = 1
var2 = var2 + "1"
var3 = var3 + "1"
var4 = var4 + "1"
End If

With CreateObject("WScript.Shell")
    .Run "nircmd setcursor 761 334‬", 0, True
    .Run "nircmd sendmouse left click", 0, True
End With
WScript.Sleep 500

If var1 = "1" Then
WshShell.SendKeys "a"
ElseIf var1 = "2" Then
WshShell.SendKeys "b"
ElseIf var1 = "3" Then
WshShell.SendKeys "c"
ElseIf var1 = "4" Then
WshShell.SendKeys "d"
ElseIf var1 = "5" Then
WshShell.SendKeys "e"
ElseIf var1 = "6" Then
WshShell.SendKeys "f"
ElseIf var1 = "7" Then
WshShell.SendKeys "g"
ElseIf var1 = "8" Then
WshShell.SendKeys "h"
ElseIf var1 = "9" Then
WshShell.SendKeys "i"
ElseIf var1 = "10" Then
WshShell.SendKeys "j"
ElseIf var1 = "11" Then
WshShell.SendKeys "k"
ElseIf var1 = "12" Then
WshShell.SendKeys "l"
ElseIf var1 = "13" Then
WshShell.SendKeys "m"
ElseIf var1 = "14" Then
WshShell.SendKeys "n"
ElseIf var1 = "15" Then
WshShell.SendKeys "o"
ElseIf var1 = "16" Then
WshShell.SendKeys "p"
ElseIf var1 = "17" Then
WshShell.SendKeys "q"
ElseIf var1 = "18" Then
WshShell.SendKeys "r"
ElseIf var1 = "19" Then
WshShell.SendKeys "s"
ElseIf var1 = "20" Then
WshShell.SendKeys "t"
ElseIf var1 = "21" Then
WshShell.SendKeys "u"
ElseIf var1 = "22" Then
WshShell.SendKeys "v"
ElseIf var1 = "23" Then
WshShell.SendKeys "w"
ElseIf var1 = "24" Then
WshShell.SendKeys "x"
ElseIf var1 = "25" Then
WshShell.SendKeys "y"
ElseIf var1 = "26" Then
WshShell.SendKeys "z"
End If

If var2 = "1" Then
 WshShell.SendKeys "a"
 ElseIf var2 = "2" Then
 WshShell.SendKeys "b"
 ElseIf var2 = "3" Then
 WshShell.SendKeys "c"
 ElseIf var2 = "4" Then
 WshShell.SendKeys "d"
 ElseIf var2 = "5" Then
 WshShell.SendKeys "e"
 ElseIf var2 = "6" Then
 WshShell.SendKeys "f"
 ElseIf var2 = "7" Then
 WshShell.SendKeys "g"
 ElseIf var2 = "8" Then
 WshShell.SendKeys "h"
 ElseIf var2 = "9" Then
 WshShell.SendKeys "i"
 ElseIf var2 = "10" Then
 WshShell.SendKeys "j"
 ElseIf var2 = "11" Then
 WshShell.SendKeys "k"
 ElseIf var2 = "12" Then
 WshShell.SendKeys "l"
 ElseIf var2 = "13" Then
 WshShell.SendKeys "m"
 ElseIf var2 = "14" Then
 WshShell.SendKeys "n"
 ElseIf var2 = "15" Then
 WshShell.SendKeys "o"
 ElseIf var2 = "16" Then
 WshShell.SendKeys "p"
 ElseIf var2 = "17" Then
 WshShell.SendKeys "q"
 ElseIf var2 = "18" Then
 WshShell.SendKeys "r"
 ElseIf var2 = "19" Then
 WshShell.SendKeys "s"
 ElseIf var2 = "20" Then
 WshShell.SendKeys "t"
 ElseIf var2 = "21" Then
 WshShell.SendKeys "u"
 ElseIf var2 = "22" Then
 WshShell.SendKeys "v"
 ElseIf var2 = "23" Then
 WshShell.SendKeys "w"
 ElseIf var2 = "24" Then
 WshShell.SendKeys "x"
 ElseIf var2 = "25" Then
 WshShell.SendKeys "y"
 ElseIf var2 = "26" Then
 WshShell.SendKeys "z"
 var3 = var3 + "1"
 End If

If var3 = "1" Then
WshShell.SendKeys "a"
ElseIf var3 = "2" Then
WshShell.SendKeys "b"
ElseIf var3 = "3" Then
WshShell.SendKeys "c"
ElseIf var3 = "4" Then
WshShell.SendKeys "d"
ElseIf var3 = "5" Then
WshShell.SendKeys "e"
ElseIf var3 = "6" Then
WshShell.SendKeys "f"
ElseIf var3 = "7" Then
WshShell.SendKeys "g"
ElseIf var3 = "8" Then
WshShell.SendKeys "h"
ElseIf var3 = "9" Then
WshShell.SendKeys "i"
ElseIf var3 = "10" Then
WshShell.SendKeys "j"
ElseIf var3 = "11" Then
WshShell.SendKeys "k"
ElseIf var3 = "12" Then
WshShell.SendKeys "l"
ElseIf var3 = "13" Then
WshShell.SendKeys "m"
ElseIf var3 = "14" Then
WshShell.SendKeys "n"
ElseIf var3 = "15" Then
WshShell.SendKeys "o"
ElseIf var3 = "16" Then
WshShell.SendKeys "p"
ElseIf var3 = "17" Then
WshShell.SendKeys "q"
ElseIf var3 = "18" Then
WshShell.SendKeys "r"
ElseIf var3 = "19" Then
WshShell.SendKeys "s"
ElseIf var3 = "20" Then
WshShell.SendKeys "t"
ElseIf var3 = "21" Then
WshShell.SendKeys "u"
ElseIf var3 = "22" Then
WshShell.SendKeys "v"
ElseIf var3 = "23" Then
WshShell.SendKeys "w"
ElseIf var3 = "24" Then
WshShell.SendKeys "x"
ElseIf var3 = "25" Then
WshShell.SendKeys "y"
ElseIf var3 = "26" Then
WshShell.SendKeys "z"
End If

If var4 = "1" Then
WshShell.SendKeys "a"
ElseIf var4 = "2" Then
WshShell.SendKeys "b"
ElseIf var4 = "3" Then
WshShell.SendKeys "c"
ElseIf var4 = "4" Then
WshShell.SendKeys "d"
ElseIf var4 = "5" Then
WshShell.SendKeys "e"
ElseIf var4 = "6" Then
WshShell.SendKeys "f"
ElseIf var4 = "7" Then
WshShell.SendKeys "g"
ElseIf var4 = "8" Then
WshShell.SendKeys "h"
ElseIf var4 = "9" Then
WshShell.SendKeys "i"
ElseIf var4 = "10" Then
WshShell.SendKeys "j"
ElseIf var4 = "11" Then
WshShell.SendKeys "k"
ElseIf var4 = "12" Then
WshShell.SendKeys "l"
ElseIf var4 = "13" Then
WshShell.SendKeys "m"
ElseIf var4 = "14" Then
WshShell.SendKeys "n"
ElseIf var4 = "15" Then
WshShell.SendKeys "o"
ElseIf var4 = "16" Then
WshShell.SendKeys "p"
ElseIf var4 = "17" Then
WshShell.SendKeys "q"
ElseIf var4 = "18" Then
WshShell.SendKeys "r"
ElseIf var4 = "19" Then
WshShell.SendKeys "s"
ElseIf var4 = "20" Then
WshShell.SendKeys "t"
ElseIf var4 = "21" Then
WshShell.SendKeys "u"
ElseIf var4 = "22" Then
WshShell.SendKeys "v"
ElseIf var4 = "23" Then
WshShell.SendKeys "w"
ElseIf var4 = "24" Then
WshShell.SendKeys "x"
ElseIf var4 = "25" Then
WshShell.SendKeys "y"
ElseIf var4 = "26" Then
WshShell.SendKeys "z"
End If

WshShell.SendKeys "{ENTER}"
var1 = var1 + 1
loop
